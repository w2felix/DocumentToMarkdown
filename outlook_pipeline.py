"""
Outlook Email Processing Pipeline
Ingests emails from an Outlook folder, groups by conversation thread,
and produces structured markdown with processed attachments.

Uses win32com to access the local Outlook desktop app (no OAuth needed).
Tracks processed emails by EntryID for incremental re-runs.
"""

import os
import re
import json
import html
import argparse
import tempfile
import logging
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from datetime import datetime
from collections import defaultdict

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

from pipeline_security import validate_output_path, sanitize_filename


class OutlookPipeline:
    """Pipeline for processing Outlook emails into structured markdown threads."""

    OL_MAIL_ITEM = 43
    OL_FOLDER_INBOX = 6
    MAX_ATTACHMENT_SIZE_MB = 50

    ATTACHMENT_PIPELINES = {
        '.pdf': 'paper',
        '.pptx': 'presentation',
        '.docx': 'docx',
        '.xlsx': 'xlsx',
        '.xls': 'xlsx',
    }

    IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.svg'}

    SUBJECT_PREFIX_RE = re.compile(
        r'^(?:(?:Re|Fwd|FW|RE|AW|WG|Antwort|Weitergeleitet)\s*:\s*)+',
        re.IGNORECASE
    )

    SIGNATURE_DELIMITERS = [
        '\n-- \n',       # standard sig delimiter
        '\n___\n',       # Outlook-style
        '\n---\n',       # markdown-style
        '\n_____',       # long underscore
    ]

    SIGNATURE_START_RE = re.compile(
        r'\n(?:'
        r'(?:Mit freundlichen Gr[uü][sß]en|'
        r'Best regards|Kind regards|Regards|Thanks|Cheers|Sincerely|'
        r'Viele Gr[uü][sß]e|Liebe Gr[uü][sß]e|MfG|VG|BR|'
        r'Sent from my iPhone|Sent from my iPad|'
        r'Get Outlook for)'
        r')[\s,]*\n',
        re.IGNORECASE
    )

    INLINE_IMAGE_MAX_KB = 15

    def __init__(self, folder_path: str, output_dir: str, verbose: bool = False,
                 no_attachments: bool = False, reprocess: bool = False, limit: int = 0):
        self.folder_path = folder_path
        self.output_dir = Path(output_dir).resolve()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.verbose = verbose
        self.no_attachments = no_attachments
        self.reprocess = reprocess
        self.limit = limit

        self._staging_dir = self.output_dir / ".staging"
        self._staging_dir.mkdir(parents=True, exist_ok=True)

        self.state_path = self.output_dir / "processed_state.json"
        self.processing_log_path = self.output_dir / "processing_log.tsv"
        self._state = self._load_state()
        self._init_processing_log()

    # ─── COM Connection ─────────────────────────────────────────────────

    def _connect_outlook(self):
        """Connect to running Outlook instance via COM."""
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except ImportError:
            pass

        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            return namespace
        except Exception as e:
            raise RuntimeError(
                f"Cannot connect to Outlook. Is Outlook running?\n"
                f"Error: {e}\n"
                f"Ensure Outlook is open and pywin32 is installed."
            )

    def _navigate_to_folder(self, namespace, folder_path: str):
        """Navigate to specified folder within Outlook namespace."""
        parts = re.split(r'[/\\]', folder_path.strip().strip('/\\'))
        parts = [p for p in parts if p]

        if not parts:
            raise ValueError("Folder path is empty")

        first = parts[0].lower()
        if first == 'inbox' or first == 'posteingang':
            folder = namespace.GetDefaultFolder(self.OL_FOLDER_INBOX)
            parts = parts[1:]
        else:
            try:
                folder = namespace.GetDefaultFolder(self.OL_FOLDER_INBOX)
                parent = folder.Parent
                folder = parent.Folders[parts[0]]
                parts = parts[1:]
            except Exception:
                folder = namespace.GetDefaultFolder(self.OL_FOLDER_INBOX)
                # If first part is something like "Inbox" in another language, try subfolders
                pass

        for part in parts:
            try:
                folder = folder.Folders[part]
            except Exception:
                available = [folder.Folders.Item(i + 1).Name
                             for i in range(folder.Folders.Count)]
                raise ValueError(
                    f"Subfolder '{part}' not found. "
                    f"Available folders: {available}"
                )

        return folder

    # ─── State Management ───────────────────────────────────────────────

    def _load_state(self) -> Dict:
        """Load processed email state from JSON file."""
        if self.state_path.exists():
            try:
                data = json.loads(self.state_path.read_text(encoding='utf-8'))
                if isinstance(data, dict) and 'entries' in data:
                    return data
            except (json.JSONDecodeError, ValueError) as e:
                backup = self.state_path.with_suffix('.json.bak')
                logger.warning(f"Corrupt state file, backing up to {backup}: {e}")
                self.state_path.rename(backup)

        return {"version": 1, "entries": {}}

    def _save_state(self):
        """Atomically write state file."""
        tmp_path = self.state_path.with_suffix('.json.tmp')
        tmp_path.write_text(json.dumps(self._state, indent=2, default=str), encoding='utf-8')
        tmp_path.replace(self.state_path)

    def _is_processed(self, entry_id: str) -> bool:
        return entry_id in self._state["entries"]

    def _mark_processed(self, entry_id: str, subject: str, date: str, thread_id: str):
        self._state["entries"][entry_id] = {
            "subject": subject,
            "date": date,
            "thread_id": thread_id,
            "processed_timestamp": datetime.now().isoformat(),
        }

    # ─── Email Iteration & Threading ────────────────────────────────────

    def _snapshot_mail(self, mail_item) -> Dict:
        """Snapshot all COM properties into a plain dict to avoid stale references."""
        entry_id = ''
        try:
            entry_id = mail_item.EntryID
        except Exception:
            entry_id = f"unknown_{id(mail_item)}"

        subject = ''
        try:
            subject = mail_item.Subject or ''
        except Exception:
            pass

        sender_name = ''
        sender_email = ''
        try:
            sender_name = mail_item.SenderName or ''
        except Exception:
            pass
        try:
            sender_email = mail_item.SenderEmailAddress or ''
        except Exception:
            pass

        received = None
        try:
            received = mail_item.ReceivedTime
        except Exception:
            pass

        body = ''
        try:
            body = mail_item.Body or ''
        except Exception:
            pass
        html_body = ''
        if not body.strip():
            try:
                html_body = mail_item.HTMLBody or ''
            except Exception:
                pass

        thread_id = ''
        try:
            thread_id = mail_item.ConversationID or ''
        except Exception:
            pass

        to_list = []
        cc_list = []
        recipients_names = []
        try:
            for j in range(1, mail_item.Recipients.Count + 1):
                recip = mail_item.Recipients.Item(j)
                name = getattr(recip, 'Name', '') or getattr(recip, 'Address', '') or ''
                rtype = getattr(recip, 'Type', 1)
                if rtype == 1:
                    to_list.append(name)
                elif rtype == 2:
                    cc_list.append(name)
                if name:
                    recipients_names.append(name)
        except Exception:
            pass

        # Save attachments eagerly to staging dir while COM is alive
        saved_attachments = []
        try:
            att_count = mail_item.Attachments.Count
            for i in range(1, att_count + 1):
                try:
                    att = mail_item.Attachments.Item(i)
                    filename = att.FileName
                    if not filename:
                        continue
                    filename = sanitize_filename(filename)
                    save_path = self._staging_dir / f"{entry_id[:16]}_{i}_{filename}"
                    att.SaveAsFile(str(save_path))
                    saved_attachments.append({
                        'original_name': filename,
                        'staged_path': save_path,
                    })
                except Exception as e:
                    logger.debug(f"Could not save attachment {i}: {e}")
        except Exception:
            pass

        return {
            'entry_id': entry_id,
            'subject': subject,
            'sender_name': sender_name,
            'sender_email': sender_email,
            'received': received,
            'body': body.strip(),
            'html_body': html_body,
            'thread_id': thread_id,
            'to': to_list,
            'cc': cc_list,
            'recipients_names': recipients_names,
            'saved_attachments': saved_attachments,
        }

    def _get_mail_items(self, folder) -> List[Dict]:
        """Retrieve and snapshot mail items from folder."""
        items = folder.Items
        items.Sort("[ReceivedTime]", False)  # oldest first

        mail_items = []
        count = items.Count
        for i in range(1, count + 1):
            try:
                item = items.Item(i)
                if item.Class == self.OL_MAIL_ITEM:
                    snapshot = self._snapshot_mail(item)
                    mail_items.append(snapshot)
                    if self.limit and len(mail_items) >= self.limit:
                        break
            except Exception as e:
                logger.debug(f"Skipping item {i}: {e}")
                continue

        logger.info(f"Found {len(mail_items)} emails in folder '{self.folder_path}'")
        return mail_items

    def _get_thread_id_from_snapshot(self, snap: Dict) -> str:
        """Get thread identifier from a snapshot dict."""
        if snap['thread_id']:
            return snap['thread_id']
        subject = snap['subject'] or 'no_subject'
        return f"subj_{self._normalize_subject(subject)}"

    def _normalize_subject(self, subject: str) -> str:
        """Strip reply/forward prefixes and normalize for grouping."""
        cleaned = self.SUBJECT_PREFIX_RE.sub('', subject)
        return cleaned.strip().lower()

    def _group_into_threads(self, mail_items: List[Dict]) -> Dict[str, List[Dict]]:
        """Group snapshot dicts by conversation thread."""
        threads = defaultdict(list)
        for snap in mail_items:
            thread_id = self._get_thread_id_from_snapshot(snap)
            threads[thread_id].append(snap)

        for thread_id in threads:
            threads[thread_id].sort(
                key=lambda m: m['received'] or datetime.min
            )

        logger.info(f"Grouped into {len(threads)} conversation threads")
        return dict(threads)

    MEETING_INVITE_RE = re.compile(
        r'(?:'
        r'Microsoft Teams|Join the meeting|Meeting ID:|'
        r'Dial in by phone|Phone conference ID|'
        r'Join Zoom Meeting|zoom\.us/j/|'
        r'Google Meet|meet\.google\.com|'
        r'Join on a video conferencing device|'
        r'For organizers: Meeting options|'
        r'You\'re invited to|RSVP|'
        r'When:.*\d{4}.*Where:'
        r')',
        re.IGNORECASE
    )

    MANDATORY_INFO_RE = re.compile(
        r'(?:Mandatory information|Pflichtangaben|mandatories\.merckgroup|'
        r'Privacy Note:.*personal data|'
        r'This email.*confidential|'
        r'CONFIDENTIALITY NOTICE).*',
        re.IGNORECASE | re.DOTALL
    )

    # ─── Signature Detection ───────────────────────────────────────────

    def _split_signature(self, body: str) -> Tuple[str, Optional[str]]:
        """Split email body into (content, signature). Returns (body, None) if no signature found."""
        if not body:
            return body, None

        # First strip meeting invites from body
        body = self._strip_meeting_invite(body)

        # Try explicit delimiters first
        for delim in self.SIGNATURE_DELIMITERS:
            idx = body.find(delim)
            if idx > 0:
                content = body[:idx].rstrip()
                sig = body[idx + len(delim):].strip()
                if sig and len(sig) > 10:
                    sig = self._clean_signature(sig)
                    if sig:
                        return content, sig
                    return content, None

        # Try common sign-off phrases
        match = self.SIGNATURE_START_RE.search(body)
        if match and match.start() > len(body) * 0.3:
            content = body[:match.start()].rstrip()
            sig = body[match.start():].strip()
            if sig and len(sig) > 10:
                sig = self._clean_signature(sig)
                if sig:
                    return content, sig
                return content, None

        return body, None

    def _strip_meeting_invite(self, text: str) -> str:
        """Remove Teams/Zoom/Google meeting invite blocks from text."""
        # Remove block starting from meeting markers to end or next delimiter
        lines = text.split('\n')
        result = []
        skip = False
        for line in lines:
            if self.MEETING_INVITE_RE.search(line):
                skip = True
                continue
            if skip and (line.strip().startswith('___') or line.strip().startswith('---')):
                skip = False
                continue
            if not skip:
                result.append(line)
        return '\n'.join(result)

    def _clean_signature(self, sig: str) -> Optional[str]:
        """Clean up a signature: remove meeting invites, mandatory notices, format nicely."""
        # If the whole thing is a meeting invite, discard
        if self.MEETING_INVITE_RE.search(sig[:200]):
            return None

        # Strip mandatory/privacy notices
        sig = self.MANDATORY_INFO_RE.sub('', sig)

        # Remove urldefense wrappers: [text](url) → text
        sig = re.sub(r'\[([^\]]+)\]\(https?://urldefense[^)]+\)', r'\1', sig)
        # Remove raw urldefense links
        sig = re.sub(r'<https?://urldefense[^>]+>', '', sig)
        # Clean markdown links with display text: [text](url) → text
        sig = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', sig)
        # Remove <mailto:...> and <http...> angle-bracket links
        sig = re.sub(r'<(?:mailto:|https?://)[^>]+>', '', sig)
        # Remove standalone http URLs
        sig = re.sub(r'https?://\S+', '', sig)

        # Remove excessive underscores (separator lines)
        sig = re.sub(r'_{3,}', '', sig)
        # Collapse whitespace
        sig = re.sub(r'\n{3,}', '\n', sig)
        sig = re.sub(r'[ \t]+', ' ', sig)
        # Remove leading/trailing blank lines
        sig = '\n'.join(line.strip() for line in sig.strip().split('\n'))
        sig = re.sub(r'\n{3,}', '\n\n', sig)

        # If too short after cleaning, discard
        if not sig or len(sig.strip()) < 15:
            return None

        return sig.strip()

    def _identify_signature_owner(self, sig: str, fallback_name: str) -> str:
        """Extract the person's name from the signature content itself."""
        lines = [l.strip() for l in sig.split('\n') if l.strip()]
        # Skip sign-off line (Best regards, Thanks, etc.)
        start = 0
        for i, line in enumerate(lines):
            if re.match(
                r'^(?:Best regards|Kind regards|Regards|Thanks|Cheers|Sincerely|'
                r'Mit freundlichen Gr[uü][sß]en|Viele Gr[uü][sß]e|MfG|VG|BR)',
                line, re.IGNORECASE
            ):
                start = i + 1
                continue
            # Also skip single first-name lines right after greeting (e.g. "Carsten")
            if i == start and len(line.split()) == 1 and line[0].isupper() and len(line) < 20:
                start = i + 1
                continue
            break

        # The name is typically the first non-empty line after the sign-off
        for line in lines[start:]:
            if (len(line) < 50 and
                not any(x in line.lower() for x in ['|', 'phone:', 'e-mail:', 'www.', 'http', '@']) and
                not line.startswith('+') and
                not re.match(r'^[\d\s\-\(\)]+$', line)):
                return line

        return fallback_name

    def _strip_signoff_from_signature(self, sig: str) -> str:
        """Remove the greeting/sign-off lines from the beginning of a signature."""
        lines = sig.split('\n')
        start = 0
        found_signoff = False
        for i, line in enumerate(lines):
            stripped = line.strip()
            if not stripped:
                if found_signoff:
                    start = i + 1
                continue
            if re.match(
                r'^(?:Best regards|Kind regards|Regards|Thanks|Cheers|Sincerely|'
                r'Mit freundlichen Gr[uü][sß]en|Viele Gr[uü][sß]e|MfG|VG|BR)[,]?\s*$',
                stripped, re.IGNORECASE
            ):
                start = i + 1
                found_signoff = True
                continue
            # Single first-name line shortly after greeting
            if found_signoff and len(stripped.split()) == 1 and stripped[0].isupper() and len(stripped) < 20:
                start = i + 1
                continue
            break
        return '\n'.join(lines[start:]).strip()

    # ─── Body Extraction ────────────────────────────────────────────────

    def _extract_body(self, snap: Dict) -> str:
        """Extract email body from snapshot as plain text or converted markdown."""
        if snap['body']:
            text = snap['body']
        elif snap['html_body'] and snap['html_body'].strip():
            text = self._html_to_text(snap['html_body'])
        else:
            return "[Empty email body]"

        # Collapse excessive blank lines
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = re.sub(r'(\n[ \t]*){3,}', '\n\n', text)
        return text.strip()

    def _html_to_text(self, html_content: str) -> str:
        """Lightweight HTML to markdown-ish text conversion."""
        text = html_content
        # Remove style and script blocks
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
        # Convert links
        text = re.sub(r'<a\s+[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
                      r'[\2](\1)', text, flags=re.IGNORECASE | re.DOTALL)
        # Convert bold/strong
        text = re.sub(r'<(?:b|strong)[^>]*>(.*?)</(?:b|strong)>', r'**\1**',
                      text, flags=re.IGNORECASE | re.DOTALL)
        # Convert italic/em
        text = re.sub(r'<(?:i|em)[^>]*>(.*?)</(?:i|em)>', r'*\1*',
                      text, flags=re.IGNORECASE | re.DOTALL)
        # Convert list items
        text = re.sub(r'<li[^>]*>(.*?)</li>', r'- \1\n', text, flags=re.IGNORECASE | re.DOTALL)
        # Convert line breaks and paragraphs
        text = re.sub(r'<br\s*/?\s*>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</p>', '\n\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</?(?:div|tr)[^>]*>', '\n', text, flags=re.IGNORECASE)
        # Strip remaining tags
        text = re.sub(r'<[^>]+>', '', text)
        # Decode HTML entities
        text = html.unescape(text)
        # Collapse excessive whitespace
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = re.sub(r'[ \t]+', ' ', text)
        return text.strip()

    # ─── Attachment Handling ────────────────────────────────────────────

    def _extract_attachments(self, snap: Dict, thread_dir: Path) -> List[Dict]:
        """Move pre-saved attachments from staging to the thread directory."""
        attachments_info = []

        for staged in snap['saved_attachments']:
            staged_path = staged['staged_path']
            filename = staged['original_name']

            if not staged_path.exists():
                continue

            # Skip tiny inline images (signature logos)
            ext = Path(filename).suffix.lower()
            size_kb = staged_path.stat().st_size / 1024
            if ext in self.IMAGE_EXTENSIONS and size_kb < self.INLINE_IMAGE_MAX_KB:
                staged_path.unlink(missing_ok=True)
                continue

            dest_path = thread_dir / filename
            if dest_path.exists():
                stem = dest_path.stem
                suffix = dest_path.suffix
                counter = 1
                while dest_path.exists():
                    dest_path = thread_dir / f"{stem}_{counter}{suffix}"
                    counter += 1

            try:
                staged_path.rename(dest_path)
            except OSError:
                import shutil
                shutil.move(str(staged_path), str(dest_path))

            size_mb = dest_path.stat().st_size / (1024 * 1024)
            if size_mb > self.MAX_ATTACHMENT_SIZE_MB:
                logger.warning(f"Attachment too large ({size_mb:.1f}MB): {filename}")
                attachments_info.append({
                    'filename': dest_path.name,
                    'path': dest_path,
                    'size_mb': size_mb,
                    'type': 'too_large',
                    'processed': False,
                    'output': None,
                })
                continue

            att_type = self.ATTACHMENT_PIPELINES.get(ext, 'raw')
            if ext in self.IMAGE_EXTENSIONS:
                att_type = 'image'

            attachments_info.append({
                'filename': dest_path.name,
                'path': dest_path,
                'size_mb': size_mb,
                'type': att_type,
                'processed': False,
                'output': None,
            })

        return attachments_info

    def _route_attachment(self, att_info: Dict, thread_dir: Path) -> Optional[str]:
        """Route an attachment to the appropriate pipeline for processing."""
        file_path = att_info['path']
        att_type = att_info['type']
        output_name = f"attachment_{file_path.stem}.md"
        output_path = thread_dir / output_name

        try:
            if att_type == 'paper':
                return self._process_pdf_attachment(file_path, thread_dir, output_name)
            elif att_type == 'presentation':
                return self._process_pptx_attachment(file_path, thread_dir, output_name)
            elif att_type == 'docx':
                return self._process_docx_attachment(file_path, output_path)
            elif att_type == 'xlsx':
                return self._process_xlsx_attachment(file_path, output_path)
            elif att_type == 'image':
                return None  # Keep raw, referenced in thread.md
            else:
                return None  # Keep raw
        except Exception as e:
            logger.warning(f"Failed to process attachment {file_path.name}: {e}")
            return None

    def _classify_pdf(self, file_path: Path) -> str:
        """Classify a PDF as 'paper' (multi-page text), 'poster' (single-page visual), or 'visual'."""
        try:
            import fitz
            doc = fitz.open(str(file_path))
            page_count = len(doc)
            if page_count == 0:
                doc.close()
                return 'paper'

            # Check text density of first page
            first_page = doc[0]
            text = first_page.get_text()
            text_chars = len(text.strip())
            images = first_page.get_images()

            doc.close()

            # Single page with few text chars → visual/poster (needs vision AI)
            if page_count <= 2 and text_chars < 500 and len(images) > 0:
                return 'visual'
            # Single page with moderate text → could be poster
            if page_count == 1 and text_chars < 2000:
                return 'visual'
            # Multi-page with decent text → standard paper
            return 'paper'
        except Exception:
            return 'paper'

    def _process_pdf_attachment(self, file_path: Path, thread_dir: Path,
                                output_name: str) -> Optional[str]:
        """Process a PDF attachment — route based on document type."""
        pdf_type = self._classify_pdf(file_path)

        if pdf_type == 'visual':
            return self._process_pdf_with_vision(file_path, thread_dir, output_name)
        else:
            return self._process_pdf_as_paper(file_path, thread_dir, output_name)

    def _process_pdf_with_vision(self, file_path: Path, thread_dir: Path,
                                 output_name: str) -> Optional[str]:
        """Process a visual/poster-style PDF using direct vision AI extraction."""
        try:
            import fitz
            import base64
            from auth import get_anthropic_client

            doc = fitz.open(str(file_path))
            images_b64 = []
            for page in doc:
                pix = page.get_pixmap(dpi=150)
                img_bytes = pix.tobytes("png")
                images_b64.append(base64.standard_b64encode(img_bytes).decode())
            doc.close()

            if not images_b64:
                return None

            client = get_anthropic_client()
            content = []
            for img_b64 in images_b64:
                content.append({
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/png", "data": img_b64}
                })
            content.append({
                "type": "text",
                "text": (
                    "Extract ALL text and information from this document image. "
                    "This appears to be a poster, infographic, or visual document. "
                    "Structure your response as markdown with appropriate headings. "
                    "Include: title, authors/presenters, all text content, key data points, "
                    "and descriptions of any figures, charts, or diagrams."
                )
            })

            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4000,
                messages=[{"role": "user", "content": content}]
            )

            extracted = response.content[0].text
            output_path = thread_dir / output_name

            md_content = f"---\ntitle: \"{self._escape_yaml(file_path.stem)}\"\n"
            md_content += f"source_file: \"{file_path.name}\"\n"
            md_content += f"extraction_method: vision_ai\n"
            md_content += f"processing_date: \"{datetime.now().strftime('%Y-%m-%d')}\"\n"
            md_content += "---\n\n"
            md_content += extracted

            output_path.write_text(md_content, encoding='utf-8')
            return output_name

        except Exception as e:
            logger.warning(f"Vision AI extraction failed for {file_path.name}: {e}")
            return self._process_pdf_as_paper(file_path, thread_dir, output_name)

    def _process_pdf_as_paper(self, file_path: Path, thread_dir: Path,
                              output_name: str) -> Optional[str]:
        """Process a multi-page PDF through the paper pipeline."""
        try:
            from paper_pipeline import PaperPipeline
            pipeline = PaperPipeline(
                input_folder=str(thread_dir),
                output_dir=str(thread_dir),
            )
            pipeline.process_single_paper(file_path, skip_existing=False)
            for f in thread_dir.glob('*.md'):
                if f.stem != 'thread' and f.stem != 'signatures' and file_path.stem.lower() in f.stem.lower():
                    if f.name != output_name:
                        f.rename(thread_dir / output_name)
                    return output_name
            return None
        except ImportError:
            logger.warning("paper_pipeline not available for PDF processing")
            return None

    def _process_pptx_attachment(self, file_path: Path, thread_dir: Path,
                                 output_name: str) -> Optional[str]:
        """Process a PPTX attachment through the presentation pipeline."""
        try:
            from presentation_pipeline import PresentationPipeline
            pipeline = PresentationPipeline(
                input_folder=str(thread_dir),
                output_dir=str(thread_dir),
            )
            pipeline.process_single_presentation(file_path, skip_existing=False)
            for f in thread_dir.glob('*.md'):
                if f.stem != 'thread' and file_path.stem.lower() in f.stem.lower():
                    if f.name != output_name:
                        f.rename(thread_dir / output_name)
                    return output_name
            return None
        except ImportError:
            logger.warning("presentation_pipeline not available for PPTX processing")
            return None

    def _process_docx_attachment(self, file_path: Path, output_path: Path) -> Optional[str]:
        """Extract text from DOCX and write as markdown."""
        try:
            from docx import Document
            doc = Document(str(file_path))
            lines = []
            lines.append(f"# {file_path.stem}\n")
            for para in doc.paragraphs:
                text = para.text.strip()
                if not text:
                    lines.append("")
                    continue
                style = para.style.name.lower() if para.style else ''
                if 'heading 1' in style:
                    lines.append(f"# {text}")
                elif 'heading 2' in style:
                    lines.append(f"## {text}")
                elif 'heading 3' in style:
                    lines.append(f"### {text}")
                else:
                    lines.append(text)

            output_path.write_text('\n'.join(lines), encoding='utf-8')
            return output_path.name
        except ImportError:
            logger.warning("python-docx not installed — .docx files kept raw")
            return None
        except Exception as e:
            logger.warning(f"DOCX extraction failed for {file_path.name}: {e}")
            return None

    def _process_xlsx_attachment(self, file_path: Path, output_path: Path) -> Optional[str]:
        """Convert Excel sheets to markdown tables."""
        try:
            import pandas as pd
            xls = pd.ExcelFile(file_path)
            lines = [f"# {file_path.stem}\n"]

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, nrows=500)
                if df.empty:
                    continue
                lines.append(f"## {sheet_name}\n")
                lines.append(df.to_markdown(index=False))
                lines.append("")

            output_path.write_text('\n'.join(lines), encoding='utf-8')
            return output_path.name
        except Exception as e:
            logger.warning(f"Excel extraction failed for {file_path.name}: {e}")
            return None

    # ─── Markdown Generation ────────────────────────────────────────────

    def _thread_dir_name(self, subject: str) -> str:
        """Generate a directory name for a thread."""
        normalized = self._normalize_subject(subject)
        slug = re.sub(r'[^a-z0-9]+', '_', normalized)
        slug = slug.strip('_')[:80]
        if not slug:
            slug = 'unnamed_thread'
        return f"thread_{slug}"

    def _generate_thread_markdown(self, thread_id: str, emails: List[Dict],
                                  all_attachments: Dict[str, List[Dict]],
                                  folder_path: str) -> str:
        """Generate the full thread.md content from snapshot dicts."""
        subjects = []
        participants = set()
        dates = []

        for snap in emails:
            subj = snap['subject'] or 'No Subject'
            subjects.append(subj)
            sender_display = snap['sender_name'] or snap['sender_email']
            if sender_display and not sender_display.startswith('/'):
                participants.add(sender_display)
            elif snap['sender_name']:
                participants.add(snap['sender_name'])
            for name in snap['recipients_names']:
                if name:
                    participants.add(name)
            if snap['received']:
                dates.append(snap['received'])

        title = subjects[-1] if subjects else 'Untitled Thread'
        date_range_start = min(dates).strftime('%Y-%m-%d') if dates else 'unknown'
        date_range_end = max(dates).strftime('%Y-%m-%d') if dates else 'unknown'

        all_att_list = []
        for entry_id, atts in all_attachments.items():
            for att in atts:
                all_att_list.append({
                    'name': att['filename'],
                    'processed': att['processed'],
                    'output': att.get('output'),
                })

        lines = ['---']
        lines.append(f'title: "{self._escape_yaml(title)}"')
        lines.append(f'thread_id: "{thread_id[:64]}"')
        lines.append(f'folder: "{self._escape_yaml(folder_path)}"')
        lines.append('participants:')
        for p in sorted(participants):
            lines.append(f'  - "{self._escape_yaml(p)}"')
        lines.append(f'date_range: "{date_range_start} to {date_range_end}"')
        lines.append(f'email_count: {len(emails)}')
        if all_att_list:
            lines.append('attachments:')
            for att in all_att_list:
                lines.append(f'  - name: "{self._escape_yaml(att["name"])}"')
                lines.append(f'    processed: {str(att["processed"]).lower()}')
                if att['output']:
                    lines.append(f'    output: "{att["output"]}"')
        lines.append(f'processing_date: "{datetime.now().strftime("%Y-%m-%d")}"')
        lines.append('---')
        lines.append('')

        lines.append(f'# {title}')
        lines.append('')
        lines.append(
            f'**Thread**: {len(emails)} emails | '
            f'**Period**: {date_range_start} to {date_range_end} | '
            f'**Folder**: {folder_path}'
        )
        lines.append('')

        signatures: Dict[str, str] = {}  # sender -> signature text

        for idx, snap in enumerate(emails, 1):
            lines.append('---')
            lines.append('')

            received = snap['received'].strftime('%Y-%m-%d %H:%M') if snap['received'] else 'unknown date'

            sender_name = snap['sender_name']
            sender_email = snap['sender_email']
            if sender_email and not sender_email.startswith('/'):
                sender = f"{sender_name} <{sender_email}>" if sender_name else sender_email
            else:
                sender = sender_name or 'Unknown'
            subject = snap['subject'] or 'No Subject'

            lines.append(f'## Email {idx} — {received}')
            lines.append('')
            lines.append(f'**From**: {sender}  ')
            if snap['to']:
                lines.append(f'**To**: {", ".join(snap["to"])}  ')
            if snap['cc']:
                lines.append(f'**Cc**: {", ".join(snap["cc"])}  ')
            lines.append(f'**Subject**: {subject}')
            lines.append('')

            body = self._extract_body(snap)
            content, sig = self._split_signature(body)
            lines.append(content)
            lines.append('')

            if sig:
                sig_key = self._identify_signature_owner(sig, sender_name)
                if sig_key not in signatures:
                    signatures[sig_key] = self._strip_signoff_from_signature(sig)

            entry_id = snap['entry_id']
            if entry_id in all_attachments and all_attachments[entry_id]:
                atts = all_attachments[entry_id]
                att_links = []
                for att in atts:
                    if att.get('output'):
                        att_links.append(f"[{att['filename']}]({att['output']})")
                    else:
                        att_links.append(f"{att['filename']}")
                lines.append(f'**Attachments**: {" | ".join(att_links)}')
                lines.append('')

        # Store signatures dict for separate file generation
        self._thread_signatures = signatures

        return '\n'.join(lines)

    def _escape_yaml(self, value: str) -> str:
        """Escape a string for safe YAML embedding."""
        return value.replace('\\', '\\\\').replace('"', '\\"').replace('\n', ' ')

    # ─── Processing Log ─────────────────────────────────────────────────

    def _init_processing_log(self):
        """Initialize the processing log TSV if it doesn't exist."""
        if not self.processing_log_path.exists():
            header = "TIMESTAMP\tTHREAD_SUBJECT\tEMAIL_COUNT\tATTACHMENTS\tSTATUS\n"
            self.processing_log_path.write_text(header, encoding='utf-8')

    def _log_processing(self, subject: str, email_count: int,
                        attachment_count: int, status: str):
        """Append a line to the processing log."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        safe_subject = subject.replace('\t', ' ')[:100]
        line = f"{timestamp}\t{safe_subject}\t{email_count}\t{attachment_count}\t{status}\n"
        with open(self.processing_log_path, 'a', encoding='utf-8') as f:
            f.write(line)

    # ─── Main Processing ────────────────────────────────────────────────

    def process_thread(self, thread_id: str, emails: List[Dict]) -> bool:
        """Process a single conversation thread (emails are snapshot dicts)."""
        if not self.reprocess:
            all_processed = all(
                self._is_processed(snap['entry_id']) for snap in emails
            )
            if all_processed:
                logger.debug(f"Thread already fully processed, skipping")
                return True

        subject = emails[-1]['subject'] or 'No Subject'
        dir_name = self._thread_dir_name(subject)
        thread_dir = self.output_dir / dir_name
        thread_dir.mkdir(parents=True, exist_ok=True)

        logger.info(f"Processing thread: {subject} ({len(emails)} emails)")

        # Phase 1: Extract ALL attachments first (while COM refs may still be valid)
        all_attachments: Dict[str, List[Dict]] = {}
        attachment_count = 0

        if not self.no_attachments:
            seen_attachments = set()
            for snap in emails:
                entry_id = snap['entry_id']
                atts = self._extract_attachments(snap, thread_dir)

                unique_atts = []
                for att in atts:
                    key = (att['filename'], f"{att['size_mb']:.2f}")
                    if key not in seen_attachments:
                        seen_attachments.add(key)
                        unique_atts.append(att)
                    else:
                        try:
                            att['path'].unlink()
                        except Exception:
                            pass

                all_attachments[entry_id] = unique_atts
                attachment_count += len(unique_atts)

        # Phase 2: Route attachments to pipelines (may take minutes, COM not needed)
        if not self.no_attachments:
            for entry_id, atts in all_attachments.items():
                for att in atts:
                    if att['type'] not in ('raw', 'image', 'too_large'):
                        output = self._route_attachment(att, thread_dir)
                        if output:
                            att['processed'] = True
                            att['output'] = output

        # Phase 3: Generate thread markdown (uses only snapshot data)
        self._thread_signatures = {}
        thread_md = self._generate_thread_markdown(
            thread_id, emails, all_attachments, self.folder_path
        )
        thread_path = thread_dir / "thread.md"
        thread_path.write_text(thread_md, encoding='utf-8')

        # Generate signatures.md if any signatures were detected
        if self._thread_signatures:
            sig_lines = ['---', 'title: "Participant Signatures"', '---', '', '# Signatures', '']
            for person, sig_text in self._thread_signatures.items():
                sig_lines.append(f'## {person}')
                sig_lines.append('')
                sig_lines.append(sig_text)
                sig_lines.append('')
                sig_lines.append('---')
                sig_lines.append('')
            sig_path = thread_dir / "signatures.md"
            sig_path.write_text('\n'.join(sig_lines), encoding='utf-8')

        # Mark all emails as processed
        for snap in emails:
            date_str = snap['received'].strftime('%Y-%m-%dT%H:%M:%S') if snap['received'] else ''
            self._mark_processed(snap['entry_id'], snap['subject'], date_str, thread_id)

        self._log_processing(subject, len(emails), attachment_count, 'SAVED')
        return True

    def process_all(self):
        """Main entry point: connect, snapshot all emails, then process threads."""
        logger.info(f"Connecting to Outlook...")
        namespace = self._connect_outlook()

        logger.info(f"Navigating to folder: {self.folder_path}")
        folder = self._navigate_to_folder(namespace, self.folder_path)

        logger.info(f"Retrieving emails...")
        mail_items = self._get_mail_items(folder)

        if not mail_items:
            logger.info("No emails found in folder.")
            return

        threads = self._group_into_threads(mail_items)

        success_count = 0
        fail_count = 0
        skip_count = 0

        for thread_id, emails in threads.items():
            if not self.reprocess:
                all_processed = all(
                    self._is_processed(snap['entry_id']) for snap in emails
                )
                if all_processed:
                    skip_count += 1
                    continue

            try:
                if self.process_thread(thread_id, emails):
                    success_count += 1
                else:
                    fail_count += 1
            except Exception as e:
                subject = emails[-1]['subject'] if emails else 'unknown'
                logger.error(f"Failed to process thread '{subject}': {e}")
                self._log_processing(subject or 'unknown', len(emails), 0, 'FAILED')
                fail_count += 1

        self._save_state()

        # Clean up staging directory
        try:
            import shutil
            if self._staging_dir.exists():
                shutil.rmtree(self._staging_dir, ignore_errors=True)
        except Exception:
            pass

        logger.info(f"\n{'=' * 70}")
        logger.info(f"PIPELINE COMPLETE: {success_count} processed, "
                    f"{skip_count} skipped, {fail_count} failed "
                    f"out of {len(threads)} threads")
        logger.info(f"Output: {self.output_dir}")
        logger.info(f"{'=' * 70}")


def main():
    parser = argparse.ArgumentParser(
        description='Process Outlook emails into structured markdown threads',
        epilog='Reads from local Outlook via COM — no OAuth needed. Outlook must be running.'
    )
    parser.add_argument('--folder', type=str, required=True,
                        help='Outlook folder path (e.g., "Inbox/CI Reports")')
    parser.add_argument('--output', type=str, default='output_outlook',
                        help='Output directory for markdown (default: output_outlook)')
    parser.add_argument('--no-attachments', action='store_true',
                        help='Skip attachment extraction and processing')
    parser.add_argument('--no-skip', action='store_true',
                        help='Reprocess threads even if already in state')
    parser.add_argument('--limit', type=int, default=0,
                        help='Max emails to retrieve (0=all)')
    parser.add_argument('--verbose', action='store_true',
                        help='Enable debug logging')

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    try:
        validate_output_path(args.output)
    except ValueError as e:
        logger.error(f"Output path validation failed: {e}")
        return

    pipeline = OutlookPipeline(
        folder_path=args.folder,
        output_dir=args.output,
        verbose=args.verbose,
        no_attachments=args.no_attachments,
        reprocess=args.no_skip,
        limit=args.limit,
    )

    pipeline.process_all()


if __name__ == '__main__':
    main()
