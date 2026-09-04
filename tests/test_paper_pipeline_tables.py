"""Unit tests for PaperPipeline table extraction/rendering.

Regression coverage for a real corruption bug: pdfplumber preserves the
literal line breaks of a cell that wraps across multiple lines in a narrow
PDF column (e.g. "Onvanser\ntib"). Left unsanitized, that raw newline broke
the `|`-delimited markdown row into two lines and shifted every column
boundary for the rest of the table.

No PDF fixture or API credentials needed — these exercise the pure cell
sanitization / markdown-rendering helpers directly.

Run: python -m pytest tests/test_paper_pipeline_tables.py -v
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from paper_pipeline import PaperPipeline  # noqa: E402


def _pipeline(tmp_path: Path) -> PaperPipeline:
    return PaperPipeline(input_folder=str(tmp_path), output_dir=str(tmp_path / "out"))


class TestCleanTableCell:
    def test_collapses_embedded_newline_from_wrapped_cell(self):
        assert PaperPipeline._clean_table_cell("Onvanser\ntib") == "Onvanser tib"

    def test_collapses_multi_line_wrap(self):
        assert PaperPipeline._clean_table_cell("Selective\nSmall-\nMolecule\nPLK1\nInhibitor") == \
            "Selective Small- Molecule PLK1 Inhibitor"

    def test_collapses_tabs_and_repeated_whitespace(self):
        assert PaperPipeline._clean_table_cell("Phase\t\t2  ") == "Phase 2"

    def test_none_and_empty_become_empty_string(self):
        assert PaperPipeline._clean_table_cell(None) == ''
        assert PaperPipeline._clean_table_cell('') == ''

    def test_plain_cell_unaffected(self):
        assert PaperPipeline._clean_table_cell("Cardiff Oncology") == "Cardiff Oncology"


class TestEscapeTableCell:
    def test_escapes_literal_pipe(self):
        pipeline = _pipeline(Path.cwd())
        assert pipeline._escape_table_cell("FOLFIRI | bevacizumab") == r"FOLFIRI \| bevacizumab"

    def test_defensively_cleans_unsanitized_input(self):
        # Guards against a future caller feeding _format_table_markdown a
        # table dict that bypassed extract_tables()'s own cleaning.
        pipeline = _pipeline(Path.cwd())
        assert pipeline._escape_table_cell("Onvanser\ntib") == "Onvanser tib"


class TestFormatTableMarkdown:
    def test_wrapped_cell_does_not_corrupt_row_structure(self):
        """The bug: an unsanitized wrapped cell used to split one row into
        two lines, shifting every '|' column boundary after it."""
        pipeline = _pipeline(Path.cwd())
        table = {
            'table_number': 1,
            'caption': 'PLK1 competitors',
            'header': ['Candidate', 'Modality', 'Sponsor'],
            'rows': [
                ['Onvanser\ntib', 'Selective\nSmall-\nMolecule\nPLK1\nInhibitor', 'Cardiff\nOncology'],
                ['Volasertib', 'Pan-PLK inhibitor', 'Boehringer Ingelheim'],
            ],
        }
        md = pipeline._format_table_markdown(table)
        lines = [l for l in md.split('\n') if l.strip()]

        # Header line, separator line, then exactly one line per data row.
        table_lines = [l for l in lines if l.startswith('|')]
        assert len(table_lines) == 4  # header + separator + 2 rows

        header_line, sep_line, row1, row2 = table_lines
        assert header_line == '| Candidate | Modality | Sponsor |'
        assert row1 == '| Onvanser tib | Selective Small- Molecule PLK1 Inhibitor | Cardiff Oncology |'
        assert row2 == '| Volasertib | Pan-PLK inhibitor | Boehringer Ingelheim |'

        # Every table row has the same column count as the header.
        expected_pipes = header_line.count('|')
        for line in table_lines:
            assert line.count('|') == expected_pipes

    def test_literal_pipe_in_cell_does_not_add_a_column(self):
        pipeline = _pipeline(Path.cwd())
        table = {
            'table_number': 2,
            'caption': '',
            'header': ['Regimen', 'Notes'],
            'rows': [['FOLFIRI | bevacizumab', 'ORR 72.2%']],
        }
        md = pipeline._format_table_markdown(table)
        row_line = [l for l in md.split('\n') if l.startswith('| FOLFIRI')][0]
        # The escaped `\|` must not be readable as a real column boundary:
        # exactly 2 columns (Regimen, Notes), not 3.
        assert row_line == r'| FOLFIRI \| bevacizumab | ORR 72.2% |'
