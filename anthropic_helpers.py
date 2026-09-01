"""Small helpers for Anthropic SDK responses.

The Foundry proxy (and any request that enables extended thinking) can return
a `ThinkingBlock` as `resp.content[0]`, ahead of the `TextBlock`. Blind
indexing (`resp.content[0].text`) then raises
`AttributeError: 'ThinkingBlock' object has no attribute 'text'`. Every call
site should route through `first_text` instead.
"""
from __future__ import annotations


def first_text(resp) -> str:
    """Return the text of the first text-bearing content block.

    Skips ThinkingBlock, ToolUseBlock, and any other non-text block that the
    Anthropic SDK may put in front of the model's textual answer. Raises
    ValueError if no text block is present (so callers can distinguish a
    genuinely empty response from an extraction bug).
    """
    for block in getattr(resp, "content", None) or []:
        # SDK TextBlock has type == "text"; be permissive in case the client
        # returns a dict-shaped payload from a proxy.
        btype = getattr(block, "type", None) or (
            block.get("type") if isinstance(block, dict) else None
        )
        if btype == "text":
            text = getattr(block, "text", None)
            if text is None and isinstance(block, dict):
                text = block.get("text")
            if text is not None:
                return text
        # Fallback: some proxies emit text blocks without an explicit type
        # field but still expose `.text`. Accept these too, but only when no
        # `type` was present at all, so we don't misread a ThinkingBlock that
        # happens to carry a `.text`-adjacent attribute in a future SDK.
        if btype is None and hasattr(block, "text"):
            text = getattr(block, "text")
            if text:
                return text
    raise ValueError("Anthropic response contained no text block")
