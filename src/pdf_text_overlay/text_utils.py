"""Text normalization helpers for pdf_text_overlay."""

from __future__ import annotations

import unicodedata
from typing import Iterable, List


def is_cjk(char: str) -> bool:
    """Return True if the character belongs to the CJK ranges."""
    code = ord(char)
    return (
        0x4E00 <= code <= 0x9FFF  # CJK Unified Ideographs
        or 0x3400 <= code <= 0x4DBF  # Extension A
        or 0x20000 <= code <= 0x2CEAF  # Extension B-D
        or 0xF900 <= code <= 0xFAFF  # Compatibility Ideographs
    )


def normalize_token(
    text: str,
    keep_spaces: bool = False,
    cjk_join: bool = False,
) -> str:
    """Apply Unicode normalization, zero-width removal, and optional spacing tweaks."""
    normalized = unicodedata.normalize("NFKC", text or "")
    normalized = (
        normalized.replace("\u200b", " ")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
    )
    if not keep_spaces:
        normalized = " ".join(normalized.split())

    if cjk_join:
        normalized = _trim_cjk_spaces(normalized)
    return normalized


def _trim_cjk_spaces(text: str) -> str:
    if not text:
        return text
    chars: List[str] = []
    length = len(text)
    for idx, char in enumerate(text):
        if char.isspace():
            prev_char = text[idx - 1] if idx > 0 else ""
            next_char = text[idx + 1] if idx + 1 < length else ""
            if is_cjk(prev_char) and is_cjk(next_char):
                continue
        chars.append(char)
    return "".join(chars)


def dehyphenize(lines: Iterable[str]) -> List[str]:
    """Join hyphenated line endings according to simple rules."""
    result: List[str] = []
    buffer = ""
    for line in lines:
        segment = line.strip()
        if not segment:
            if buffer:
                result.append(buffer)
                buffer = ""
            result.append("")
            continue
        if buffer.endswith("-") and segment and segment[0].islower():
            buffer = buffer[:-1] + segment
            continue
        if buffer:
            result.append(buffer)
        buffer = segment
    if buffer:
        result.append(buffer)
    return result
