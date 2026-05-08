"""Shared CTD utilities used across all modules.

Single source of truth for section-ID regex so changes only happen here.
"""
from __future__ import annotations

import re


def section_flexible_regex(section: str) -> str:
    """Build a regex that matches a CTD section ID with flexible separators.

    Handles variations like '3.2.S.2.2', '3.2 . S . 2 . 2', '3_2_S_2_2'.
    """
    parts = [re.escape(p) for p in section.split(".") if p]
    if not parts:
        return re.escape(section)
    return r"\b" + r"\s*[\.\s_\-]\s*".join(parts) + r"\b"


def compile_section_pattern(section: str, flags: int = re.IGNORECASE) -> re.Pattern[str]:
    """Return a compiled pattern for a CTD section ID."""
    return re.compile(section_flexible_regex(section), flags)


REFER_SECTION_RE = re.compile(
    r"^refer\s+section\s+3\.2\.[sp]\.",
    re.IGNORECASE,
)
