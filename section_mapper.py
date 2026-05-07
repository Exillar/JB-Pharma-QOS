from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class ResolvedSection:
    refer_section: str
    resolved_pdf: Path
    warning: str | None = None


class SectionMapper:
    CTD_PATTERN = re.compile(r"(?i)(3\.2\.[A-Z]\.\d+(?:\.\d+)*)")
    FLEX_CTD_PATTERN = re.compile(
        r"(?i)\b3[\s._-]*2[\s._-]*([SP])((?:[\s._-]*\d+)+)\b"
    )

    def __init__(self, module3_root: Path) -> None:
        self.module3_root = module3_root
        self._pdfs = list(module3_root.rglob("*.pdf"))
        self._index = self._build_index(self._pdfs)
        self._content_index: dict[str, list[Path]] = {}

    @staticmethod
    def _normalize(token: str) -> str:
        cleaned = re.sub(r"[\s._-]+", ".", token.strip().upper())
        cleaned = re.sub(r"\.+", ".", cleaned).strip(".")
        return cleaned

    @classmethod
    def _extract_flexible_sections(cls, text: str) -> set[str]:
        sections: set[str] = set()
        for alpha, tail in cls.FLEX_CTD_PATTERN.findall(text):
            nums = re.findall(r"\d+", tail)
            if nums:
                sections.add(cls._normalize(f"3.2.{alpha}." + ".".join(nums)))
        return sections

    def _build_index(self, pdfs: list[Path]) -> dict[str, list[Path]]:
        index: dict[str, list[Path]] = {}
        for pdf in pdfs:
            candidates = set()
            base = pdf.stem
            matches = self.CTD_PATTERN.findall(base)
            for m in matches:
                candidates.add(self._normalize(m))
            candidates.update(self._extract_flexible_sections(base))

            # Also infer from the leading part of filename if it starts with 3.2.
            if base.upper().startswith("3.2."):
                head = re.split(r"[\s\-_]", base, maxsplit=1)[0]
                if self.CTD_PATTERN.match(head):
                    candidates.add(self._normalize(head))

            for c in candidates:
                index.setdefault(c, []).append(pdf)
        return index

    @staticmethod
    def _pick_best(candidates: list[Path]) -> Path:
        # Prefer shorter filename token and deterministic ordering.
        sorted_candidates = sorted(candidates, key=lambda p: (len(p.stem), str(p).lower()))
        return sorted_candidates[0]

    @staticmethod
    def _section_flexible_regex(section: str) -> re.Pattern[str]:
        parts = [re.escape(p) for p in section.split(".") if p]
        if not parts:
            return re.compile(re.escape(section), re.IGNORECASE)
        pat = r"\b" + r"\s*[\.\s_\-]\s*".join(parts) + r"\b"
        return re.compile(pat, re.IGNORECASE)

    def _pdf_contains_section(self, pdf: Path, section_re: re.Pattern[str], *, max_pages: int = 3) -> bool:
        try:
            import fitz  # PyMuPDF
        except Exception:
            return False

        try:
            with fitz.open(pdf) as doc:
                limit = min(max_pages, doc.page_count)
                for i in range(limit):
                    try:
                        text = doc.load_page(i).get_text("text", sort=True)
                    except Exception:
                        continue
                    if section_re.search(text):
                        return True
        except Exception:
            return False
        return False

    def _content_scan(self, refer_section: str) -> list[Path]:
        key = self._normalize(refer_section)
        cached = self._content_index.get(key)
        if cached is not None:
            return cached

        section_re = self._section_flexible_regex(key)
        hits: list[Path] = []
        for pdf in self._pdfs:
            if self._pdf_contains_section(pdf, section_re):
                hits.append(pdf)
        self._content_index[key] = hits
        return hits

    def resolve(self, refer_section: str) -> ResolvedSection:
        normalized = self._normalize(refer_section)
        exact = self._index.get(normalized)
        if exact:
            return ResolvedSection(
                refer_section=refer_section,
                resolved_pdf=self._pick_best(exact),
                warning=None,
            )

        # Fallback by trimming last numeric node: 3.2.S.1.1 -> 3.2.S.1
        parts = normalized.split(".")
        cursor = parts
        while len(cursor) > 4:
            cursor = cursor[:-1]
            key = ".".join(cursor)
            candidates = self._index.get(key)
            if candidates:
                chosen = self._pick_best(candidates)
                return ResolvedSection(
                    refer_section=refer_section,
                    resolved_pdf=chosen,
                    warning=(
                        f"No exact PDF for {refer_section}; used fallback source {chosen.name}"
                    ),
                )

        # Final heuristic: prefix match by section root.
        prefix = normalized
        heuristic_hits = []
        for token, paths in self._index.items():
            if token.startswith(prefix) or prefix.startswith(token):
                heuristic_hits.extend(paths)

        if heuristic_hits:
            chosen = self._pick_best(heuristic_hits)
            return ResolvedSection(
                refer_section=refer_section,
                resolved_pdf=chosen,
                warning=f"Heuristic match used for {refer_section}: {chosen.name}",
            )

        # Last-resort: scan PDF content for the section anchor (robust to poor filenames).
        scanned = self._content_scan(refer_section)
        if scanned:
            chosen = self._pick_best(scanned)
            return ResolvedSection(
                refer_section=refer_section,
                resolved_pdf=chosen,
                warning=f"Content-scan match used for {refer_section}: {chosen.name}",
            )

        raise FileNotFoundError(f"Unable to map refer section: {refer_section}")
