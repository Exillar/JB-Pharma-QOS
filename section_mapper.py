from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ResolvedSection:
    refer_section: str
    resolved_pdf: Path
    match_mode: str  # exact | fallback
    warning: str | None = None


class SectionMapper:
    CTD_PATTERN = re.compile(r"(?i)(3\.2\.[A-Z]\.\d+(?:\.\d+)*)")

    def __init__(self, module3_root: Path) -> None:
        self.module3_root = module3_root
        self._pdfs = list(module3_root.rglob("*.pdf"))
        self._index = self._build_index(self._pdfs)

    @staticmethod
    def _normalize(token: str) -> str:
        return re.sub(r"\s+", "", token.strip().upper())

    def _build_index(self, pdfs: list[Path]) -> dict[str, list[Path]]:
        index: dict[str, list[Path]] = {}
        for pdf in pdfs:
            candidates = set()
            base = pdf.stem
            matches = self.CTD_PATTERN.findall(base)
            for m in matches:
                candidates.add(self._normalize(m))

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

    def resolve(self, refer_section: str) -> ResolvedSection:
        normalized = self._normalize(refer_section)
        exact = self._index.get(normalized)
        if exact:
            return ResolvedSection(
                refer_section=refer_section,
                resolved_pdf=self._pick_best(exact),
                match_mode="exact",
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
                    match_mode="fallback",
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
                match_mode="fallback",
                warning=f"Heuristic match used for {refer_section}: {chosen.name}",
            )

        raise FileNotFoundError(f"Unable to map refer section: {refer_section}")
