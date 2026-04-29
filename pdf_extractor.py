from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import fitz

from section_mapper import ResolvedSection


@dataclass
class ExtractedSectionContent:
    refer_section: str
    source_pdf: Path
    anchor_found: bool
    warning: str | None
    raw_text: str
    tables: list[list[list[str]]] = field(default_factory=list)
    image_paths: list[Path] = field(default_factory=list)


class PdfSectionExtractor:
    def __init__(self, images_dir: Path) -> None:
        self.images_dir = images_dir

    @staticmethod
    def _section_flexible_regex(section: str) -> str:
        parts = [re.escape(p) for p in section.split(".") if p]
        if not parts:
            return re.escape(section)
        # Accept dot and/or whitespace separators and optional extra whitespace.
        return r"\b" + r"\s*[\.\s]\s*".join(parts) + r"\b"

    @staticmethod
    def _anchor_patterns(refer_section: str) -> list[re.Pattern[str]]:
        escaped = re.escape(refer_section)
        flexible = PdfSectionExtractor._section_flexible_regex(refer_section)
        return [
            re.compile(rf"\b{escaped}\b", re.IGNORECASE),
            re.compile(flexible, re.IGNORECASE),
        ]

    @staticmethod
    def _next_section_candidates(refer_section: str) -> list[str]:
        # e.g. 3.2.S.1.1 -> 3.2.S.1.2
        parts = refer_section.split(".")
        if not parts:
            return []
        try:
            n = int(parts[-1])
        except ValueError:
            return []
        sibling = parts[:-1] + [str(n + 1)]
        return [".".join(sibling)]

    def _find_page_with_patterns(self, doc: fitz.Document, patterns: list[re.Pattern[str]]) -> int | None:
        def compact(s: str) -> str:
            return re.sub(r"[^0-9A-Za-z]", "", s).lower()

        # Fallback target from first regex pattern body (best-effort extraction).
        target = ""
        if patterns:
            raw = patterns[0].pattern
            target = compact(raw)

        for i in range(doc.page_count):
            text = doc.load_page(i).get_text("text", sort=True)
            if any(p.search(text) for p in patterns):
                return i
            if target:
                if target in compact(text):
                    return i
        return None

    def _extract_tables(self, page: fitz.Page) -> list[list[list[str]]]:
        tables_out: list[list[list[str]]] = []
        try:
            found = page.find_tables()
            for tab in found.tables:
                rows = tab.extract()
                rows_normalized: list[list[str]] = []
                for row in rows:
                    rows_normalized.append(["" if cell is None else str(cell) for cell in row])
                if rows_normalized:
                    tables_out.append(rows_normalized)
        except Exception:
            # Keep extractor resilient; table detection can fail on some PDFs.
            return tables_out
        return tables_out

    def _extract_images(self, doc: fitz.Document, page_index: int, end_page: int, refer_section: str) -> list[Path]:
        paths: list[Path] = []
        candidate_pages = [page_index]
        if refer_section == "3.2.S.2.2":
            candidate_pages = list(range(page_index, min(end_page, page_index + 2) + 1))

        ranked: list[tuple[int, Path]] = []

        img_idx = 0
        for pidx in candidate_pages:
            page = doc.load_page(pidx)
            image_list = page.get_images(full=True)
            for img in image_list:
                img_idx += 1
                xref = img[0]
                out = self.images_dir / f"{refer_section.replace('.', '_')}_p{pidx+1}_{img_idx}.png"
                try:
                    pix = fitz.Pixmap(doc, xref)
                    if pix.alpha:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    area = pix.width * pix.height
                    pix.save(out)
                    ranked.append((area, out))
                except Exception:
                    continue

        if ranked:
            ranked.sort(key=lambda x: x[0], reverse=True)
            # Keep the single most likely content image; avoids tiny logos/headers.
            paths.append(ranked[0][1])

        # Fallback rendering only for structure section where image is expected.
        if not paths and refer_section in {"3.2.S.1.2", "3.2.S.2.2"}:
            render_idx = page_index
            if refer_section == "3.2.S.2.2":
                # Prefer a page that mentions Figure / Flow Diagram for visual extraction.
                for pidx in range(page_index, min(end_page, page_index + 4) + 1):
                    txt = doc.load_page(pidx).get_text("text", sort=True).lower()
                    if "figure" in txt or "flow diagram" in txt:
                        render_idx = pidx
                        break
            page = doc.load_page(render_idx)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            out = self.images_dir / f"{refer_section.replace('.', '_')}_p{render_idx+1}_render.png"
            pix.save(out)
            paths.append(out)

        return paths

    @staticmethod
    def _crop_text_between_anchors(raw_text: str, refer_section: str, next_candidates: list[str]) -> str:
        def anchor_regex(section: str) -> str:
            return PdfSectionExtractor._section_flexible_regex(section)

        start_m = re.search(anchor_regex(refer_section), raw_text, flags=re.IGNORECASE)
        if not start_m:
            return raw_text

        tail = raw_text[start_m.start():]
        end_pos = len(tail)
        for nxt in next_candidates:
            em = re.search(anchor_regex(nxt), tail, flags=re.IGNORECASE)
            if em and em.start() > 0:
                end_pos = min(end_pos, em.start())
        return tail[:end_pos].strip()

    @staticmethod
    def _remove_noise_lines(text: str) -> str:
        lines = [ln.rstrip() for ln in text.splitlines()]
        freq: dict[str, int] = {}
        for ln in lines:
            key = re.sub(r"\s+", " ", ln.strip())
            if not key:
                continue
            freq[key] = freq.get(key, 0) + 1

        cleaned: list[str] = []
        for ln in lines:
            key = re.sub(r"\s+", " ", ln.strip())
            if not key:
                cleaned.append(ln)
                continue
            if re.fullmatch(r"\d+\s+of\s+\d+", key, flags=re.IGNORECASE):
                continue
            if key.lower().startswith("unique pharmaceutical laboratories"):
                continue
            if key.lower().startswith("(a div. of"):
                continue
            # Drop aggressively repeated header/footer noise.
            if freq.get(key, 0) >= 3 and len(key) > 12:
                continue
            cleaned.append(ln)

        # Collapse excessive blank lines.
        out: list[str] = []
        blank_run = 0
        for ln in cleaned:
            if not ln.strip():
                blank_run += 1
                if blank_run <= 1:
                    out.append(ln)
            else:
                blank_run = 0
                out.append(ln)
        return "\n".join(out).strip()

    def extract(self, resolved: ResolvedSection) -> ExtractedSectionContent:
        refer = resolved.refer_section
        source_pdf = resolved.resolved_pdf

        with fitz.open(source_pdf) as doc:
            start_patterns = self._anchor_patterns(refer)
            start_page = self._find_page_with_patterns(doc, start_patterns)

            anchor_found = start_page is not None
            warning = resolved.warning

            if start_page is None:
                start_page = 0
                if warning:
                    warning = warning + f" | Anchor {refer} not found in PDF; used page-0 best guess"
                else:
                    warning = f"Anchor {refer} not found in PDF; used page-0 best guess"

            next_candidates = self._next_section_candidates(refer)
            end_page = doc.page_count - 1
            if next_candidates:
                for candidate in next_candidates:
                    cand_patterns = self._anchor_patterns(candidate)
                    cand_page = self._find_page_with_patterns(doc, cand_patterns)
                    if cand_page is not None and cand_page >= start_page:
                        end_page = max(start_page, cand_page - 1)
                        break

            page_texts: list[str] = []
            all_tables: list[list[list[str]]] = []
            all_images: list[Path] = []

            for pidx in range(start_page, end_page + 1):
                page = doc.load_page(pidx)
                txt = page.get_text("text", sort=True)
                page_texts.append(txt)
                all_tables.extend(self._extract_tables(page))

            all_images.extend(self._extract_images(doc, start_page, end_page, refer))

        raw_text = "\n\n".join(page_texts)
        raw_text = self._crop_text_between_anchors(raw_text, refer, next_candidates)
        raw_text = self._remove_noise_lines(raw_text)

        return ExtractedSectionContent(
            refer_section=refer,
            source_pdf=source_pdf,
            anchor_found=anchor_found,
            warning=warning,
            raw_text=raw_text,
            tables=all_tables,
            image_paths=all_images,
        )
