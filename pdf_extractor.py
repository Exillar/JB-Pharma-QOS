"""
Module: pdf_extractor
Responsibility: Extracts text, tables and images from CTD Module 3 PDFs.

Key design decisions
---------------------
- All tuning constants (DPI scale, crop fractions, drawing thresholds, keywords)
  live in `config_loader.DiagramConfig` / `NoiseConfig` — nothing is hardcoded
  in this module.
- Vector-drawn flow diagrams (no embedded image XREFs) are detected via a
  two-signal heuristic: drawing count + chars-per-drawing ratio.  Both
  thresholds are config-driven.
- Header/footer noise is removed by two complementary strategies:
    1. Auto-detection: frequency analysis of top/bottom margin text across
       pages (ported from JB-Pharma-QIS) — works for any dossier.
    2. Config-supplied prefix list: hard-coded company names in config.yaml
       (overridable per deployment).
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import fitz

from config_loader import DiagramConfig, NoiseConfig
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


# ---------------------------------------------------------------------------
# Auto noise-blocklist (ported from JB-Pharma-QIS / pdf_extractor.py)
# ---------------------------------------------------------------------------

def _build_noise_blocklist(doc: fitz.Document, cfg: NoiseConfig) -> set[str]:
    """Detect header/footer text by frequency analysis of page margin zones.

    Any normalised text string that appears in the top or bottom margin of
    at least `cfg.noise_page_threshold` pages is considered noise and added
    to the returned blocklist.
    """
    total = doc.page_count
    if total <= 1:
        return set()

    page_sets: list[set[str]] = []
    for page in doc:
        h = page.rect.height
        w = page.rect.width
        page_set: set[str] = set()
        clips = [
            fitz.Rect(0, 0, w, h * cfg.noise_top_margin_frac),
            fitz.Rect(0, h * (1.0 - cfg.noise_bottom_margin_frac), w, h),
        ]
        for clip in clips:
            for block in page.get_text("blocks", clip=clip):
                text = block[4].strip() if len(block) > 4 else ""
                if not text:
                    continue
                norm = " ".join(text.lower().split())
                if len(norm) >= 3:
                    page_set.add(norm)
        page_sets.append(page_set)

    freq: Counter[str] = Counter()
    for ps in page_sets:
        for t in ps:
            freq[t] += 1

    threshold = min(cfg.noise_page_threshold, total)
    return {t for t, c in freq.items() if c >= threshold}


# ---------------------------------------------------------------------------
# Main extractor class
# ---------------------------------------------------------------------------

class PdfSectionExtractor:
    """Extracts content from a single resolved CTD section PDF."""

    def __init__(
        self,
        images_dir: Path,
        diagram_cfg: DiagramConfig | None = None,
        noise_cfg: NoiseConfig | None = None,
        preserve_keywords: tuple[str, ...] = (),
    ) -> None:
        self.images_dir = images_dir
        self._dcfg = diagram_cfg or DiagramConfig()
        self._ncfg = noise_cfg or NoiseConfig()
        self._preserve_keywords = tuple(k.lower() for k in preserve_keywords if k)

    # ------------------------------------------------------------------
    # Section-ID helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _section_flexible_regex(section: str) -> str:
        """Build a regex that matches a CTD section ID with flexible separators."""
        parts = [re.escape(p) for p in section.split(".") if p]
        if not parts:
            return re.escape(section)
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
        """Return the most likely sibling section ID (e.g. 3.2.S.1.1 → 3.2.S.1.2)."""
        parts = refer_section.split(".")
        if not parts:
            return []
        try:
            n = int(parts[-1])
        except ValueError:
            return []
        return [".".join(parts[:-1] + [str(n + 1)])]

    # ------------------------------------------------------------------
    # Page-finding helpers
    # ------------------------------------------------------------------

    def _find_page_with_patterns(
        self, doc: fitz.Document, patterns: list[re.Pattern[str]]
    ) -> int | None:
        def compact(s: str) -> str:
            return re.sub(r"[^0-9A-Za-z]", "", s).lower()

        # Compact form of the first pattern for a last-resort character-level match.
        target = compact(patterns[0].pattern) if patterns else ""

        for i in range(doc.page_count):
            text = doc.load_page(i).get_text("text", sort=True)
            if any(p.search(text) for p in patterns):
                return i
            if target and target in compact(text):
                return i
        return None

    # ------------------------------------------------------------------
    # Table extraction
    # ------------------------------------------------------------------

    @staticmethod
    def _extract_tables(page: fitz.Page) -> list[list[list[str]]]:
        tables_out: list[list[list[str]]] = []
        try:
            found = page.find_tables()
            for tab in found.tables:
                rows = [
                    ["" if cell is None else str(cell) for cell in row]
                    for row in tab.extract()
                ]
                if rows:
                    tables_out.append(rows)
        except Exception:
            pass
        return tables_out

    # ------------------------------------------------------------------
    # Diagram-page detection (config-driven, no drug-specific strings)
    # ------------------------------------------------------------------

    def _is_diagram_page(self, page: fitz.Page, *, require_keywords: bool = True) -> bool:
        """Return True when the page is primarily a vector-drawn flow diagram.

        Two signals:
        1. Drawing count ≥ min_diagram_drawings (rules out text-only pages).
        2. Chars-per-drawing ratio ≤ threshold (rules out text-heavy pages
           whose thin table-border lines inflate the drawing count).
        3. At least one generic flow/process keyword present in the text.
        """
        n_drawings = len(page.get_drawings())
        if n_drawings < self._dcfg.min_diagram_drawings:
            return False

        txt = page.get_text("text", sort=True)
        n_chars = len(txt.strip())
        chars_per_drawing = n_chars / max(1, n_drawings)
        if chars_per_drawing > self._dcfg.chars_per_drawing_threshold:
            return False

        if not require_keywords:
            return True

        txt_lower = txt.lower()
        if any(kw in txt_lower for kw in self._dcfg.diagram_exclude_keywords):
            return False
        return any(kw in txt_lower for kw in self._dcfg.diagram_keywords)

    def _render_diagram_page(self, page: fitz.Page, out_path: Path) -> Path:
        """Render a single page to PNG cropped to diagram drawings only."""
        rect = page.rect
        scale = self._dcfg.render_dpi_scale

        clip = self._diagram_drawing_clip(page)
        if clip is None:
            clip = fitz.Rect(
                rect.x0,
                rect.y0 + rect.height * self._dcfg.header_crop_frac,
                rect.x1,
                rect.y1 - rect.height * self._dcfg.footer_crop_frac,
            )
        stop_y = self._find_diagram_tail_marker_y(page)
        if stop_y is not None and stop_y > clip.y0 + 24:
            clip = fitz.Rect(clip.x0, clip.y0, clip.x1, min(clip.y1, stop_y - 4))
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
        pix.save(str(out_path))
        return out_path

    @staticmethod
    def _find_diagram_tail_marker_y(page: fitz.Page) -> float | None:
        """
        Find common subsection markers that indicate diagram content has ended.
        Keeps extraction generic and avoids including (b)/(c)/(d) prompt text.
        """
        marker_tokens = (
            "brief narrative description",
            "alternate processes",
            "reprocessing steps",
            "(b)",
            "(c)",
            "(d)",
        )
        ys: list[float] = []
        for token in marker_tokens:
            try:
                hits = page.search_for(token)
            except Exception:
                hits = []
            for r in hits:
                ys.append(r.y0)
        return min(ys) if ys else None

    def _diagram_drawing_clip(self, page: fitz.Page) -> fitz.Rect | None:
        drawings = page.get_drawings()
        if not drawings:
            return None

        union = None
        rect = page.rect
        top_cut = rect.y0 + rect.height * self._dcfg.header_crop_frac
        bottom_cut = rect.y1 - rect.height * self._dcfg.footer_crop_frac

        for drawing in drawings:
            rect = drawing.get("rect")
            if rect is None:
                continue
            if rect.y1 <= top_cut or rect.y0 >= bottom_cut:
                continue
            union = rect if union is None else union | rect

        if union is None:
            return None

        # Add small padding so borders are not clipped.
        pad_x = max(4.0, union.width * 0.01)
        pad_y = max(4.0, union.height * 0.01)
        clip = fitz.Rect(
            union.x0 - pad_x,
            union.y0 - pad_y,
            union.x1 + pad_x,
            union.y1 + pad_y,
        )
        return clip & page.rect

    # ------------------------------------------------------------------
    # Image extraction (vector-draw pages + embedded XREF fallback)
    # ------------------------------------------------------------------

    def _extract_images(
        self,
        doc: fitz.Document,
        page_index: int,
        end_page: int,
        refer_section: str,
    ) -> list[Path]:
        safe_name = refer_section.replace(".", "_")

        # ---- Vector-drawn diagrams (e.g. 3.2.S.2.2 flow diagrams) --------
        if refer_section in self._dcfg.vector_diagram_sections:
            diagram_paths: list[Path] = []
            for pidx in range(page_index, end_page + 1):
                page = doc.load_page(pidx)
                if not self._is_diagram_page(page):
                    continue
                out = self.images_dir / f"{safe_name}_p{pidx + 1}_diagram.png"
                self._render_diagram_page(page, out)
                diagram_paths.append(out)
            return diagram_paths

        # ---- Embedded image XREFs (default path) --------------------------
        page = doc.load_page(page_index)
        image_list = page.get_images(full=True)
        best_pix = None
        best_area = 0
        best_path: Path | None = None

        for img_idx, img in enumerate(image_list, start=1):
            xref = img[0]
            out = self.images_dir / f"{safe_name}_p{page_index + 1}_{img_idx}.png"
            try:
                pix = fitz.Pixmap(doc, xref)
                if pix.alpha:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                area = pix.width * pix.height
                if area > best_area:
                    best_area = area
                    best_pix = pix
                    best_path = out
            except Exception:
                continue

        if best_pix is not None and best_path is not None:
            # Save only the largest image (structural formula / main graphic).
            best_pix.save(str(best_path))
            return [best_path]

        # ---- Page-render fallback (e.g. 3.2.S.1.2 structural formula) ----
        if refer_section in self._dcfg.page_render_sections:
            out = self.images_dir / f"{safe_name}_p{page_index + 1}_render.png"
            scale = self._dcfg.render_dpi_scale
            page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False).save(str(out))
            return [out]

        return []

    # ------------------------------------------------------------------
    # Text cleaning
    # ------------------------------------------------------------------

    @staticmethod
    def _crop_text_between_anchors(
        raw_text: str, refer_section: str, next_candidates: list[str]
    ) -> str:
        def anchor_re(sec: str) -> str:
            return PdfSectionExtractor._section_flexible_regex(sec)

        m = re.search(anchor_re(refer_section), raw_text, flags=re.IGNORECASE)
        if not m:
            return raw_text

        tail = raw_text[m.start():]
        end_pos = len(tail)
        for nxt in next_candidates:
            em = re.search(anchor_re(nxt), tail, flags=re.IGNORECASE)
            if em and em.start() > 0:
                end_pos = min(end_pos, em.start())
        return tail[:end_pos].strip()

    def _remove_noise_lines(self, text: str, blocklist: set[str]) -> str:
        """Strip header/footer lines using both auto-detected blocklist and
        config-supplied company-name prefixes.
        """
        lines = [ln.rstrip() for ln in text.splitlines()]

        # Build frequency map for repeated-line detection.
        freq: Counter[str] = Counter()
        for ln in lines:
            key = re.sub(r"\s+", " ", ln.strip())
            if key:
                freq[key] += 1

        cleaned: list[str] = []
        for ln in lines:
            key = re.sub(r"\s+", " ", ln.strip())
            key_lower = key.lower()

            if not key:
                cleaned.append(ln)
                continue

            # Page-number pattern: "3 of 9", "12 of 12", etc.
            if re.fullmatch(r"\d+\s+of\s+\d+", key, flags=re.IGNORECASE):
                continue

            # Config-supplied company-name prefixes (overridable).
            if any(key_lower.startswith(pfx.lower()) for pfx in self._ncfg.company_name_prefixes):
                continue

            # Auto-detected blocklist from margin frequency analysis.
            norm = " ".join(key_lower.split())
            if norm in blocklist:
                continue

            # Aggressively repeated lines (>= 3× and non-trivial length).
            if freq[key] >= 3 and len(key) > 12:
                if self._preserve_keywords and any(k in key_lower for k in self._preserve_keywords):
                    cleaned.append(ln)
                    continue
                continue

            cleaned.append(ln)

        # Collapse excessive blank lines to at most one in a row.
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

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def extract(self, resolved: ResolvedSection) -> ExtractedSectionContent:
        refer = resolved.refer_section
        source_pdf = resolved.resolved_pdf

        with fitz.open(source_pdf) as doc:
            # Build noise blocklist once per document (auto-detection).
            noise_blocklist = _build_noise_blocklist(doc, self._ncfg)

            start_patterns = self._anchor_patterns(refer)
            start_page = self._find_page_with_patterns(doc, start_patterns)

            anchor_found = start_page is not None
            warning = resolved.warning

            if start_page is None:
                start_page = 0
                suffix = f"Anchor {refer} not found in PDF; used page-0 best guess"
                warning = f"{warning} | {suffix}" if warning else suffix

            next_candidates = self._next_section_candidates(refer)
            end_page = doc.page_count - 1
            for candidate in next_candidates:
                cand_page = self._find_page_with_patterns(
                    doc, self._anchor_patterns(candidate)
                )
                if cand_page is not None and cand_page >= start_page:
                    end_page = max(start_page, cand_page - 1)
                    break

            page_texts: list[str] = []
            all_tables: list[list[list[str]]] = []

            for pidx in range(start_page, end_page + 1):
                page = doc.load_page(pidx)
                page_texts.append(page.get_text("text", sort=True))
                all_tables.extend(self._extract_tables(page))

            all_images = self._extract_images(doc, start_page, end_page, refer)

        raw_text = "\n\n".join(page_texts)
        raw_text = self._crop_text_between_anchors(raw_text, refer, next_candidates)
        raw_text = self._remove_noise_lines(raw_text, noise_blocklist)

        return ExtractedSectionContent(
            refer_section=refer,
            source_pdf=source_pdf,
            anchor_found=anchor_found,
            warning=warning,
            raw_text=raw_text,
            tables=all_tables,
            image_paths=all_images,
        )
