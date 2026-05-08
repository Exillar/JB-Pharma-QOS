"""builders/s32.py — S32DocxFiller: fills QOS section 2.3.S.3.2 Impurities."""
from __future__ import annotations

import re
from pathlib import Path

import fitz
from docx import Document
from docx.document import Document as _Document

from builders.base import _DocxHelper, run_artifact_cleanup
from config_loader import DiagramConfig
from pdf_extractor import ExtractedSectionContent
from ctd_utils import section_flexible_regex


class S32DocxFiller(_DocxHelper):
    SECTION_START = "2.3.S.3.2 Impurities"
    SECTION_END = "2.3.S.4 Control"

    def __init__(
        self,
        template_docx: Path,
        filled_reference_docx: Path | None = None,
        *,
        images_dir: Path | None = None,
        diagram_cfg: DiagramConfig | None = None,
    ) -> None:
        self.template_docx = template_docx
        self.filled_reference_docx = filled_reference_docx
        self.images_dir = images_dir
        self.diagram_cfg = diagram_cfg or DiagramConfig()

    def _find_section_pages(
        self,
        doc: fitz.Document,
        refer_section: str,
        *,
        next_section: str | None,
    ) -> tuple[int, int]:
        start_re = re.compile(section_flexible_regex(refer_section), re.IGNORECASE)
        start_page = None
        for i in range(doc.page_count):
            if start_re.search(doc.load_page(i).get_text("text", sort=True)):
                start_page = i
                break
        if start_page is None:
            return (0, doc.page_count - 1)

        end_page = doc.page_count - 1
        if next_section:
            next_re = re.compile(section_flexible_regex(next_section), re.IGNORECASE)
            for i in range(start_page + 1, doc.page_count):
                if next_re.search(doc.load_page(i).get_text("text", sort=True)):
                    end_page = max(start_page, i - 1)
                    break
        return (start_page, end_page)

    @staticmethod
    def _norm_cell(s) -> str:
        if s is None:
            return ""
        return re.sub(r"\s+", " ", str(s).replace(" ", " ").strip())

    def _extract_api_impurity_rows(
        self,
        pdf_path: Path,
        *,
        refer_section: str = "3.2.S.3.2",
        next_section: str = "3.2.S.3.3",
    ) -> tuple[list[dict[str, object]], list[str]]:
        warnings: list[str] = []
        if self.images_dir is None:
            return ([], ["images_dir not configured"])
        self.images_dir.mkdir(parents=True, exist_ok=True)
        scale = float(self.diagram_cfg.render_dpi_scale or 2.0)

        def is_api_table(tab_rows: list[list[object]]) -> bool:
            window = tab_rows[:4]
            flat = " ".join(
                self._norm_cell(c) for r in window for c in (r[:3] if r else [])
            ).lower()
            return ("api-related impurity" in flat) and ("structure" in flat) and ("origin" in flat)

        entry_specs: list[dict[str, object]] = []
        try:
            with fitz.open(pdf_path) as doc:
                sp, ep = self._find_section_pages(doc, refer_section, next_section=next_section)
                for pidx in range(sp, ep + 1):
                    page = doc.load_page(pidx)
                    try:
                        found = page.find_tables()
                    except Exception:
                        continue
                    for tab in found.tables:
                        if getattr(tab, "col_count", 0) != 3:
                            continue
                        try:
                            extracted = tab.extract() or []
                        except Exception:
                            extracted = []
                        if not extracted or not is_api_table(extracted):
                            continue

                        row_stream: list[dict[str, object]] = []
                        for ridx, row_text in enumerate(extracted):
                            if ridx >= len(tab.rows):
                                break
                            c0 = self._norm_cell(row_text[0] if len(row_text) > 0 else "")
                            c2 = self._norm_cell(row_text[2] if len(row_text) > 2 else "")
                            low_all = f"{c0} {c2}".lower().strip()
                            if not low_all:
                                continue
                            if "api-related impurity" in low_all or low_all in {"structure", "origin", "structure origin"}:
                                continue
                            if "descriptor" in c0.lower() or "compendial" in c0.lower():
                                continue
                            bbox1 = None
                            try:
                                bbox1 = tab.rows[ridx].cells[1]
                            except Exception:
                                pass
                            row_stream.append({"name": c0, "origin": c2, "bbox": bbox1, "page_index": pidx})

                        current: dict[str, object] | None = None
                        for item in row_stream:
                            name = str(item["name"] or "").strip()
                            origin = str(item["origin"] or "").strip()
                            bbox = item["bbox"]
                            pindex = int(item["page_index"])
                            starts_new = bool(name) and (not name.startswith("(")) and bool(origin)
                            if current is None and not starts_new:
                                continue
                            if starts_new:
                                if current is not None:
                                    entry_specs.append(current)
                                current = {
                                    "name_parts": [name] if name else [],
                                    "origin_parts": [origin] if origin else [],
                                    "page_index": pindex,
                                    "bbox_union": None,
                                }
                            else:
                                if name:
                                    current["name_parts"].append(name)
                                if origin:
                                    current["origin_parts"].append(origin)
                            if current is not None and bbox is not None:
                                try:
                                    rect = fitz.Rect(*bbox) & page.rect
                                except Exception:
                                    rect = None
                                if rect is not None and not rect.is_empty:
                                    prev = current.get("bbox_union")
                                    current["bbox_union"] = rect if prev is None else (prev | rect)
                        if current is not None:
                            entry_specs.append(current)
        except Exception as e:
            return ([], [f"failed to read PDF {pdf_path.name}: {e}"])

        out_rows: list[dict[str, object]] = []
        try:
            with fitz.open(pdf_path) as doc:
                for i, e in enumerate(entry_specs, start=1):
                    name = " ".join(p for p in e.get("name_parts", []) if p).strip()
                    origin = " ".join(p for p in e.get("origin_parts", []) if p).strip()
                    page_index = int(e.get("page_index", 0))
                    rect = e.get("bbox_union")
                    image_path: Path | None = None

                    if rect is not None:
                        page = doc.load_page(page_index)
                        cell_rect = fitz.Rect(rect) & page.rect
                        if not cell_rect.is_empty:
                            inset_x = max(4.0, cell_rect.width * 0.02)
                            inset_y = max(4.0, cell_rect.height * 0.02)
                            inner = fitz.Rect(
                                cell_rect.x0 + inset_x, cell_rect.y0 + inset_y,
                                cell_rect.x1 - inset_x, cell_rect.y1 - inset_y,
                            ) & cell_rect

                            def accept(r: fitz.Rect) -> bool:
                                if r.is_empty:
                                    return False
                                if r.width >= cell_rect.width * 0.92 or r.height >= cell_rect.height * 0.92:
                                    return False
                                if inner.is_empty:
                                    return True
                                return not (r & inner).is_empty

                            union = None
                            try:
                                for d in page.get_drawings():
                                    drect = d.get("rect")
                                    if drect is None:
                                        continue
                                    r = fitz.Rect(drect) & cell_rect
                                    if accept(r):
                                        union = r if union is None else (union | r)
                            except Exception:
                                pass
                            try:
                                for img in page.get_images(full=True):
                                    xref = img[0]
                                    for r0 in page.get_image_rects(xref):
                                        r = fitz.Rect(r0) & cell_rect
                                        if accept(r):
                                            union = r if union is None else (union | r)
                            except Exception:
                                pass

                            clip0 = union or cell_rect
                            pad_x = max(2.0, clip0.width * 0.03)
                            pad_y = max(2.0, clip0.height * 0.03)
                            clip = fitz.Rect(
                                clip0.x0 - pad_x, clip0.y0 - pad_y,
                                clip0.x1 + pad_x, clip0.y1 + pad_y,
                            ) & page.rect
                            out = self.images_dir / f"3_2_S_3_2_api_structure_{i}.png"
                            try:
                                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
                                pix.save(str(out))
                                image_path = out
                            except Exception as img_err:
                                warnings.append(f"failed to render structure image for row {i}: {img_err}")

                    if name or origin or image_path:
                        out_rows.append({"name": name, "origin": origin, "structure_image": image_path})
        except Exception as e:
            warnings.append(f"failed to render images: {e}")

        return (out_rows, warnings)

    def _estimate_table_col_width_emu(self, doc: _Document, *, ncols: int) -> int:
        available = self._get_available_page_width(doc)
        return max(int(available / max(1, ncols)), int(914400 * 1.2))

    def fill_s32_section(
        self,
        extracted: dict[str, ExtractedSectionContent],
        output_docx: Path,
    ) -> list[str]:
        doc = Document(self.template_docx)
        template_table_count = len(doc.tables)
        warnings: list[str] = []

        start_idx, end_idx = self._get_target_range(doc)

        name_line = self._resolve_name_manufacturer_line(
            self.filled_reference_docx, "2.3.S.3.2 Impurities"
        )
        idx = self._find_para_index_doc(doc, "2.3.S.3.2 Impurities", start_idx, end_idx)
        if idx is not None and name_line:
            nxt = (doc.paragraphs[idx + 1].text or "").strip() if idx + 1 < len(doc.paragraphs) else ""
            if not nxt.startswith("("):
                self._insert_paragraph_after(doc.paragraphs[idx], name_line)

        payload = extracted.get("3.2.S.3.2")
        if not payload:
            warnings.append("3.2.S.3.2: missing extracted payload")
            rows: list[dict[str, object]] = []
        else:
            rows, ws = self._extract_api_impurity_rows(payload.source_pdf)
            warnings.extend(f"3.2.S.3.2: {w}" for w in ws)

        start_idx, end_idx = self._get_target_range(doc)
        anchor_idx = self._find_para_index_doc(doc, "List of API-related impurities", start_idx, end_idx)
        if anchor_idx is None:
            warnings.append("3.2.S.3.2: API impurities anchor paragraph not found in template")
        else:
            placeholder = None
            try:
                from docx.oxml.table import CT_Tbl
                from docx.oxml.text.paragraph import CT_P
                from docx.table import Table
                from docx.text.paragraph import Paragraph as _P

                def _iter_blocks(d: _Document):
                    for child in d.element.body.iterchildren():
                        if isinstance(child, CT_P):
                            yield _P(child, d)
                        elif isinstance(child, CT_Tbl):
                            yield Table(child, d)

                blocks = list(_iter_blocks(doc))
                anchor_p = doc.paragraphs[anchor_idx]
                anchor_block_i = None
                for i, b in enumerate(blocks):
                    if hasattr(b, "_p") and getattr(b, "_p") is anchor_p._p:
                        anchor_block_i = i
                        break
                if anchor_block_i is not None:
                    for b in blocks[anchor_block_i + 1: anchor_block_i + 6]:
                        if isinstance(b, Table):
                            placeholder = b
                            break
            except Exception:
                placeholder = None

            if placeholder is None:
                warnings.append("3.2.S.3.2: API impurity placeholder table not found in template")
            else:
                needed = 1 + len(rows)
                while len(placeholder.rows) < needed:
                    placeholder.add_row()

                col_width = self._estimate_table_col_width_emu(doc, ncols=len(placeholder.columns))
                try:
                    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
                except Exception:
                    WD_CELL_VERTICAL_ALIGNMENT = None

                for i, row in enumerate(rows, start=1):
                    name = str(row.get("name") or "")
                    origin = str(row.get("origin") or "")
                    img = row.get("structure_image")

                    self._safe_set_cell_text(placeholder.cell(i, 0), name)
                    self._safe_set_cell_text(placeholder.cell(i, 2), origin)
                    self._safe_set_cell_text(placeholder.cell(i, 1), "")

                    try:
                        if WD_CELL_VERTICAL_ALIGNMENT is not None:
                            placeholder.cell(i, 1).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    except Exception:
                        pass

                    if img and isinstance(img, Path) and img.exists():
                        p = placeholder.cell(i, 1).paragraphs[0]
                        try:
                            p.alignment = 1
                            p.paragraph_format.space_before = 0
                            p.paragraph_format.space_after = 0
                        except Exception:
                            pass
                        run = p.add_run()
                        try:
                            max_w = min(int(col_width * 0.90), int(914400 * 1.85))
                            max_h = int(914400 * 1.05)
                            self._add_picture_autofit_bounds(run, img, max_width_emu=max_w, max_height_emu=max_h)
                        except Exception as e:
                            warnings.append(f"3.2.S.3.2: failed to insert structure image for row {i}: {e}")

                if not rows:
                    warnings.append("3.2.S.3.2: no API impurity rows extracted from PDF")

        cleanup_stats = run_artifact_cleanup(doc, keep_first_n_tables=template_table_count)
        if any(cleanup_stats.values()):
            warnings.append(f"cleanup: {cleanup_stats}")

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_docx)
        return warnings
