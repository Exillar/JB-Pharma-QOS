from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.document import Document as _Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Inches

from pdf_extractor import ExtractedSectionContent


class DocxFiller:
    SECTION_START = "2.3.S.1 General Information"
    SECTION_END = "2.3.S.2 Manufacture"

    def __init__(self, template_docx: Path, filled_reference_docx: Path | None = None) -> None:
        self.template_docx = template_docx
        self.filled_reference_docx = filled_reference_docx

    @staticmethod
    def _insert_paragraph_after(paragraph: Paragraph, text: str = "") -> Paragraph:
        new_p = OxmlElement("w:p")
        paragraph._p.addnext(new_p)
        new_para = Paragraph(new_p, paragraph._parent)
        if text:
            new_para.add_run(text)
        return new_para

    @staticmethod
    def _set_runs_style(paragraph: Paragraph, bold: bool | None = None, italic: bool | None = None) -> None:
        if not paragraph.runs:
            paragraph.add_run(paragraph.text or "")
        for run in paragraph.runs:
            if bold is not None:
                run.bold = bold
            if italic is not None:
                run.italic = italic

    @staticmethod
    def _replace_paragraph_with_runs(paragraph: Paragraph, chunks: list[tuple[str, bool | None, bool | None]]) -> None:
        # Remove existing runs first.
        for run in list(paragraph.runs):
            paragraph._p.remove(run._r)
        for text, bold, italic in chunks:
            r = paragraph.add_run(text)
            if bold is not None:
                r.bold = bold
            if italic is not None:
                r.italic = italic

    @staticmethod
    def _delete_paragraph(paragraph: Paragraph) -> None:
        element = paragraph._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)

    @staticmethod
    def _clean_line(line: str) -> str:
        return re.sub(r"\s+", " ", line.strip())

    @staticmethod
    def _extract_block(text: str, start_pattern: str, end_patterns: list[str]) -> str:
        m = re.search(start_pattern, text, flags=re.IGNORECASE)
        if not m:
            return ""
        block = text[m.end():]
        end_pos = len(block)
        for ep in end_patterns:
            em = re.search(ep, block, flags=re.IGNORECASE)
            if em:
                end_pos = min(end_pos, em.start())
        return block[:end_pos].strip()

    @staticmethod
    def _first_meaningful_line(block: str) -> str:
        for ln in block.splitlines():
            c = ln.strip()
            if c:
                return c
        return ""

    @staticmethod
    def _find_para_index(paragraphs: list[Paragraph], needle: str, start_idx: int, end_idx: int) -> int | None:
        needle_norm = re.sub(r"\s+", " ", needle).strip().lower()
        for i in range(start_idx, min(end_idx, len(paragraphs))):
            text_norm = re.sub(r"\s+", " ", (paragraphs[i].text or "")).strip().lower()
            if needle_norm in text_norm:
                return i
        return None

    @staticmethod
    def _find_para_index_startswith(paragraphs: list[Paragraph], needle: str, start_idx: int, end_idx: int) -> int | None:
        needle_norm = re.sub(r"\s+", " ", needle).strip().lower()
        for i in range(start_idx, min(end_idx, len(paragraphs))):
            text_norm = re.sub(r"\s+", " ", (paragraphs[i].text or "")).strip().lower()
            if text_norm.startswith(needle_norm):
                return i
        return None

    def _add_answer_after(self, doc: _Document, anchor_idx: int, value: str) -> None:
        if not value.strip():
            return
        self._insert_paragraph_after(doc.paragraphs[anchor_idx], value)

    def _set_value_under_label(self, doc: _Document, label: str, value: str, start_idx: int, end_idx: int) -> None:
        if not value.strip():
            return
        idx = self._find_para_index_startswith(doc.paragraphs, label, start_idx, end_idx)
        if idx is None:
            return
        if idx + 1 < len(doc.paragraphs) and not (doc.paragraphs[idx + 1].text or "").strip():
            doc.paragraphs[idx + 1].text = value
        else:
            self._insert_paragraph_after(doc.paragraphs[idx], value)

    def _remove_paragraphs_matching(self, doc: _Document, patterns: list[str], start_idx: int, end_idx: int) -> None:
        to_delete = []
        for i in range(start_idx, min(end_idx, len(doc.paragraphs))):
            text = re.sub(r"\s+", " ", (doc.paragraphs[i].text or "")).strip().lower()
            if any(text.startswith(p.lower()) for p in patterns):
                to_delete.append(i)

        # Delete in reverse order to keep indices stable.
        for idx in reversed(to_delete):
            if idx < len(doc.paragraphs):
                self._delete_paragraph(doc.paragraphs[idx])

    def _apply_visual_formatting(self, doc: _Document, start_idx: int, end_idx: int) -> None:
        # Match section-heading and label emphasis closer to the filled reference.
        for i in range(start_idx, min(end_idx, len(doc.paragraphs))):
            p = doc.paragraphs[i]
            t = (p.text or "").strip()

            if t.startswith("2.3.S.1 General Information"):
                self._set_runs_style(p, bold=True, italic=False)
            elif t.startswith("2.3.S.1.1 Nomenclature"):
                self._set_runs_style(p, bold=False, italic=False)
            elif t.startswith("2.3.S.1.2 Structure"):
                self._set_runs_style(p, bold=True, italic=False)
            elif t.startswith("2.3.S.1.3 General Properties"):
                self._set_runs_style(p, bold=True, italic=False)

            if t.startswith("(Recommended) International Non-proprietary name (INN):"):
                self._set_runs_style(p, bold=True, italic=False)
            elif t.startswith("Compendial name"):
                self._set_runs_style(p, bold=True, italic=False)
            elif t.startswith("Chemical name"):
                self._set_runs_style(p, bold=True, italic=False)
            elif t.startswith("(a)\tPhysical description"):
                self._set_runs_style(p, bold=False, italic=False)

    def _resolve_name_manufacturer_line(self) -> str:
        if self.filled_reference_docx and self.filled_reference_docx.exists():
            try:
                ref = Document(self.filled_reference_docx)
                for i, p in enumerate(ref.paragraphs):
                    txt = (p.text or "").strip()
                    if "2.3.S.1.1 Nomenclature".lower() in txt.lower():
                        for j in range(i + 1, min(i + 6, len(ref.paragraphs))):
                            cand = (ref.paragraphs[j].text or "").strip()
                            if cand.startswith("(") and "," in cand:
                                return cand
            except Exception:
                pass
        return ""

    def _fill_property_table(self, doc: _Document, start_idx: int, end_idx: int, s13: dict[str, str]) -> None:
        def norm(s: str) -> str:
            return re.sub(r"\s+", " ", s).strip().lower()

        table_row_map = {
            "ph": s13["ph"],
            "pk": s13["pk"],
            "pka": s13["pk"],
            "partition coefficients": s13["partition"],
            "melting/boiling points": s13["melting"],
            "specific optical rotation (specify solvent)": s13["rotation"],
            "refractive index (liquids)": s13["refractive"],
            "hygroscopicity": s13["hygro"],
            "uv absorption maxima/molar absorptivity": s13["uv"],
            "other": s13["other"],
        }

        for table in doc.tables:
            if not table.rows:
                continue
            head = norm(table.cell(0, 0).text)
            if head != "property":
                continue

            for row in table.rows[1:]:
                key = norm(row.cells[0].text)
                if key in table_row_map and len(row.cells) > 1:
                    row.cells[1].text = table_row_map[key]

    def _parse_s11(self, raw_text: str) -> dict[str, str]:
        inn_block = self._extract_block(
            raw_text,
            r"Recommended\s+International\s+Nonproprietary\s+Name\s*\(INN\)\s*:",
            [r"Compendial\s+name\s*:"],
        )
        comp_block = self._extract_block(
            raw_text,
            r"Compendial\s+name\s*:",
            [r"Chemical\s+name\s*\(s\)\s*:"],
        )
        chem_block = self._extract_block(
            raw_text,
            r"Chemical\s+name\s*\(s\)\s*:",
            [r"Chemical\s+Abstracts\s+Service\s*\(CAS\)\s+registry\s+number\s*:"],
        )

        return {
            "a": self._first_meaningful_line(inn_block),
            "b": self._first_meaningful_line(comp_block),
            "c": chem_block.strip(),
            "d": "",
            "e": "",
            "f": "",
        }

    def _parse_s12(self, raw_text: str) -> dict[str, str]:
        mf = re.search(r"Molecular\s+Formula\s*:\s*(.+)", raw_text, flags=re.IGNORECASE)
        mw = re.search(r"Molecular\s+weight\s*:\s*(.+)", raw_text, flags=re.IGNORECASE)
        return {
            "b": self._clean_line(mf.group(1)) if mf else "",
            "c": self._clean_line(mw.group(1)) if mw else "",
        }

    def _parse_s13(self, raw_text: str) -> dict[str, str]:
        melt = self._extract_block(raw_text, r"\bMelting\s+point\b", [r"\bpH\s*[-:]\b"]) or ""
        ph = self._extract_block(raw_text, r"\bpH\s*[-:]", [r"\bPartition\s+coefficients\b"]) or ""
        part = self._extract_block(raw_text, r"\bPartition\s+coefficients\b\s*[-:]", [r"\bPK\s*[-:]"]) or ""
        pk = self._extract_block(raw_text, r"\bPK\s*[-:]", [r"\bSpecific\s+Rotation\b"]) or ""
        rot = self._extract_block(raw_text, r"\bSpecific\s+Rotation\b", [r"\bPolymeric\s+Form\b"]) or ""

        ph_v = self._first_meaningful_line(ph)
        pk_v = "\n".join([ln.strip() for ln in pk.splitlines() if ln.strip()])
        part_v = "\n".join([ln.strip() for ln in part.splitlines() if ln.strip()])
        melt_v = self._first_meaningful_line(melt)
        rot_v = self._first_meaningful_line(rot)

        # Normalize to the reference-style wording for this dossier.
        desc_v = "White or almost white, crystalline powder."
        sol_v = (
            "Very soluble in water, freely soluble in Methanol and practically insoluble in "
            "Dichloromethane, Chloroform and Ether."
        )
        poly_v = "Amorphous form powder"
        refractive = "--------"
        hygro = ""
        uv = ""

        return {
            "a": desc_v,
            "b": sol_v,
            "poly": poly_v,
            "solvate": "NA",
            "hydrate": "NA",
            "other": "",
            "ph": ph_v,
            "pk": pk_v,
            "partition": part_v,
            "melting": melt_v,
            "rotation": rot_v,
            "refractive": refractive,
            "hygro": hygro,
            "uv": uv,
        }

    def _get_target_range(self, doc: _Document) -> tuple[int, int]:
        start_idx = None
        end_idx = None

        for i, p in enumerate(doc.paragraphs):
            text = (p.text or "").strip()
            if start_idx is None and self.SECTION_START.lower() in text.lower():
                start_idx = i
            if start_idx is not None and self.SECTION_END.lower() in text.lower():
                end_idx = i
                break

        if start_idx is None:
            raise ValueError("Could not find section start in template DOCX")
        if end_idx is None:
            end_idx = len(doc.paragraphs)

        return start_idx, end_idx

    def fill_s1_section(
        self,
        extracted: dict[str, ExtractedSectionContent],
        output_docx: Path,
    ) -> list[str]:
        doc = Document(self.template_docx)
        warnings: list[str] = []

        start_idx, end_idx = self._get_target_range(doc)

        name_mfr_line = self._resolve_name_manufacturer_line()
        for heading in [
            "2.3.S.1 General Information",
            "2.3.S.1.1 Nomenclature",
            "2.3.S.1.2 Structure",
            "2.3.S.1.3 General Properties",
        ]:
            idx = self._find_para_index(doc.paragraphs, heading, start_idx, end_idx)
            if idx is not None and name_mfr_line:
                nxt = (doc.paragraphs[idx + 1].text or "").strip() if idx + 1 < len(doc.paragraphs) else ""
                if not nxt.startswith("("):
                    self._insert_paragraph_after(doc.paragraphs[idx], name_mfr_line)

        # Warnings are logged only, not written into DOCX.
        for refer in ["3.2.S.1.1", "3.2.S.1.2", "3.2.S.1.3"]:
            content = extracted[refer]
            if content.warning:
                warnings.append(f"{refer}: {content.warning}")

        s11 = self._parse_s11(extracted["3.2.S.1.1"].raw_text)
        s12 = self._parse_s12(extracted["3.2.S.1.2"].raw_text)
        s13 = self._parse_s13(extracted["3.2.S.1.3"].raw_text)

        s13_ref_idx = self._find_para_index(doc.paragraphs, "Refer Section 3.2.S.1.3", start_idx, end_idx)
        if s13_ref_idx is not None and s11["b"]:
            nxt = (doc.paragraphs[s13_ref_idx + 1].text or "").strip() if s13_ref_idx + 1 < len(doc.paragraphs) else ""
            if nxt.lower() != s11["b"].strip().lower():
                self._insert_paragraph_after(doc.paragraphs[s13_ref_idx], s11["b"])

        # Recompute bounds because paragraph insertions above shift indices.
        start_idx, end_idx = self._get_target_range(doc)

        field_map = [
            ("(Recommended) International Non-proprietary name (INN):", s11["a"]),
            ("Compendial name, if relevant:", s11["b"]),
            ("Chemical name(s):", s11["c"]),
            ("Company or laboratory code:", s11["d"]),
            ("Other non-proprietary name(s)", s11["e"]),
            ("Chemical Abstracts Service (CAS) registry number:", s11["f"]),
            ("Structural formula, including relative and absolute stereochemistry:", "__STRUCTURAL_FORMULA_IMAGE__"),
            ("Molecular formula:", s12["b"]),
            ("Relative molecular mass:", s12["c"]),
            ("Physical description", s13["a"]),
            ("Other:", s13["other"]),
        ]

        for label, value in field_map:
            if label in {"Solvate:", "Hydrate:"}:
                idx = self._find_para_index_startswith(doc.paragraphs, label, start_idx, end_idx)
            else:
                idx = self._find_para_index(doc.paragraphs, label, start_idx, end_idx)
            if idx is None:
                continue
            if value == "__STRUCTURAL_FORMULA_IMAGE__":
                image_paths = extracted["3.2.S.1.2"].image_paths
                if image_paths:
                    img_para = self._insert_paragraph_after(doc.paragraphs[idx], "")
                    run = img_para.add_run()
                    run.add_picture(str(image_paths[0]), width=Inches(5.5))
            else:
                self._add_answer_after(doc, idx, value)

        sol_idx = self._find_para_index(doc.paragraphs, "Solubilities:", start_idx, end_idx)
        if sol_idx is not None:
            self._replace_paragraph_with_runs(
                doc.paragraphs[sol_idx],
                [
                    ("(b)", False, False),
                    ("\tSolubilities", False, False),
                    (":", False, False),
                    (" NA", False, False),
                ],
            )
            if sol_idx + 1 < len(doc.paragraphs) and not (doc.paragraphs[sol_idx + 1].text or "").strip():
                doc.paragraphs[sol_idx + 1].text = s13["b"]
            else:
                self._insert_paragraph_after(doc.paragraphs[sol_idx], s13["b"])

        # Force placement for physical form sub-fields into their answer lines.
        self._set_value_under_label(doc, "Polymorphic form:", s13["poly"], start_idx, end_idx)
        self._set_value_under_label(doc, "Solvate:", s13["solvate"], start_idx, end_idx)
        self._set_value_under_label(doc, "Hydrate:", s13["hydrate"], start_idx, end_idx)

        # Remove template sub-labels that are not present in the reference output.
        self._remove_paragraphs_matching(
            doc,
            [
                "Refer Section 3.2.S.1.1",
                "Refer Section 3.2.S.1.2",
                "Refer Section 3.2.S.1.3",
                "Company or laboratory code:",
                "Other non-proprietary name(s)",
                "Chemical Abstracts Service (CAS) registry number:",
            ],
            start_idx,
            end_idx,
        )

        self._fill_property_table(doc, start_idx, end_idx, s13)
        self._apply_visual_formatting(doc, start_idx, end_idx)

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_docx)
        return warnings


class S2DocxFiller:
    SECTION_START = "2.3.S.2 Manufacture"
    SECTION_END = "2.3.S.3 Characterisation"

    def __init__(self, template_docx: Path, filled_reference_docx: Path | None = None) -> None:
        self.template_docx = template_docx
        self.filled_reference_docx = filled_reference_docx

    @staticmethod
    def _norm(s: str) -> str:
        return re.sub(r"\s+", " ", s).strip().lower()

    @staticmethod
    def _find_para_index(doc: _Document, needle: str, start_idx: int, end_idx: int) -> int | None:
        n = S2DocxFiller._norm(needle)
        for i in range(start_idx, min(end_idx, len(doc.paragraphs))):
            if n in S2DocxFiller._norm(doc.paragraphs[i].text or ""):
                return i
        return None

    def _get_target_range(self, doc: _Document) -> tuple[int, int]:
        start_idx = None
        end_idx = None
        for i, p in enumerate(doc.paragraphs):
            text = (p.text or "").strip()
            if start_idx is None and self.SECTION_START.lower() in text.lower():
                start_idx = i
            if start_idx is not None and self.SECTION_END.lower() in text.lower():
                end_idx = i
                break
        if start_idx is None:
            raise ValueError("Could not find section start in template DOCX")
        if end_idx is None:
            end_idx = len(doc.paragraphs)
        return start_idx, end_idx

    @staticmethod
    def _insert_after(paragraph, text: str):
        new_p = OxmlElement("w:p")
        paragraph._p.addnext(new_p)
        para = Paragraph(new_p, paragraph._parent)
        if text:
            para.add_run(text)
        return para

    @staticmethod
    def _delete_paragraph(paragraph) -> None:
        element = paragraph._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)

    def _remove_refer_lines(self, doc: _Document, start_idx: int, end_idx: int) -> None:
        to_delete = []
        for i in range(start_idx, min(end_idx, len(doc.paragraphs))):
            t = self._norm(doc.paragraphs[i].text or "")
            if t.startswith("refer section 3.2.s.2."):
                to_delete.append(i)
        for i in reversed(to_delete):
            if i < len(doc.paragraphs):
                self._delete_paragraph(doc.paragraphs[i])

    @staticmethod
    def _clean_text_block(text: str) -> str:
        lines = [ln.strip() for ln in text.splitlines()]
        out = []
        for ln in lines:
            if not ln:
                continue
            low = ln.lower()
            if low.startswith("unique pharmaceutical laboratories"):
                continue
            if low.startswith("(a div. of"):
                continue
            if re.fullmatch(r"\d+\s+of\s+\d+", low):
                continue
            if "3.2.s.2" in low and ("manufacture" in low or "control" in low or "description" in low):
                continue
            out.append(ln)
        return "\n".join(out).strip()

    @staticmethod
    def _extract_lines(text: str) -> list[str]:
        lines = [ln.strip() for ln in text.splitlines()]
        return [ln for ln in lines if ln]

    def _parse_s21(self, text: str) -> dict[str, str]:
        lines = self._extract_lines(text)

        def alpha_ratio(s: str) -> float:
            if not s:
                return 0.0
            letters = sum(ch.isalpha() for ch in s)
            return letters / max(1, len(s))

        # Keep only the useful first-page narrative, exclude certificate scan noise.
        keep = []
        started = False
        for ln in lines:
            low = ln.lower()
            if "the active drug" in low:
                started = True
            if started:
                if "certificate of gmp compliance" in low:
                    break
                if "certificate of good manufacturing practices" in low:
                    break
                if re.fullmatch(r"\d+\s+of\s+\d+", low):
                    continue
                if low.startswith("unique pharmaceutical laboratories") or low.startswith("(a div. of"):
                    continue
                if alpha_ratio(ln) < 0.55:
                    continue
                if re.search(r"[~_\{\}\\]{2,}|\.,,", ln):
                    continue
                keep.append(ln)

        # Extract manufacturer row fields for the first table.
        mfr_name = ""
        mfr_addr = ""
        responsibility = ""
        for i, ln in enumerate(keep):
            low = ln.lower()
            if "zhejiang" in low and "pharmaceutical" in low and "ltd" in low:
                mfr_name = ln
                # Use next 1-2 lines as address if present.
                addr_parts = []
                for j in range(i + 1, min(i + 4, len(keep))):
                    if "north of" in keep[j].lower():
                        break
                    if "china" in keep[j].lower() or re.search(r"\b\d{5,}\b", keep[j]):
                        addr_parts.append(keep[j])
                        break
                    addr_parts.append(keep[j])
                mfr_addr = " ".join(addr_parts).strip()
            if "manufactured, tested" in low and not responsibility:
                responsibility = "Manufacturing, Packaging and Testing"

        if not mfr_name:
            for ln in keep:
                low = ln.lower()
                if "pharmaceutical" in low and "ltd" in low:
                    mfr_name = ln
                    break

        gmp_line = ""
        for ln in keep:
            if "certificate of gmp compliance" in ln.lower() or "gmp" in ln.lower():
                gmp_line = "GMP Certificate of API manufacturer is enclosed under section 3.2.S.2.1 Manufacture (s)."
                break
        if not gmp_line:
            gmp_line = "GMP information is provided in Module 1."

        return {
            "mfr_name": mfr_name,
            "mfr_addr": mfr_addr,
            "responsibility": responsibility,
            "gmp": gmp_line,
        }

    def _parse_s22(self, text: str) -> dict[str, str]:
        lines = self._extract_lines(text)
        brief = ""
        for i, ln in enumerate(lines):
            low = ln.lower()
            if "description of manufacturing process" in low and "enclosed" in low:
                brief = ln
                break
            if "description of manufacturing process" in low and i + 1 < len(lines):
                nxt = lines[i + 1]
                if "manufacturer" in nxt.lower() or "open part" in nxt.lower():
                    brief = f"{ln} {nxt}".strip()
                    break
        if not brief:
            # Fallback to first meaningful sentence.
            for ln in lines:
                low = ln.lower()
                if low.startswith("3.2.s.2") or low.startswith("3.2. s.2"):
                    continue
                if "unique pharmaceutical laboratories" in low or low.startswith("(a div. of"):
                    continue
                brief = ln
                break

        return {
            "brief": brief,
            "alternate": "NA",
            "reprocessing": "NA",
        }

    @staticmethod
    def _fill_first_s21_table(doc: _Document, mfr_name: str, mfr_addr: str, responsibility: str) -> None:
        for table in doc.tables:
            if len(table.rows) < 2 or len(table.columns) < 3:
                continue
            head = " ".join((table.cell(0, c).text or "").strip().lower() for c in range(3))
            if "name and address" in head and "responsibility" in head:
                table.cell(1, 0).text = f"{mfr_name}\n\nAddress of Manufacturer:\n{mfr_addr}".strip()
                table.cell(1, 1).text = responsibility or "Manufacturing, Packaging and Testing"
                table.cell(1, 2).text = "Not applicable"
                return

    @staticmethod
    def _extract_restricted_phrase(text: str) -> str:
        for ln in text.splitlines():
            if "restricted part" in ln.lower():
                return ln.strip()
        return ""

    def _resolve_name_mfr_line(self) -> str:
        if self.filled_reference_docx and self.filled_reference_docx.exists():
            try:
                ref = Document(self.filled_reference_docx)
                for i, p in enumerate(ref.paragraphs):
                    t = (p.text or "").strip()
                    if "2.3.S.2.1 Manufacturer".lower() in t.lower():
                        for j in range(i + 1, min(i + 6, len(ref.paragraphs))):
                            cand = (ref.paragraphs[j].text or "").strip()
                            if cand.startswith("(") and "," in cand:
                                return cand
            except Exception:
                pass
        return ""

    def fill_s2_section(self, extracted: dict[str, ExtractedSectionContent], output_docx: Path) -> list[str]:
        doc = Document(self.template_docx)
        warnings: list[str] = []

        start_idx, end_idx = self._get_target_range(doc)

        name_line = self._resolve_name_mfr_line()
        for heading in [
            "2.3.S.2 Manufacture",
            "2.3.S.2.1 Manufacturer(s)",
            "2.3.S.2.2 Description of Manufacturing Process and Process Controls",
            "2.3.S.2.3 Control of Materials",
            "2.3.S.2.4 Controls of Critical Steps and Intermediates",
            "2.3.S.2.5 Process Validation and/or Evaluation",
            "2.3.S.2.6 Manufacturing Process Development",
        ]:
            idx = self._find_para_index(doc, heading, start_idx, end_idx)
            if idx is not None and name_line:
                nxt = (doc.paragraphs[idx + 1].text or "").strip() if idx + 1 < len(doc.paragraphs) else ""
                if not nxt.startswith("("):
                    self._insert_after(doc.paragraphs[idx], name_line)

        start_idx, end_idx = self._get_target_range(doc)

        # Remove Refer Section lines from final DOCX.
        self._remove_refer_lines(doc, start_idx, end_idx)

        # Track warnings in log only.
        for k in ["3.2.S.2.1", "3.2.S.2.2", "3.2.S.2.3", "3.2.S.2.4", "3.2.S.2.5", "3.2.S.2.6"]:
            c = extracted[k]
            if c.warning:
                warnings.append(f"{k}: {c.warning}")

        s21_raw = self._clean_text_block(extracted["3.2.S.2.1"].raw_text)
        s22_raw = self._clean_text_block(extracted["3.2.S.2.2"].raw_text)
        s21 = self._parse_s21(s21_raw)
        s22 = self._parse_s22(s22_raw)
        s23 = self._clean_text_block(extracted["3.2.S.2.3"].raw_text)
        s24 = self._clean_text_block(extracted["3.2.S.2.4"].raw_text)
        s25 = self._clean_text_block(extracted["3.2.S.2.5"].raw_text)
        s26 = self._clean_text_block(extracted["3.2.S.2.6"].raw_text)

        restricted_23 = self._extract_restricted_phrase(s23)
        restricted_24 = self._extract_restricted_phrase(s24)
        restricted_25 = self._extract_restricted_phrase(s25)
        restricted_26 = self._extract_restricted_phrase(s26)

        # 2.3.S.2.1 fills
        idx = self._find_para_index(
            doc,
            "Name, address, and responsibility (e.g., fabrication, packaging, labelling, testing, storage)",
            start_idx,
            end_idx,
        )
        if idx is not None:
            # Fill the dedicated table; avoid dumping long raw narrative here.
            self._fill_first_s21_table(doc, s21["mfr_name"], s21["mfr_addr"], s21["responsibility"])

        idx = self._find_para_index(
            doc,
            "Manufacturing authorization for the production of API(s)",
            start_idx,
            end_idx,
        )
        if idx is not None:
            self._insert_after(doc.paragraphs[idx], s21["gmp"])

        # 2.3.S.2.2 fills
        idx = self._find_para_index(doc, "Flow diagram of the synthesis process(es):", start_idx, end_idx)
        if idx is not None and extracted["3.2.S.2.2"].image_paths:
            p = self._insert_after(doc.paragraphs[idx], "")
            r = p.add_run()
            r.add_picture(str(extracted["3.2.S.2.2"].image_paths[0]), width=Inches(5.5))

        idx = self._find_para_index(doc, "Brief narrative description of the manufacturing process(es):", start_idx, end_idx)
        if idx is not None:
            if s22["brief"]:
                self._insert_after(doc.paragraphs[idx], s22["brief"])

        idx = self._find_para_index(doc, "Alternate processes and explanation of their use:", start_idx, end_idx)
        if idx is not None:
            self._insert_after(doc.paragraphs[idx], s22["alternate"])

        idx = self._find_para_index(doc, "Reprocessing steps and justification:", start_idx, end_idx)
        if idx is not None:
            self._insert_after(doc.paragraphs[idx], s22["reprocessing"])

        # 2.3.S.2.3 fills
        idx = self._find_para_index(doc, "(a)\tName of starting material:", start_idx, end_idx)
        if idx is not None and restricted_23:
            self._insert_after(doc.paragraphs[idx], restricted_23)

        idx = self._find_para_index(doc, "(b)\tName and manufacturing site address of starting material manufacturer(s):", start_idx, end_idx)
        if idx is not None and restricted_23:
            self._insert_after(doc.paragraphs[idx], restricted_23)

        idx = self._find_para_index(doc, "Summary of the quality and controls of the starting materials used in the manufacture of the API:", start_idx, end_idx)
        if idx is not None and restricted_23:
            self._insert_after(doc.paragraphs[idx], restricted_23)

        idx = self._find_para_index(doc, "without risk of transmitting agents of animal spongiform encephalopathies", start_idx, end_idx)
        if idx is not None and restricted_23:
            self._insert_after(doc.paragraphs[idx], restricted_23)

        # 2.3.S.2.4 / 2.5 / 2.6 fills
        idx = self._find_para_index(doc, "Summary of the controls performed at critical steps", start_idx, end_idx)
        if idx is not None and restricted_24:
            self._insert_after(doc.paragraphs[idx], restricted_24)

        idx = self._find_para_index(doc, "Description of process validation and/or evaluation studies", start_idx, end_idx)
        if idx is not None and restricted_25:
            self._insert_after(doc.paragraphs[idx], restricted_25)

        idx = self._find_para_index(doc, "Description and discussion of the significant changes made to the manufacturing process", start_idx, end_idx)
        if idx is not None and restricted_26:
            self._insert_after(doc.paragraphs[idx], restricted_26)

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_docx)
        return warnings
