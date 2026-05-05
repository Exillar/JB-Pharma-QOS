from __future__ import annotations

import re
from pathlib import Path

from docx import Document


class SectionVerifier:
    DEFAULT_START = "2.3.S.1 General Information"
    DEFAULT_END = "2.3.S.2 Manufacture"

    @staticmethod
    def _extract_section_text(doc_path: Path, start_label: str, end_label: str) -> str:
        doc = Document(doc_path)
        lines: list[str] = []
        in_section = False

        for p in doc.paragraphs:
            text = (p.text or "").strip()
            if not in_section and start_label.lower() in text.lower():
                in_section = True
            if in_section and end_label.lower() in text.lower():
                break
            if in_section:
                lines.append(text)

        return "\n".join(lines)

    @staticmethod
    def _tokenize(text: str) -> set[str]:
        return set(re.findall(r"[A-Za-z0-9\.\-/]+", text.lower()))

    def verify(
        self,
        generated_docx: Path,
        filled_reference_docx: Path,
        report_path: Path,
        start_label: str = DEFAULT_START,
        end_label: str = DEFAULT_END,
    ) -> None:
        gen = self._extract_section_text(generated_docx, start_label, end_label)
        ref = self._extract_section_text(filled_reference_docx, start_label, end_label)

        gen_tokens = self._tokenize(gen)
        ref_tokens = self._tokenize(ref)

        overlap = len(gen_tokens & ref_tokens)
        denom = max(1, len(ref_tokens))
        ratio = overlap / denom

        report_lines = [
            f"QOS Verification Report ({start_label})",
            "===============================",
            f"Generated file: {generated_docx}",
            f"Reference file: {filled_reference_docx}",
            "",
            f"Reference token count: {len(ref_tokens)}",
            f"Generated token count: {len(gen_tokens)}",
            f"Token overlap count: {overlap}",
            f"Coverage vs reference: {ratio:.2%}",
            "",
            "Note: This is a lexical overlap check for quick verification only.",
            "Manual QA is still required for regulatory submission quality.",
        ]

        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text("\n".join(report_lines), encoding="utf-8")
