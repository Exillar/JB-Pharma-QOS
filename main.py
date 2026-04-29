from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from config_loader import AppConfig, ConfigLoader, PipelineConfig
from docx_builder import DocxFiller, S2DocxFiller
from pdf_extractor import ExtractedSectionContent, PdfSectionExtractor
from section_mapper import SectionMapper
from verifier import SectionVerifier


@dataclass(frozen=True)
class SectionSpec:
    refer_sections: list[str]
    filler_factory: Callable[[PipelineConfig], object]
    fill_method: str
    start_label: str
    end_label: str
    log_title: str


class QosPipeline:
    SECTION_SPECS: dict[str, SectionSpec] = {
        "s1": SectionSpec(
            refer_sections=["3.2.S.1.1", "3.2.S.1.2", "3.2.S.1.3"],
            filler_factory=lambda cfg: DocxFiller(cfg.template_docx, cfg.filled_reference_docx),
            fill_method="fill_s1_section",
            start_label="2.3.S.1 General Information",
            end_label="2.3.S.2 Manufacture",
            log_title="QOS 2.3.S.1 generation warnings",
        ),
        "s2": SectionSpec(
            refer_sections=[
                "3.2.S.2.1",
                "3.2.S.2.2",
                "3.2.S.2.3",
                "3.2.S.2.4",
                "3.2.S.2.5",
                "3.2.S.2.6",
            ],
            filler_factory=lambda cfg: S2DocxFiller(cfg.template_docx, cfg.filled_reference_docx),
            fill_method="fill_s2_section",
            start_label="2.3.S.2 Manufacture",
            end_label="2.3.S.3 Characterisation",
            log_title="QOS 2.3.S.2 generation warnings",
        ),
    }

    def __init__(self, config: PipelineConfig, section: str) -> None:
        self.config = config
        self.section = section.lower().strip()

    def run(self) -> "PipelineResult":
        spec = self.SECTION_SPECS.get(self.section)
        if not spec:
            raise ValueError(f"Unsupported section: {self.section}")

        self.config.ensure_directories()

        mapper = SectionMapper(self.config.module3_root)
        if self.config.extractor_backend == "pymupdf":
            extractor = PdfSectionExtractor(self.config.image_artifacts_dir)
        else:
            raise ValueError(f"Unsupported extractor backend: {self.config.extractor_backend}")
        filler = spec.filler_factory(self.config)
        verifier = SectionVerifier()

        extracted_payload: dict[str, ExtractedSectionContent] = {}
        warnings: list[str] = []

        for refer in spec.refer_sections:
            resolved = mapper.resolve(refer)
            content = extractor.extract(resolved)
            extracted_payload[refer] = content
            if content.warning:
                warnings.append(f"{refer}: {content.warning}")

        output_docx = self.config.output_docx
        try:
            fill = getattr(filler, spec.fill_method)
            docx_warnings = fill(extracted_payload, output_docx)
        except PermissionError:
            from datetime import datetime

            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            alt = output_docx.with_name(f"{output_docx.stem}_{stamp}{output_docx.suffix}")
            fill = getattr(filler, spec.fill_method)
            docx_warnings = fill(extracted_payload, alt)
            output_docx = alt
        warnings.extend(docx_warnings)

        verifier.verify(
            generated_docx=output_docx,
            filled_reference_docx=self.config.filled_reference_docx,
            report_path=self.config.verification_report_path,
            start_label=spec.start_label,
            end_label=spec.end_label,
        )

        unique_warnings = sorted(set(warnings))
        log_lines = [
            spec.log_title,
            "===============================",
        ]
        if unique_warnings:
            log_lines.extend(unique_warnings)
        else:
            log_lines.append("No warnings")
        self.config.generation_log_path.write_text("\n".join(log_lines), encoding="utf-8")

        return PipelineResult(
            output_docx=output_docx,
            verification_report=self.config.verification_report_path,
            warnings=unique_warnings,
        )


@dataclass(frozen=True)
class PipelineResult:
    output_docx: Path
    verification_report: Path
    warnings: list[str]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate QOS sections from dossier PDFs")
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to YAML config (defaults to ./config.yaml)",
    )
    parser.add_argument("--template", default="", help="Path to Quality Overall Summary template DOCX")
    parser.add_argument("--dossier-root", default="", help="Path to dossier root folder (e.g., Cardiolek)")
    parser.add_argument("--filled-reference", default="", help="Path to already-filled QOS DOCX for verification")
    parser.add_argument("--output", default="", help="Output DOCX path")
    parser.add_argument(
        "--artifacts-dir",
        default=str(Path(__file__).resolve().parent / "artifacts"),
        help="Artifacts folder for images and reports",
    )
    parser.add_argument(
        "--backend",
        default="pymupdf",
        choices=["pymupdf"],
        help="Extraction backend",
    )
    parser.add_argument(
        "--section",
        default="s1",
        choices=["s1", "s2"],
        help="QOS section automation target",
    )
    return parser


def _apply_overrides(base: AppConfig, args: argparse.Namespace) -> AppConfig:
    return base.with_overrides(
        template_docx=args.template or None,
        dossier_root=args.dossier_root or None,
        filled_reference_docx=args.filled_reference or None,
        output_docx=args.output or None,
        artifacts_dir=args.artifacts_dir or None,
        extractor_backend=args.backend or None,
        section=args.section or None,
    )


def main() -> None:
    args = build_parser().parse_args()

    loader = ConfigLoader(Path(args.config))
    app_config = _apply_overrides(loader.load(), args)

    config = PipelineConfig(
        template_docx=app_config.template_docx,
        dossier_root=app_config.dossier_root,
        filled_reference_docx=app_config.filled_reference_docx,
        output_docx=app_config.output_docx,
        artifacts_dir=app_config.artifacts_dir,
        extractor_backend=app_config.extractor_backend,
    )

    pipeline = QosPipeline(config, app_config.section)
    result = pipeline.run()

    print(f"=== QOS {app_config.section.upper()} Generation Summary ===")
    print(f"Output DOCX: {result.output_docx}")
    print(f"Verification report: {result.verification_report}")
    print("Warnings:")
    for warning in result.warnings:
        print(f"- {warning}")


if __name__ == "__main__":
    main()
