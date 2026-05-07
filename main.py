from __future__ import annotations

import argparse
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from config_loader import AppConfig, ConfigLoader, PipelineConfig
from docx_builder import DocxFiller, S2DocxFiller, S3DocxFiller, S32DocxFiller, S4DocxFiller
from pdf_extractor import ExtractedSectionContent, PdfSectionExtractor
from section_mapper import SectionMapper
from verifier import SectionVerifier


@dataclass(frozen=True)
class SectionSpec:
    refer_sections: list[str]
    filler_factory: Callable[[PipelineConfig], object]
    fill_runner: Callable[[object, dict[str, ExtractedSectionContent], Path], list[str]]
    start_label: str
    end_label: str
    log_title: str


class QosPipeline:
    SECTION_SPECS: dict[str, SectionSpec] = {
        "s1": SectionSpec(
            refer_sections=["3.2.S.1.1", "3.2.S.1.2", "3.2.S.1.3"],
            filler_factory=lambda cfg: DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
            ),
            fill_runner=lambda filler, payload, out: filler.fill_s1_section(payload, out),
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
            filler_factory=lambda cfg: S2DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
                images_dir=cfg.image_artifacts_dir,
                s2_fill_cfg=cfg.s2_fill,
                noise_cfg=cfg.noise,
                diagram_cfg=cfg.diagram,
            ),
            fill_runner=lambda filler, payload, out: filler.fill_s2_section(payload, out),
            start_label="2.3.S.2 Manufacture",
            end_label="2.3.S.3 Characterisation",
            log_title="QOS 2.3.S.2 generation warnings",
        ),
        "s3": SectionSpec(
            refer_sections=["3.2.S.3.1"],
            filler_factory=lambda cfg: S3DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
                s2_fill_cfg=cfg.s2_fill,
            ),
            fill_runner=lambda filler, payload, out: filler.fill_s3_section(payload, out),
            start_label="2.3.S.3 Characterisation",
            end_label="2.3.S.4 Control of Drug Substance",
            log_title="QOS 2.3.S.3 generation warnings",
        ),
        "s32": SectionSpec(
            refer_sections=["3.2.S.3.2"],
            filler_factory=lambda cfg: S32DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
                images_dir=cfg.image_artifacts_dir,
                diagram_cfg=cfg.diagram,
            ),
            fill_runner=lambda filler, payload, out: filler.fill_s32_section(payload, out),
            start_label="2.3.S.3.2 Impurities",
            end_label="2.3.S.4 Control",
            log_title="QOS 2.3.S.3.2 generation warnings",
        ),
        "s4": SectionSpec(
            refer_sections=["3.2.S.4.1", "3.2.S.4.2"],
            filler_factory=lambda cfg: S4DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
            ),
            fill_runner=lambda filler, payload, out: filler.fill_s4_section(payload, out),
            start_label="2.3.S.4 Control of the API",
            end_label="2.3.S.5 Reference Standards or Materials",
            log_title="QOS 2.3.S.4 generation warnings",
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
            extractor = PdfSectionExtractor(
                self.config.image_artifacts_dir,
                diagram_cfg=self.config.diagram,
                noise_cfg=self.config.noise,
            )
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
            docx_warnings = spec.fill_runner(filler, extracted_payload, output_docx)
        except (PermissionError, FileNotFoundError, OSError):
            from datetime import datetime

            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_dir = self.config.artifacts_dir / "fallback_outputs"
            safe_dir.mkdir(parents=True, exist_ok=True)
            alt = safe_dir / f"{output_docx.stem}_{stamp}{output_docx.suffix}"
            docx_warnings = spec.fill_runner(filler, extracted_payload, alt)
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
    parser.add_argument("--template", default=None, help="Path to Quality Overall Summary template DOCX")
    parser.add_argument("--dossier-root", default=None, help="Path to dossier root folder")
    parser.add_argument("--filled-reference", default=None, help="Path to already-filled QOS DOCX for verification")
    parser.add_argument("--output", default=None, help="Output DOCX path")
    parser.add_argument(
        "--artifacts-dir",
        default="",
        help="Artifacts folder for images and reports",
    )
    parser.add_argument(
        "--backend",
        default="",
        choices=["", "pymupdf"],
        help="Extraction backend",
    )
    parser.add_argument(
        "--section",
        default="",
        choices=["", "s1", "s2", "s3", "s32", "s4", "all"],
        help="QOS section automation target",
    )
    return parser


def _apply_overrides(base: AppConfig, args: argparse.Namespace) -> AppConfig:
    return base.with_overrides(
        template_docx=args.template,
        dossier_root=args.dossier_root,
        filled_reference_docx=args.filled_reference,
        output_docx=args.output,
        artifacts_dir=args.artifacts_dir,
        extractor_backend=args.backend,
        section=args.section,
    )


def main() -> None:
    args = build_parser().parse_args()

    loader = ConfigLoader(Path(args.config))
    app_config = _apply_overrides(loader.load(), args)
    from datetime import datetime
    run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    def resolve_writable_artifacts_dir() -> Path:
        """
        artifacts_dir in config may point to a protected folder (common in corporate paths).
        Fall back to a local workspace folder when we cannot create subfolders.
        """
        try:
            app_config.artifacts_dir.mkdir(parents=True, exist_ok=True)
            test_dir = app_config.artifacts_dir / "images"
            test_dir.mkdir(parents=True, exist_ok=True)
            return app_config.artifacts_dir
        except Exception:
            fallback = Path.cwd() / "artifacts" / f"run_{run_stamp}"
            fallback.mkdir(parents=True, exist_ok=True)
            (fallback / "images").mkdir(parents=True, exist_ok=True)
            return fallback

    def resolve_final_output_path() -> Path:
        # Prefer a conventional dossier_root/Output folder when present (or creatable),
        # so "all" runs produce a single deliverable in one place.
        preferred_dir = app_config.dossier_root / "Output"
        try:
            preferred_dir.mkdir(parents=True, exist_ok=True)
            return preferred_dir / app_config.output_docx.name
        except Exception:
            return app_config.output_docx

    def cleanup_images_dir() -> None:
        shutil.rmtree(app_config.artifacts_dir / "images", ignore_errors=True)

    def run_section(
        section: str,
        template_docx: Path,
        output_docx: Path,
        *,
        verification_report_name: str = "verification_report.txt",
        generation_log_name: str = "generation.log",
    ) -> PipelineResult:
        effective_artifacts_dir = resolve_writable_artifacts_dir()
        config = PipelineConfig(
            template_docx=template_docx,
            dossier_root=app_config.dossier_root,
            filled_reference_docx=app_config.filled_reference_docx,
            output_docx=output_docx,
            artifacts_dir=effective_artifacts_dir,
            extractor_backend=app_config.extractor_backend,
            diagram=app_config.diagram,
            noise=app_config.noise,
            s2_fill=app_config.s2_fill,
            verification_report_name=verification_report_name,
            generation_log_name=generation_log_name,
        )
        pipeline = QosPipeline(config, section)
        return pipeline.run()

    if app_config.section == "all":
        intermediates_dir = app_config.artifacts_dir
        try:
            candidate = app_config.artifacts_dir / "intermediates"
            candidate.mkdir(parents=True, exist_ok=True)
            intermediates_dir = candidate
        except Exception:
            # Some dossier folders allow writing files but block new subfolders.
            intermediates_dir = app_config.artifacts_dir
        intermediate_s1 = intermediates_dir / "qos_s1_intermediate.docx"
        intermediate_s2 = intermediates_dir / "qos_s2_intermediate.docx"
        intermediate_s3 = intermediates_dir / "qos_s3_intermediate.docx"
        intermediate_s32 = intermediates_dir / "qos_s32_intermediate.docx"
        final_output = resolve_final_output_path()

        result_s1 = run_section(
            "s1",
            app_config.template_docx,
            intermediate_s1,
            verification_report_name="verification_report_part_a.txt",
            generation_log_name="generation_part_a.log",
        )
        result_s2 = run_section(
            "s2",
            intermediate_s1,
            intermediate_s2,
            verification_report_name="verification_report_part_b.txt",
            generation_log_name="generation_part_b.log",
        )
        result_s3 = run_section(
            "s3",
            intermediate_s2,
            intermediate_s3,
            verification_report_name="verification_report_part_c.txt",
            generation_log_name="generation_part_c.log",
        )
        result_s32 = run_section(
            "s32",
            intermediate_s3,
            intermediate_s32,
            verification_report_name="verification_report_part_d.txt",
            generation_log_name="generation_part_d.log",
        )
        result_s4 = run_section(
            "s4",
            intermediate_s32,
            final_output,
            verification_report_name="verification_report_part_e.txt",
            generation_log_name="generation_part_e.log",
        )

        cleanup_images_dir()

        # Ensure a single deliverable lands in dossier_root/Output when possible.
        desired_output = app_config.dossier_root / "Output" / app_config.output_docx.name
        final_path = result_s4.output_docx
        try:
            desired_output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(final_path, desired_output)
            final_path = desired_output
        except Exception:
            pass

        # Keep only the final deliverable in Output; intermediates stay under artifacts_dir (or artifacts_dir/intermediates when possible).
        # If you want intermediates removed entirely, uncomment this block.
        # try:
        #     shutil.rmtree(intermediates_dir, ignore_errors=True)
        # except Exception:
        #     pass

        print("=== QOS ALL Generation Summary ===")
        print(f"Final Output DOCX: {final_path}")
        print(f"Verification report (part A): {result_s1.verification_report}")
        print(f"Verification report (part B): {result_s2.verification_report}")
        print(f"Verification report (part C): {result_s3.verification_report}")
        print(f"Verification report (part D): {result_s32.verification_report}")
        print(f"Verification report (part E): {result_s4.verification_report}")
        if result_s1.warnings or result_s2.warnings or result_s3.warnings or result_s32.warnings or result_s4.warnings:
            print("Warnings:")
            for warning in (
                result_s1.warnings
                + result_s2.warnings
                + result_s3.warnings
                + result_s32.warnings
                + result_s4.warnings
            ):
                print(f"- {warning}")
    else:
        result = run_section(app_config.section, app_config.template_docx, app_config.output_docx)

        cleanup_images_dir()

        print(f"=== QOS {app_config.section.upper()} Generation Summary ===")
        print(f"Output DOCX: {result.output_docx}")
        print(f"Verification report: {result.verification_report}")
        print("Warnings:")
        for warning in result.warnings:
            print(f"- {warning}")


if __name__ == "__main__":
    main()
