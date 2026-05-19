from __future__ import annotations

import argparse
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from builders import (
    S1DocxFiller,
    S2DocxFiller,
    S3DocxFiller,
    S32DocxFiller,
    S4DocxFiller,
    S5DocxFiller,
    GenericSectionFiller,
    P1DocxFiller,
)
from config_loader import AppConfig, ConfigLoader, PipelineConfig
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
            filler_factory=lambda cfg: S1DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
            ),
            fill_runner=lambda f, p, o: f.fill_s1_section(p, o),
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
            fill_runner=lambda f, p, o: f.fill_s2_section(p, o),
            start_label="2.3.S.2 Manufacture",
            end_label="2.3.S.3 Characterisation",
            log_title="QOS 2.3.S.2 generation warnings",
        ),
        "s3": SectionSpec(
            refer_sections=["3.2.S.3.1"],
            filler_factory=lambda cfg: S3DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
                s3_cfg=cfg.s3_fill,
            ),
            fill_runner=lambda f, p, o: f.fill_s3_section(p, o),
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
                preserve_repeated_patterns=cfg.s2_fill.restricted_phrase_keywords,
            ),
            fill_runner=lambda f, p, o: f.fill_s32_section(p, o),
            start_label="2.3.S.3.2 Impurities",
            end_label="2.3.S.4 Control",
            log_title="QOS 2.3.S.3.2 generation warnings",
        ),
        "s4": SectionSpec(
            refer_sections=["3.2.S.4.1", "3.2.S.4.2", "3.2.S.4.5"],
            filler_factory=lambda cfg: S4DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
                preserve_repeated_patterns=cfg.s2_fill.restricted_phrase_keywords,
            ),
            fill_runner=lambda f, p, o: f.fill_s4_section(p, o),
            start_label="2.3.S.4 Control of the API",
            end_label="2.3.S.5 Reference Standards or Materials",
            log_title="QOS 2.3.S.4 generation warnings",
        ),
        "s5": SectionSpec(
            refer_sections=["3.2.S.5"],
            filler_factory=lambda cfg: S5DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
                images_dir=cfg.image_artifacts_dir,
                diagram_cfg=cfg.diagram,
                preserve_repeated_patterns=cfg.s2_fill.restricted_phrase_keywords,
            ),
            fill_runner=lambda f, p, o: f.fill_s5_section(p, o),
            start_label="2.3.S.5 Reference Standards or Materials",
            end_label="2.3.S.6 Container Closure System",
            log_title="QOS 2.3.S.5 generation warnings",
        ),
        "s6": SectionSpec(
            refer_sections=["3.2.S.6"],
            filler_factory=lambda cfg: GenericSectionFiller(
                cfg.template_docx,
                "2.3.S.6 Container Closure System",
                "2.3.S.7 Stability",
                cfg.filled_reference_docx,
                preserve_repeated_patterns=cfg.s2_fill.restricted_phrase_keywords,
            ),
            fill_runner=lambda f, p, o: f.fill_section(p, o),
            start_label="2.3.S.6 Container Closure System",
            end_label="2.3.S.7 Stability",
            log_title="QOS 2.3.S.6 generation warnings",
        ),
        "s7": SectionSpec(
            refer_sections=["3.2.S.7.1", "3.2.S.7.2", "3.2.S.7.3"],
            filler_factory=lambda cfg: GenericSectionFiller(
                cfg.template_docx,
                "2.3.S.7 Stability",
                "2.3.P",
                cfg.filled_reference_docx,
                preserve_repeated_patterns=cfg.s2_fill.restricted_phrase_keywords,
            ),
            fill_runner=lambda f, p, o: f.fill_section(p, o),
            start_label="2.3.S.7 Stability",
            end_label="2.3.P",
            log_title="QOS 2.3.S.7 generation warnings",
        ),
        "p1": SectionSpec(
            refer_sections=["3.2.P.1"],
            filler_factory=lambda cfg: P1DocxFiller(
                cfg.template_docx,
                cfg.filled_reference_docx,
                preserve_repeated_patterns=cfg.s2_fill.restricted_phrase_keywords,
            ),
            fill_runner=lambda f, p, o: f.fill_p1_section(p, o),
            start_label="2.3.P.1 Description and Composition of the FPP",
            end_label="2.3.P.2 Pharmaceutical Development",
            log_title="QOS 2.3.P.1 generation warnings",
        ),
    }

    # Ordered pipeline for --section=all.
    # Each entry is (section_key, output_stem).
    # The output of step N becomes the template input for step N+1.
    ALL_SECTIONS: list[str] = ["s1", "s2", "s3", "s32", "s4", "s5", "s6", "s7", "p1"]

    def __init__(self, config: PipelineConfig, section: str) -> None:
        self.config = config
        self.section = section.lower().strip()

    @staticmethod
    def _extract_product_name(text: str) -> str:
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        for i, ln in enumerate(lines):
            low = ln.lower()
            if "compendial name" in low:
                if ":" in ln:
                    tail = ln.split(":", 1)[1].strip()
                    if tail:
                        return tail
                if i + 1 < len(lines):
                    return lines[i + 1].strip()
        for i, ln in enumerate(lines):
            low = ln.lower()
            if (
                "recommended international nonproprietary name" in low
                or "inn" in low
            ):
                if ":" in ln:
                    tail = ln.split(":", 1)[1].strip()
                    if tail:
                        return tail
                if i + 1 < len(lines):
                    return lines[i + 1].strip()
        return ""

    @staticmethod
    def _extract_manufacturer_name(text: str) -> str:
        if not text:
            return ""
        m = re.search(r"\bby\s+([^\.;]{5,160})", text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        for ln in lines:
            low = ln.lower()
            if "pharmaceutical" in low and ("co" in low or "ltd" in low):
                return ln
        return ""

    def _derive_name_mfr_line(
        self,
        mapper: SectionMapper,
        extractor: PdfSectionExtractor,
    ) -> str:
        try:
            s11 = extractor.extract(mapper.resolve("3.2.S.1.1")).raw_text
            s21 = extractor.extract(mapper.resolve("3.2.S.2.1")).raw_text
        except Exception:
            return ""

        def _clean(val: str) -> str:
            if not val:
                return ""
            compact = re.sub(r"\s+", " ", val).strip()
            compact = re.sub(r"\s+,", ",", compact)
            compact = re.sub(r"[\s\.]+$", "", compact)
            return compact

        product = _clean(self._extract_product_name(s11))
        manufacturer = _clean(self._extract_manufacturer_name(s21))

        if product and manufacturer:
            return f"({product}, {manufacturer})"
        if product:
            return f"({product})"
        if manufacturer:
            return f"({manufacturer})"
        return ""

    def run(self) -> "PipelineResult":
        spec = self.SECTION_SPECS.get(self.section)
        if not spec:
            raise ValueError(f"Unsupported section: {self.section!r}")

        self.config.ensure_directories()

        mapper = SectionMapper(self.config.module3_root)
        if self.config.extractor_backend != "pymupdf":
            raise ValueError(f"Unsupported extractor backend: {self.config.extractor_backend!r}")
        extractor = PdfSectionExtractor(
            self.config.image_artifacts_dir,
            diagram_cfg=self.config.diagram,
            noise_cfg=self.config.noise,
            preserve_keywords=self.config.s2_fill.restricted_phrase_keywords,
        )

        name_mfr_line = self._derive_name_mfr_line(mapper, extractor)
        filler = spec.filler_factory(self.config)
        if name_mfr_line:
            setattr(filler, "_name_mfr_line", name_mfr_line)
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
        log_lines = [spec.log_title, "=" * 40]
        log_lines += unique_warnings if unique_warnings else ["No warnings"]
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
    valid_sections = ["s1", "s2", "s3", "s32", "s4", "s5", "s6", "s7", "p1", "all"]
    parser = argparse.ArgumentParser(description="Generate QOS sections from dossier PDFs")
    parser.add_argument("--config", default="config.yaml", help="Path to YAML config")
    parser.add_argument("--template", default=None, help="QOS template DOCX path")
    parser.add_argument("--dossier-root", default=None, help="Dossier root folder path")
    parser.add_argument("--filled-reference", default=None, help="Filled QOS DOCX for verification")
    parser.add_argument("--output", default=None, help="Output DOCX path")
    parser.add_argument("--artifacts-dir", default="", help="Folder for images and reports")
    parser.add_argument("--backend", default="", choices=["", "pymupdf"], help="Extraction backend")
    parser.add_argument(
        "--section", default="", choices=[""] + valid_sections, help="QOS section to generate"
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
        try:
            app_config.artifacts_dir.mkdir(parents=True, exist_ok=True)
            (app_config.artifacts_dir / "images").mkdir(parents=True, exist_ok=True)
            return app_config.artifacts_dir
        except Exception:
            fallback = Path.cwd() / "artifacts" / f"run_{run_stamp}"
            (fallback / "images").mkdir(parents=True, exist_ok=True)
            return fallback

    def resolve_final_output_path() -> Path:
        preferred_dir = app_config.dossier_root / "Output"
        try:
            preferred_dir.mkdir(parents=True, exist_ok=True)
            return preferred_dir / app_config.output_docx.name
        except Exception:
            return app_config.output_docx

    def cleanup_images_dir() -> None:
        shutil.rmtree(resolve_writable_artifacts_dir() / "images", ignore_errors=True)

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
            s3_fill=app_config.s3_fill,
            verification_report_name=verification_report_name,
            generation_log_name=generation_log_name,
        )
        return QosPipeline(config, section).run()

    if app_config.section == "all":
        intermediates_dir = resolve_writable_artifacts_dir()
        try:
            candidate = intermediates_dir / "intermediates"
            candidate.mkdir(parents=True, exist_ok=True)
            intermediates_dir = candidate
        except Exception:
            pass

        sections_to_run = [
            s for s in QosPipeline.ALL_SECTIONS if s in QosPipeline.SECTION_SPECS
        ]
        final_output = resolve_final_output_path()

        results: list[PipelineResult] = []
        current_template = app_config.template_docx

        for i, section in enumerate(sections_to_run):
            is_last = i == len(sections_to_run) - 1
            letter = chr(ord("a") + i)
            output_path = (
                final_output
                if is_last
                else intermediates_dir / f"qos_{section}_intermediate.docx"
            )
            result = run_section(
                section,
                current_template,
                output_path,
                verification_report_name=f"verification_report_part_{letter}.txt",
                generation_log_name=f"generation_part_{letter}.log",
            )
            results.append(result)
            current_template = output_path

        cleanup_images_dir()

        final_path = results[-1].output_docx
        desired_output = app_config.dossier_root / "Output" / app_config.output_docx.name
        try:
            desired_output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(final_path, desired_output)
            final_path = desired_output
        except Exception:
            pass

        print("=== QOS ALL Generation Summary ===")
        print(f"Final Output DOCX: {final_path}")
        for section, result in zip(sections_to_run, results):
            print(f"  [{section.upper()}] Verification: {result.verification_report}")
        all_warnings = [w for r in results for w in r.warnings]
        if all_warnings:
            print("Warnings:")
            for warning in all_warnings:
                print(f"  - {warning}")
    else:
        result = run_section(
            app_config.section, app_config.template_docx, app_config.output_docx
        )
        cleanup_images_dir()

        print(f"=== QOS {app_config.section.upper()} Generation Summary ===")
        print(f"Output DOCX: {result.output_docx}")
        print(f"Verification report: {result.verification_report}")
        if result.warnings:
            print("Warnings:")
            for warning in result.warnings:
                print(f"  - {warning}")


if __name__ == "__main__":
    main()
