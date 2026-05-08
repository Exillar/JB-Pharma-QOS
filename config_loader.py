from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


# ---------------------------------------------------------------------------
# Diagram extraction tuning
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class DiagramConfig:
    """Controls how vector-drawn flow-diagram pages are detected and rendered.

    Defaults are generic and should work for common CTD-style dossiers.
    Override in config.yaml under the ``diagram:`` key when a dossier has
    different header/footer heights or drawing densities.
    """

    # Sections whose diagrams are purely vector-drawn (no embedded XREFs).
    # Every page in these sections that passes the heuristic is rendered.
    vector_diagram_sections: tuple[str, ...] = ("3.2.S.2.2",)

    # Sections where the page itself is rendered as a fallback when no
    # embedded image XREF is found (e.g. structural-formula pages).
    page_render_sections: tuple[str, ...] = ("3.2.S.1.2",)

    # Minimum number of vector drawing paths required before a page is even
    # considered as a diagram candidate.
    min_diagram_drawings: int = 40

    # If chars_per_drawing exceeds this, the page is mostly text (table with
    # thin border lines counts as many "drawings" but is not a diagram).
    chars_per_drawing_threshold: float = 30.0

    # PyMuPDF Matrix scale factor applied when rendering a page to PNG.
    # 2.0 → ~144 DPI at 72 DPI base; increase for higher resolution output.
    render_dpi_scale: float = 2.0

    # Fraction of page height to remove from the top (covers the dossier
    # header table rows: company name, drug name, document title).
    header_crop_frac: float = 0.135

    # Fraction of page height to remove from the bottom (covers the footer:
    # company line + "N of M" page number).
    footer_crop_frac: float = 0.09

    # Generic flow-diagram keywords used by the heuristic filter.  These are
    # process-neutral terms; do NOT add compound-specific names here.
    diagram_keywords: tuple[str, ...] = (
        "flow diagram",
        "stage",
        "figure",
        "filtration",
        "crystallization",
        "distillation",
        "recrystallization",
        "reduction",
        "neutralization",
        "polymerization",
        "desalting",
        "decoloration",
    )

    # Keywords that signal non-flow-diagram pages (e.g., chemical pathway pages).
    diagram_exclude_keywords: tuple[str, ...] = (
        "pathway",
        "synthetic",
        "synthetical",
        "reaction scheme",
    )


# ---------------------------------------------------------------------------
# Noise / header-footer removal
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class NoiseConfig:
    """Strings that appear as running headers/footers in dossier PDFs.

    Add entries when a new dossier uses different company names or page-number
    formats so the extractor can strip them automatically.
    """

    # Line-prefix strings that should always be suppressed (case-insensitive).
    company_name_prefixes: tuple[str, ...] = ()

    # Minimum number of pages a string must appear on to be auto-detected as
    # header/footer noise (mirrors the QIS _build_noise_blocklist threshold).
    noise_page_threshold: int = 3

    # Fraction of page height defining the "top margin" zone for auto-detection.
    noise_top_margin_frac: float = 0.12

    # Fraction of page height defining the "bottom margin" zone.
    noise_bottom_margin_frac: float = 0.10


# ---------------------------------------------------------------------------
# S2 fill defaults
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class S2FillConfig:
    """Default answers used when parsed values cannot be extracted from the PDF.

    All of these are last-resort fallbacks with an associated warning logged.
    If the source PDF contains the information, the parsed value takes priority.
    """

    # 2.3.S.2.2 fixed answers (pharmaceutical manufacturing boilerplate)
    alternate_processes_default: str = "NA"
    reprocessing_steps_default: str = "NA"

    # 2.3.S.2.1 manufacturer table
    manufacturer_table_responsibility_default: str = "Manufacturing, Packaging and Testing"
    manufacturer_table_apprx_col: str = "Not applicable"

    # GMP fallback sentences (used when no GMP phrase found in source text)
    gmp_found_sentence: str = (
        "GMP Certificate of API manufacturer is enclosed under section 3.2.S.2.1 Manufacture (s)."
    )
    gmp_fallback_sentence: str = "GMP information is provided in Module 1."

    # Keywords used to detect the GMP block in source text
    gmp_keywords: tuple[str, ...] = (
        "certificate of gmp compliance",
        "gmp",
    )

    # Generic confidentiality cues seen in S2.3-S2.6 text.
    # First line matching any keyword is reused where section text is redacted.
    restricted_phrase_keywords: tuple[str, ...] = (
        "restricted part",
        "drug master file",
        "confidential",
    )

    # 2.3.S.2.3 inline fallback next to point (b) manufacturer-address label.
    s23_manufacturer_not_available_default: str = "NA"

    # 2.3.S.2.3 first table header normalization.
    s23_table_first_header_default: str = "Step / Starting Material"

    # Keywords used to detect that source text is the main narrative
    # (as opposed to certificate scan noise, table-of-contents lines, etc.)
    narrative_start_keywords: tuple[str, ...] = ("the active drug",)

    # Lines containing these phrases signal the end of the useful narrative
    narrative_end_keywords: tuple[str, ...] = (
        "certificate of gmp compliance",
        "certificate of good manufacturing practices",
    )


# ---------------------------------------------------------------------------
# S3 fill defaults (split out from S2FillConfig for clarity)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class S3Config:
    """Parsing controls and defaults for QOS section 2.3.S.3.1 Characterisation."""

    s31_summary_start_keywords: tuple[str, ...] = (
        "the structural elucidation",
        "elucidation of structure",
    )
    s31_summary_stop_keywords: tuple[str, ...] = (
        "for details of elucidation of structure",
        "3.2.s.3 ",
        "3.2.s.3.1.1",
    )
    s31_max_summary_lines: int = 14
    s31_isomerism_default: str = "NA"
    s31_polymorph_reference_default: str = "Refer to Module 3 Section 3.2.S.3.1 for details."
    s31_particle_size_default: str = "NA"
    s31_other_characteristics_default: str = "NA"


# ---------------------------------------------------------------------------
# Top-level application config
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class AppConfig:
    template_docx: Path
    dossier_root: Path
    filled_reference_docx: Path
    output_docx: Path
    artifacts_dir: Path
    extractor_backend: str
    section: str
    diagram: DiagramConfig = field(default_factory=DiagramConfig)
    noise: NoiseConfig = field(default_factory=NoiseConfig)
    s2_fill: S2FillConfig = field(default_factory=S2FillConfig)
    s3_fill: S3Config = field(default_factory=S3Config)

    def with_overrides(
        self,
        *,
        template_docx: str | None = None,
        dossier_root: str | None = None,
        filled_reference_docx: str | None = None,
        output_docx: str | None = None,
        artifacts_dir: str | None = None,
        extractor_backend: str | None = None,
        section: str | None = None,
    ) -> "AppConfig":
        return AppConfig(
            template_docx=Path(template_docx) if template_docx else self.template_docx,
            dossier_root=Path(dossier_root) if dossier_root else self.dossier_root,
            filled_reference_docx=(
                Path(filled_reference_docx) if filled_reference_docx else self.filled_reference_docx
            ),
            output_docx=Path(output_docx) if output_docx else self.output_docx,
            artifacts_dir=Path(artifacts_dir) if artifacts_dir else self.artifacts_dir,
            extractor_backend=extractor_backend or self.extractor_backend,
            section=section or self.section,
            diagram=self.diagram,
            noise=self.noise,
            s2_fill=self.s2_fill,
            s3_fill=self.s3_fill,
        )


@dataclass(frozen=True)
class PipelineConfig:
    template_docx: Path
    dossier_root: Path
    filled_reference_docx: Path
    output_docx: Path
    artifacts_dir: Path
    extractor_backend: str = "pymupdf"
    diagram: DiagramConfig = field(default_factory=DiagramConfig)
    noise: NoiseConfig = field(default_factory=NoiseConfig)
    s2_fill: S2FillConfig = field(default_factory=S2FillConfig)
    s3_fill: S3Config = field(default_factory=S3Config)
    verification_report_name: str = "verification_report.txt"
    generation_log_name: str = "generation.log"

    @property
    def module3_root(self) -> Path:
        # Some deployments point dossier_root at the dossier root (which contains "Module 3"),
        # while others point dossier_root directly at the "Module 3" folder itself.
        nested = self.dossier_root / "Module 3"
        if nested.exists():
            return nested
        return self.dossier_root

    @property
    def image_artifacts_dir(self) -> Path:
        return self.artifacts_dir / "images"

    @property
    def verification_report_path(self) -> Path:
        return self.artifacts_dir / self.verification_report_name

    @property
    def generation_log_path(self) -> Path:
        return self.artifacts_dir / self.generation_log_name

    def ensure_directories(self) -> None:
        self.artifacts_dir.mkdir(parents=True, exist_ok=True)
        self.image_artifacts_dir.mkdir(parents=True, exist_ok=True)
        try:
            self.output_docx.parent.mkdir(parents=True, exist_ok=True)
        except PermissionError:
            # Output folder may be locked down (e.g., OneDrive / corporate paths).
            pass


# ---------------------------------------------------------------------------
# YAML loader
# ---------------------------------------------------------------------------

class ConfigLoader:
    def __init__(self, config_path: Path) -> None:
        self.config_path = config_path

    def load(self) -> AppConfig:
        if not self.config_path.exists():
            raise FileNotFoundError(f"Config not found: {self.config_path}")

        data = self._load_yaml(self.config_path)
        base_dir = self.config_path.parent

        return AppConfig(
            template_docx=self._resolve_path(data, "template_docx", base_dir),
            dossier_root=self._resolve_path(data, "dossier_root", base_dir),
            filled_reference_docx=self._resolve_path(data, "filled_reference_docx", base_dir),
            output_docx=self._resolve_path(data, "output_docx", base_dir),
            artifacts_dir=self._resolve_path(data, "artifacts_dir", base_dir),
            extractor_backend=str(data.get("extractor_backend", "pymupdf")),
            section=str(data.get("section", "s1")).lower(),
            diagram=self._parse_frozen_dataclass(DiagramConfig, data.get("diagram") or {}),
            noise=self._parse_frozen_dataclass(NoiseConfig, data.get("noise") or {}),
            s2_fill=self._parse_frozen_dataclass(S2FillConfig, data.get("s2_fill") or {}),
            s3_fill=self._parse_frozen_dataclass(S3Config, data.get("s3_fill") or {}),
        )

    @staticmethod
    def _parse_frozen_dataclass(cls, raw: dict[str, Any]):
        """Generic YAML→dataclass coercion for any frozen dataclass with scalar/tuple fields."""
        defaults = cls()
        kwargs: dict[str, Any] = {}
        for f in dataclasses.fields(defaults):
            yaml_val = raw.get(f.name)
            default_val = getattr(defaults, f.name)
            if yaml_val is None:
                kwargs[f.name] = default_val
                continue
            if isinstance(default_val, tuple):
                kwargs[f.name] = (
                    tuple(yaml_val) if isinstance(yaml_val, (list, tuple)) else (str(yaml_val),)
                )
            elif isinstance(default_val, bool):
                kwargs[f.name] = bool(yaml_val)
            elif isinstance(default_val, int):
                kwargs[f.name] = int(yaml_val)
            elif isinstance(default_val, float):
                kwargs[f.name] = float(yaml_val)
            else:
                kwargs[f.name] = str(yaml_val)
        return cls(**kwargs)

    @staticmethod
    def _load_yaml(path: Path) -> dict[str, Any]:
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        if not isinstance(data, dict):
            raise ValueError("Config YAML must be a mapping at the root")
        return data

    @staticmethod
    def _clean_path_str(val: object) -> str:
        s = str(val).strip()
        if len(s) >= 2 and s[0] == s[-1] and s[0] in ("'", '"'):
            s = s[1:-1].strip()
        return s

    @staticmethod
    def _resolve_path(data: dict[str, Any], key: str, base_dir: Path) -> Path:
        val = data.get(key)
        if not val:
            raise ValueError(f"Missing required config key: {key}")
        path = Path(ConfigLoader._clean_path_str(val))
        return path if path.is_absolute() else (base_dir / path).resolve()
