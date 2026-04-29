from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml


@dataclass(frozen=True)
class AppConfig:
    template_docx: Path
    dossier_root: Path
    filled_reference_docx: Path
    output_docx: Path
    artifacts_dir: Path
    extractor_backend: str
    section: str

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
        )


@dataclass(frozen=True)
class PipelineConfig:
    template_docx: Path
    dossier_root: Path
    filled_reference_docx: Path
    output_docx: Path
    artifacts_dir: Path
    extractor_backend: str = "pymupdf"

    @property
    def module3_root(self) -> Path:
        return self.dossier_root / "Module 3"

    @property
    def image_artifacts_dir(self) -> Path:
        return self.artifacts_dir / "images"

    @property
    def verification_report_path(self) -> Path:
        return self.artifacts_dir / "verification_report.txt"

    @property
    def generation_log_path(self) -> Path:
        return self.artifacts_dir / "generation.log"

    def ensure_directories(self) -> None:
        self.artifacts_dir.mkdir(parents=True, exist_ok=True)
        self.image_artifacts_dir.mkdir(parents=True, exist_ok=True)
        self.output_docx.parent.mkdir(parents=True, exist_ok=True)


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
        )

    @staticmethod
    def _load_yaml(path: Path) -> dict[str, Any]:
        raw = path.read_text(encoding="utf-8")
        data = yaml.safe_load(raw) or {}
        if not isinstance(data, dict):
            raise ValueError("Config YAML must be a mapping at the root")
        return data

    @staticmethod
    def _resolve_path(data: dict[str, Any], key: str, base_dir: Path) -> Path:
        val = data.get(key)
        if not val:
            raise ValueError(f"Missing required config key: {key}")
        path = Path(str(val))
        return path if path.is_absolute() else (base_dir / path).resolve()
