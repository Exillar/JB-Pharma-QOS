"""docx_builder — backward-compatibility shim.

All filler classes now live in the ``builders/`` package.
This module re-exports them so that any external scripts that previously
imported from ``docx_builder`` continue to work unchanged.

New code should import directly from ``builders``.
"""
from builders.s1 import S1DocxFiller, DocxFiller  # noqa: F401
from builders.s2 import S2DocxFiller  # noqa: F401
from builders.s3 import S3DocxFiller  # noqa: F401
from builders.s32 import S32DocxFiller  # noqa: F401
from builders.s4 import S4DocxFiller  # noqa: F401
from builders.s5 import S5DocxFiller  # noqa: F401
from builders.generic import GenericSectionFiller  # noqa: F401
from builders.base import _DocxHelper, run_artifact_cleanup  # noqa: F401

__all__ = [
    "DocxFiller",
    "S1DocxFiller",
    "S2DocxFiller",
    "S3DocxFiller",
    "S32DocxFiller",
    "S4DocxFiller",
    "S5DocxFiller",
    "GenericSectionFiller",
    "_DocxHelper",
    "run_artifact_cleanup",
]
