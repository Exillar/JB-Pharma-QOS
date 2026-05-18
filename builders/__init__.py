"""builders — section-filler package for QOS document generation."""
from builders.s1 import S1DocxFiller, DocxFiller  # DocxFiller = backward-compat alias
from builders.s2 import S2DocxFiller
from builders.s3 import S3DocxFiller
from builders.s32 import S32DocxFiller
from builders.s4 import S4DocxFiller
from builders.s5 import S5DocxFiller
from builders.generic import GenericSectionFiller
from builders.p1 import P1DocxFiller

__all__ = [
    "S1DocxFiller",
    "DocxFiller",
    "S2DocxFiller",
    "S3DocxFiller",
    "S32DocxFiller",
    "S4DocxFiller",
    "S5DocxFiller",
    "GenericSectionFiller",
    "P1DocxFiller",
]
