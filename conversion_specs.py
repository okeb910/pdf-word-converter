"""Conversion direction metadata shared by the GUI and batch runner."""

from __future__ import annotations

from dataclasses import dataclass


PDF_TARGET_WORD = "word"
PDF_TARGET_POWERPOINT = "powerpoint"


@dataclass(frozen=True)
class ConversionSpec:
    key: str
    source_kind: str
    target_kind: str
    source_extensions: tuple[str, ...]
    target_suffix: str
    label: str
    engine_policy: str


PDF_TO_WORD = ConversionSpec(
    "pdf_to_word", "pdf", "word", (".pdf",), ".docx", "PDF -> Word", "selectable"
)
PDF_TO_POWERPOINT = ConversionSpec(
    "pdf_to_powerpoint",
    "pdf",
    "powerpoint",
    (".pdf",),
    ".pptx",
    "PDF -> PowerPoint",
    "builtin_images",
)
WORD_TO_PDF = ConversionSpec(
    "word_to_pdf", "word", "pdf", (".docx",), ".pdf", "Word -> PDF", "automatic"
)
POWERPOINT_TO_PDF = ConversionSpec(
    "powerpoint_to_pdf",
    "powerpoint",
    "pdf",
    (".ppt", ".pptx"),
    ".pdf",
    "PowerPoint -> PDF",
    "automatic",
)


SPECS_BY_KEY = {
    spec.key: spec
    for spec in (PDF_TO_WORD, PDF_TO_POWERPOINT, WORD_TO_PDF, POWERPOINT_TO_PDF)
}


def resolve_conversion_spec(source_kind: str, pdf_target: str = PDF_TARGET_WORD) -> ConversionSpec:
    if source_kind == "pdf":
        if pdf_target == PDF_TARGET_WORD:
            return PDF_TO_WORD
        if pdf_target == PDF_TARGET_POWERPOINT:
            return PDF_TO_POWERPOINT
        raise ValueError(f"Unsupported PDF target: {pdf_target}")
    if source_kind == "word":
        return WORD_TO_PDF
    if source_kind == "powerpoint":
        return POWERPOINT_TO_PDF
    raise ValueError(f"Unsupported source kind: {source_kind}")


def source_kind_for_extension(extension: str) -> str:
    normalized = extension.lower()
    if normalized == ".pdf":
        return "pdf"
    if normalized == ".docx":
        return "word"
    if normalized in (".ppt", ".pptx"):
        return "powerpoint"
    raise ValueError(f"Unsupported source extension: {extension}")
