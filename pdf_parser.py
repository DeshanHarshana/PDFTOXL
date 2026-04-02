"""
Extract structured data from Wärtsilä cylinder-head overhaul PDF reports.

These PDFs use AcroForm (fillable form fields) to store measurement
values — the text layer only contains labels/headings.

Public API
----------
parse_pdf(path)  →  dict   with keys matching config.EXCEL_COLUMNS values
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1

from config import CYL_HEAD_SN_KEY, FORM_FIELD_MAP


class PDFParseError(Exception):
    """Raised when a PDF cannot be parsed or is missing required fields."""


# ---------------------------------------------------------------------------
# Byte-string decoding (handles BOM, UTF-8, and latin-1 gracefully)
# ---------------------------------------------------------------------------

def _smart_decode(raw: Any) -> str | None:
    """Decode a raw PDF field value (bytes, int, float, …) to a Python str."""
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return str(raw)
    if not isinstance(raw, bytes):
        return str(raw)

    # UTF-16 with BOM
    if raw[:2] in (b"\xfe\xff", b"\xff\xfe"):
        return raw.decode("utf-16").strip("\x00").strip()

    # Plain bytes — try UTF-8 first (covers ASCII), then latin-1
    try:
        return raw.decode("utf-8").strip()
    except UnicodeDecodeError:
        return raw.decode("latin-1").strip()


# ---------------------------------------------------------------------------
# AcroForm extraction
# ---------------------------------------------------------------------------

def _extract_form_fields(filepath: Path) -> dict[str, str]:
    """Return ``{form_field_name: value_string}`` for every field in the PDF."""
    with open(filepath, "rb") as fh:
        parser = PDFParser(fh)
        doc = PDFDocument(parser)

        catalog = doc.catalog
        if "AcroForm" not in catalog:
            raise PDFParseError("PDF has no AcroForm (no fillable fields found).")

        acroform = resolve1(catalog["AcroForm"])
        raw_fields = resolve1(acroform.get("Fields", []))

        result: dict[str, str] = {}
        for ref in raw_fields:
            field = resolve1(ref)
            name = _smart_decode(field.get("T"))
            value = _smart_decode(field.get("V"))
            if name is not None and value is not None and value != "":
                result[name] = value

    return result


# ---------------------------------------------------------------------------
# Safe float conversion
# ---------------------------------------------------------------------------

def _safe_float(value: str | None) -> float | None:
    """Convert a string value to float, returning None for blanks."""
    if value is None:
        return None
    value = value.strip()
    if value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def parse_pdf(path: str | Path) -> dict[str, Any]:
    """Parse a single cylinder-head PDF and return a flat dict whose keys
    match the values in ``config.EXCEL_COLUMNS``.

    Raises ``PDFParseError`` with a human-readable message on failure.
    """
    path = Path(path)
    if not path.is_file():
        raise PDFParseError(f"File not found: {path}")

    try:
        form_fields = _extract_form_fields(path)
    except PDFParseError:
        raise
    except Exception as exc:
        raise PDFParseError(f"Cannot read PDF form fields: {exc}") from exc

    # Map form-field names → internal keys
    mapped: dict[str, str] = {}
    for field_name, internal_key in FORM_FIELD_MAP.items():
        if field_name in form_fields:
            mapped[internal_key] = form_fields[field_name]

    # Validate required field
    cyl_head_sn = mapped.get(CYL_HEAD_SN_KEY, "").strip()
    if not cyl_head_sn:
        raise PDFParseError(
            "Required field 'Cyl head cast no' is empty or missing in the PDF."
        )

    # Build result dict with the keys expected by excel_writer
    result: dict[str, Any] = {"cyl_head_sn": cyl_head_sn}

    measurement_keys = [
        "a_stem_i", "a_stem_ii", "a_stem_iii",
        "b_stem_i", "b_stem_ii", "b_stem_iii",
        "a_vg_i", "a_vg_ii",
        "b_vg_i", "b_vg_ii",
        "a_clearance", "b_clearance",
    ]
    for key in measurement_keys:
        result[key] = _safe_float(mapped.get(key))

    return result
