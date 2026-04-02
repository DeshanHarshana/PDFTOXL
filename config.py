"""
Central configuration for PDF form-field names, Excel column mappings,
template formatting, and formula templates.
"""

# ---------------------------------------------------------------------------
# PDF AcroForm field name → internal data key
# ---------------------------------------------------------------------------
# The PDF stores values in fillable form fields (AcroForm), not in the
# text layer.  Keys below are the form-field names exactly as stored in the
# PDF; values are the internal keys used throughout the app.

FORM_FIELD_MAP: dict[str, str] = {
    # Header / identification fields
    "Installation name":        "installation_name",
    "Application":              "application",
    "Engine No":                "engine_no",
    "Engine type":              "engine_type",
    "Fuel type":                "fuel_type",
    "Engine ref":               "engine_ref_type",
    "Engine rhrs":              "engine_rhrs",
    "Cooling water additive":   "cooling_water_additive",
    "Cylinder head no":         "cylinder_head_no",
    "Cyl head rhrs":            "cyl_head_rhrs",
    "Cyl head cast no":         "cyl_head_cast_no",

    # Inlet Valve A — stem diameter
    "IA":   "a_stem_i",
    "IIA":  "a_stem_ii",
    "IIIA": "a_stem_iii",

    # Inlet Valve B — stem diameter
    "IB":   "b_stem_i",
    "IIB":  "b_stem_ii",
    "IIIB": "b_stem_iii",

    # Inlet Valve A — valve guide
    "IA_2":  "a_vg_i",
    "IIA_2": "a_vg_ii",

    # Inlet Valve B — valve guide
    "IB_2":  "b_vg_i",
    "IIB_2": "b_vg_ii",

    # Clearance (calculated)
    "calculatedA": "a_clearance",
    "calculatedB": "b_clearance",
}

# The internal key that maps to the Excel "Cyl Head_SN" column
CYL_HEAD_SN_KEY = "cyl_head_cast_no"

# ---------------------------------------------------------------------------
# Excel column mapping  (1-based column numbers matching the template)
# ---------------------------------------------------------------------------
EXCEL_COLUMNS: dict[int, str] = {
    1: "cyl_head_sn",       # A: Cyl Head_SN
    2: "a_stem_i",          # B: A-I
    3: "a_stem_ii",         # C: A-II
    4: "a_stem_iii",        # D: A-III
    5: "b_stem_i",          # E: B-I
    6: "b_stem_ii",         # F: B-II
    7: "b_stem_iii",        # G: B-III
    8: "a_vg_i",            # H: A-VG-I
    9: "a_vg_ii",           # I: A-VG-II
    10: "b_vg_i",           # J: B-VG-I
    11: "b_vg_ii",          # K: B-VG-II
    12: "a_clearance",      # L: A-Clearance
    13: "b_clearance",      # M: B-Clearance
}

# ---------------------------------------------------------------------------
# Excel formatting (copied from the template's row-6 style)
# ---------------------------------------------------------------------------
HEADER_ROW = 5          # Row with column headers (Location, Cyl Head_SN, …)
DATA_START_ROW = 6      # First data row

# Number formats per column (matching template)
NUMBER_FORMATS: dict[int, str] = {
    1: "0",       # A  Cyl Head_SN
    2: "0.00",    # B  A-I
    3: "0.00",    # C  A-II
    4: "0.00",    # D  A-III
    5: "0.00",    # E  B-I
    6: "0.00",    # F  B-II
    7: "0.00",    # G  B-III
    8: "0.00",    # H  A-VG-I
    9: "0.00",    # I  A-VG-II
    10: "0.00",   # J  B-VG-I
    11: "0.00",   # K  B-VG-II
    12: "0.000",  # L  A-Clearance
    13: "0.000",  # M  B-Clearance
}

REJECTION_LIMIT = 0.156
