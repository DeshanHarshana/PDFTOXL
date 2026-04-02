# PDF to Excel Converter — Developer Guide

## Table of Contents

1. [Overview](#overview)
2. [Technology Stack](#technology-stack)
3. [Project Structure](#project-structure)
4. [Architecture & Data Flow](#architecture--data-flow)
5. [Module Reference](#module-reference)
   - [config.py](#configpy)
   - [pdf_parser.py](#pdf_parserpy)
   - [excel_writer.py](#excel_writerpy)
   - [gui.py](#guipy)
   - [main.py](#mainpy)
6. [PDF Form Field Mapping](#pdf-form-field-mapping)
7. [Excel Output Structure](#excel-output-structure)
8. [How to Run](#how-to-run)
9. [How to Build the .exe](#how-to-build-the-exe)
10. [Common Maintenance Tasks](#common-maintenance-tasks)
11. [Troubleshooting](#troubleshooting)

---

## Overview

This desktop application reads **Wärtsilä cylinder-head overhaul PDF reports** (fillable AcroForm PDFs), extracts measurement data from each one, and writes all records into a **single Excel (.xlsx) file** — one row per PDF.

The target PDFs are **not** plain-text PDFs. They use **AcroForm fillable fields** to store values. The visible text layer only contains labels and headings; actual data lives in the form-field layer.

---

## Technology Stack

| Component | Library | Purpose |
|-----------|---------|---------|
| GUI | **tkinter** (stdlib) | Desktop window, buttons, file dialogs, progress bar |
| PDF parsing | **pdfminer.six** | Read AcroForm fields from the PDF binary structure |
| Excel writing | **openpyxl** | Create `.xlsx` files with formatting, formulas, merged cells |
| Utility | **pdfplumber** | Installed as a dependency (wraps pdfminer); available if table extraction is needed later |
| Utility | **pandas** | Installed; available for future data manipulation needs |

**Python version:** 3.10+ (uses `dict[str, str]` type hints and `X | Y` union syntax).

---

## Project Structure

```
pdf-to-excel-app/
├── main.py              # Entry point — launches the GUI
├── gui.py               # Tkinter application window and UI logic
├── pdf_parser.py        # PDF AcroForm field extraction
├── excel_writer.py      # Excel file generation with template formatting
├── config.py            # All mappings, column definitions, formatting constants
├── requirements.txt     # Python dependencies
├── DEVELOPER_GUIDE.md   # This file
├── .venv/               # Virtual environment (not committed)
├── build/               # PyInstaller build artifacts (not committed)
├── dist/                # Contains PDFtoExcel.exe after build
└── PDFtoExcel.spec      # PyInstaller spec file (auto-generated)
```

---

## Architecture & Data Flow

```
┌──────────────┐     ┌──────────────┐     ┌────────────────┐     ┌──────────┐
│  User picks  │     │  pdf_parser  │     │  excel_writer   │     │  .xlsx   │
│  PDF files   │────>│  parse_pdf() │────>│  write_excel()  │────>│  output  │
│  via GUI     │     │              │     │                 │     │  file    │
└──────────────┘     └──────────────┘     └────────────────┘     └──────────┘
       │                    │                      │
       │              config.py               config.py
       │           FORM_FIELD_MAP          EXCEL_COLUMNS
       │           CYL_HEAD_SN_KEY         NUMBER_FORMATS
       │                                   REJECTION_LIMIT
       │
    gui.py
  - Upload / remove / clear list
  - Convert button → save dialog
  - Background thread for conversion
  - Progress bar updates
```

### Step-by-step flow

1. **User uploads PDFs** via the file dialog in `gui.py`.
2. **User clicks "Convert to Excel"** → a save-file dialog asks where to write the output.
3. **Background thread** iterates over each PDF:
   - `pdf_parser.parse_pdf(path)` opens the PDF binary, reads AcroForm fields, maps them to internal keys using `config.FORM_FIELD_MAP`, and returns a flat `dict`.
4. After all PDFs are parsed, `excel_writer.write_excel(records, output_path)` creates the Excel file:
   - Writes header rows (title, group headers, column names with colours).
   - Writes one data row per record.
   - Writes footer (COUNTIF formulas + rejection limit).
5. **Progress bar** updates after each PDF is processed.
6. **Completion dialog** shows success or lists any PDFs that failed.

---

## Module Reference

### config.py

**Purpose:** Single source of truth for all mappings and constants. Change this file to adjust field mappings, column order, number formats, or the rejection limit.

| Constant | Type | Description |
|----------|------|-------------|
| `FORM_FIELD_MAP` | `dict[str, str]` | Maps PDF AcroForm field names (e.g. `"Cyl head cast no"`) to internal keys (e.g. `"cyl_head_cast_no"`). |
| `CYL_HEAD_SN_KEY` | `str` | The internal key for the required "Cyl Head_SN" field. Parser raises an error if this field is empty. |
| `EXCEL_COLUMNS` | `dict[int, str]` | Maps 1-based Excel column numbers to internal data keys. Defines column order in the output. |
| `NUMBER_FORMATS` | `dict[int, str]` | Excel number format strings per column (e.g. `"0.00"` for 2 decimals). |
| `HEADER_ROW` | `int` | Row number (5) where column header names are written. |
| `DATA_START_ROW` | `int` | Row number (6) where data begins. |
| `REJECTION_LIMIT` | `float` | Clearance rejection threshold (0.156). Used in COUNTIF formulas. |

### pdf_parser.py

**Purpose:** Extract data from a single PDF file.

| Symbol | Description |
|--------|-------------|
| `PDFParseError` | Custom exception raised when a PDF can't be parsed or is missing required fields. |
| `parse_pdf(path)` | **Main entry point.** Returns a `dict` with keys matching `EXCEL_COLUMNS` values. |
| `_extract_form_fields(filepath)` | Low-level: opens the PDF binary via pdfminer, reads the AcroForm catalog, iterates all field objects, decodes names and values. |
| `_smart_decode(raw)` | Handles byte-string decoding: checks for UTF-16 BOM, falls back to UTF-8, then latin-1. Required because PDF field values can be stored in different encodings. |
| `_safe_float(value)` | Converts a string to `float`, returns `None` for blanks or non-numeric strings. |

**Important:** The PDFs use AcroForm fields. Regular text extraction (`pdfplumber.extract_text()`) will NOT find the measurement values — they are stored in form-field objects, not in the page content stream.

### excel_writer.py

**Purpose:** Generate a formatted `.xlsx` file from parsed records.

| Symbol | Description |
|--------|-------------|
| `write_excel(records, output_path, progress_callback=None)` | **Main entry point.** Creates the workbook, writes headers, data rows, and footer. |
| `_write_headers(ws)` | Writes rows 1–5: merged title, group headers (14pt bold), column names with background colours and borders. |
| `_write_data_row(ws, row, record)` | Writes one data row. Converts values to float where possible. Applies number format and alignment per column. |
| `_write_footer(ws, last_data_row)` | Writes COUNTIF formulas for clearance columns, the rejection limit value (0.156), and the "Rejection Limit" label. |
| `_set_column_widths(ws)` | Sets column widths to match the template. |

**Header row colours (row 5):**

| Columns | Colour | Hex Code |
|---------|--------|----------|
| A-I, A-II, A-III (B–D) | Light green | `#C6EFCE` |
| B-I, B-II, B-III (E–G) | Light orange | `#FCD5B4` |
| A-VG-I, A-VG-II, B-VG-I, B-VG-II (H–K) | Light blue | `#BDD7EE` |
| A-Clearance, B-Clearance (L–M) | Light ash/grey | `#D9D9D9` |

### gui.py

**Purpose:** Tkinter-based desktop UI.

| Class / Method | Description |
|---------------|-------------|
| `App(tk.Tk)` | Main window. Manages the PDF path list, button states, and conversion. |
| `_build_ui()` | Constructs all widgets: title label, button bar (Upload/Remove/Clear), scrollable listbox, Convert button, progress bar + status label. |
| `_refresh_button_states()` | Enables/disables buttons based on whether files are loaded and whether conversion is in progress. |
| `_on_upload()` | Opens multi-file dialog, appends unique PDF paths to the list. |
| `_on_remove()` | Removes selected items from the listbox and internal list. |
| `_on_clear()` | Clears all PDFs and resets the progress bar. |
| `_on_convert()` | Opens a save-file dialog, then starts conversion in a background thread. |
| `_convert_worker()` | Runs on a daemon thread: parses each PDF, collects records and errors, then calls `write_excel()`. Posts UI updates via `self.after()`. |
| `_conversion_done()` | Shows a success or warning message box with any errors listed. |

**Threading model:** Conversion runs on a background `threading.Thread` (daemon). UI updates are marshalled to the main thread using `tk.Tk.after(0, callback)`. This prevents the GUI from freezing during long conversions.

### main.py

**Purpose:** Entry point. Creates and runs the `App` instance.

```python
python main.py          # Run from source
PDFtoExcel.exe          # Run as standalone executable
```

---

## PDF Form Field Mapping

The Wärtsilä cylinder-head PDFs contain **100 AcroForm fields**. The app extracts a subset relevant to the Excel summary. Here is the complete mapping:

### Identification Fields

| PDF Field Name | Internal Key | Used in Excel? |
|----------------|-------------|----------------|
| `Installation name` | `installation_name` | No (available for future use) |
| `Application` | `application` | No |
| `Engine No` | `engine_no` | No |
| `Engine type` | `engine_type` | No |
| `Fuel type` | `fuel_type` | No |
| `Engine ref` | `engine_ref_type` | No |
| `Engine rhrs` | `engine_rhrs` | No |
| `Cooling water additive` | `cooling_water_additive` | No |
| `Cylinder head no` | `cylinder_head_no` | No |
| `Cyl head rhrs` | `cyl_head_rhrs` | No |
| **`Cyl head cast no`** | **`cyl_head_cast_no`** | **Yes → Column A (Cyl Head_SN)** |

### Measurement Fields

| PDF Field Name | Internal Key | Excel Column |
|----------------|-------------|-------------|
| `IA` | `a_stem_i` | B (A-I) |
| `IIA` | `a_stem_ii` | C (A-II) |
| `IIIA` | `a_stem_iii` | D (A-III) |
| `IB` | `b_stem_i` | E (B-I) |
| `IIB` | `b_stem_ii` | F (B-II) |
| `IIIB` | `b_stem_iii` | G (B-III) |
| `IA_2` | `a_vg_i` | H (A-VG-I) |
| `IIA_2` | `a_vg_ii` | I (A-VG-II) |
| `IB_2` | `b_vg_i` | J (B-VG-I) |
| `IIB_2` | `b_vg_ii` | K (B-VG-II) |
| `calculatedA` | `a_clearance` | L (A-Clearance) |
| `calculatedB` | `b_clearance` | M (B-Clearance) |

### Discovering Fields in New PDFs

If a new PDF revision has different field names, run this diagnostic script:

```python
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1

with open("new_report.pdf", "rb") as f:
    parser = PDFParser(f)
    doc = PDFDocument(parser)
    acroform = resolve1(doc.catalog["AcroForm"])
    fields = resolve1(acroform["Fields"])
    for i, ref in enumerate(fields):
        field = resolve1(ref)
        name = field.get("T")  # bytes
        value = field.get("V")
        print(f"{i}: name={name}  value={value}")
```

Then update `FORM_FIELD_MAP` in `config.py` accordingly.

---

## Excel Output Structure

### Layout

```
Row 1-2:  "CYLINDER HEAD MEASUREMENT RECORD SUMMARY" (merged A1:M2, bold 12pt)
Row 3:    Group headers — "Stem diameter (ØD1) in mm" | "Valve guide (Ød1) in mm" | "Clearance" (bold 14pt)
Row 4:    Sub-group headers — "INLET VALVE - A" | "INLET VALVE - B" | "Calculated" (bold 14pt)
Row 5:    Column names with coloured backgrounds and thin borders
Row 6+:   Data rows (one per PDF)
Row N+1:  COUNTIF formulas counting clearances ≥ rejection limit
Row N+2:  Rejection limit value (0.156, merged L:M)
Row N+3:  "Rejection Limit" label (bold, merged L:M)
```

### Column Layout (Row 5 onward)

| Column | Header | Number Format | Background |
|--------|--------|--------------|------------|
| A | Cyl Head_SN | `0` (integer) | — |
| B | A-I | `0.00` | Light green |
| C | A-II | `0.00` | Light green |
| D | A-III | `0.00` | Light green |
| E | B-I | `0.00` | Light orange |
| F | B-II | `0.00` | Light orange |
| G | B-III | `0.00` | Light orange |
| H | A-VG-I | `0.00` | Light blue |
| I | A-VG-II | `0.00` | Light blue |
| J | B-VG-I | `0.00` | Light blue |
| K | B-VG-II | `0.00` | Light blue |
| L | A-Clearance | `0.000` | Light ash |
| M | B-Clearance | `0.000` | Light ash |

---

## How to Run

### From Source

```powershell
cd c:\Work\Bappa\pdf-to-excel-app
.\.venv\Scripts\Activate.ps1
python main.py
```

### As Standalone Executable

```powershell
dist\PDFtoExcel.exe
```

No Python installation required on the target machine.

---

## How to Build the .exe

```powershell
cd c:\Work\Bappa\pdf-to-excel-app
.\.venv\Scripts\Activate.ps1
pip install pyinstaller
pyinstaller --onefile --windowed --name "PDFtoExcel" main.py
```

The output is `dist\PDFtoExcel.exe` (~39 MB). To add an icon:

```powershell
pyinstaller --onefile --windowed --name "PDFtoExcel" --icon=app_icon.ico main.py
```

---

## Common Maintenance Tasks

### Adding a New Excel Column

1. **config.py** — Add the PDF field mapping in `FORM_FIELD_MAP`:
   ```python
   "NewFieldName": "new_internal_key",
   ```
2. **config.py** — Add the column in `EXCEL_COLUMNS` (pick the next column number) and in `NUMBER_FORMATS`:
   ```python
   EXCEL_COLUMNS = {
       ...
       14: "new_internal_key",   # N: New Column
   }
   NUMBER_FORMATS = {
       ...
       14: "0.00",
   }
   ```
3. **excel_writer.py** — Add the column name in `_write_headers()` → `col_names` dict. Update merged cell ranges if the new column falls under an existing group header. Add a background fill in `col_fills` if desired.
4. **excel_writer.py** — Update `_COL_WIDTHS` to include the new column letter.
5. **pdf_parser.py** — If it's a measurement, add the key to the `measurement_keys` list in `parse_pdf()`.

### Changing a Column Colour

Edit `_write_headers()` in `excel_writer.py`. The fill colours are defined as `PatternFill` objects:

```python
fill_green  = PatternFill(start_color="C6EFCE", ...)  # A-I, A-II, A-III
fill_orange = PatternFill(start_color="FCD5B4", ...)  # B-I, B-II, B-III
fill_blue   = PatternFill(start_color="BDD7EE", ...)  # VG columns
fill_ash    = PatternFill(start_color="D9D9D9", ...)  # Clearance columns
```

Change the hex colour code as needed.

### Changing the Rejection Limit

Edit `REJECTION_LIMIT` in `config.py`:

```python
REJECTION_LIMIT = 0.200  # new value
```

### Supporting a New PDF Format

1. Run the field-discovery script (see [Discovering Fields in New PDFs](#discovering-fields-in-new-pdfs)) to list all form field names in the new PDF.
2. Update `FORM_FIELD_MAP` in `config.py` with any new or changed field names.
3. If the new PDF is not an AcroForm (plain text), `pdf_parser.py` would need a new extraction path using `pdfplumber.extract_text()` or `pdfplumber.extract_tables()`.

### Adding a New Required Field

In `pdf_parser.py` → `parse_pdf()`, add validation after the existing `cyl_head_sn` check:

```python
new_field = mapped.get("new_internal_key", "").strip()
if not new_field:
    raise PDFParseError("Required field 'X' is empty or missing.")
```

---

## Troubleshooting

### "Required field 'Cyl head cast no' is empty or missing"

- The PDF might not have AcroForm fields (it could be a scanned image or flat PDF).
- The form field name might have changed in a new document revision. Run the discovery script to check.

### Garbled text in extracted values

The `_smart_decode()` function handles UTF-16 (with BOM), UTF-8, and latin-1. If a new encoding appears, extend the decode chain in `pdf_parser.py`.

### GUI freezes during conversion

This should not happen — conversion runs on a background thread. If it does, check that `self.after(0, ...)` calls are being used (not direct widget access from the worker thread).

### PyInstaller build fails

- Ensure the virtual environment is activated before running PyInstaller.
- If a hidden import is missing, add it: `pyinstaller --hidden-import=module_name ...`
- Check `build\PDFtoExcel\warn-PDFtoExcel.txt` for missing module warnings.

### Excel file won't open / is corrupted

- Ensure `openpyxl` is up to date: `pip install --upgrade openpyxl`.
- Check that no two `merge_cells()` calls overlap in `_write_headers()`.
