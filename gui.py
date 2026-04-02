"""
Tkinter GUI for the PDF → Excel converter.

Features
--------
- Upload multiple PDFs  (file dialog)
- View uploaded file list
- Remove a single PDF or clear all
- Convert button (enabled only when ≥ 1 PDF)
- Save-file dialog on convert
- Progress bar with *n / total* feedback
- Error messages displayed gracefully (no crash)
"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from pdf_parser import PDFParseError, parse_pdf
from excel_writer import write_excel


class App(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("PDF to Excel Converter")
        self.geometry("620x480")
        self.minsize(520, 400)
        self.resizable(True, True)

        self._pdf_paths: list[Path] = []
        self._is_converting = False

        self._build_ui()
        self._refresh_button_states()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        pad = dict(padx=8, pady=4)

        # --- Title ---
        ttk.Label(
            self, text="Cylinder Head  PDF \u2192 Excel",
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor=tk.W, **pad)

        # --- Button bar ---
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, **pad)

        self._btn_upload = ttk.Button(
            btn_frame, text="Upload PDFs\u2026", command=self._on_upload,
        )
        self._btn_upload.pack(side=tk.LEFT, padx=(0, 4))

        self._btn_remove = ttk.Button(
            btn_frame, text="Remove Selected", command=self._on_remove,
        )
        self._btn_remove.pack(side=tk.LEFT, padx=(0, 4))

        self._btn_clear = ttk.Button(
            btn_frame, text="Clear All", command=self._on_clear,
        )
        self._btn_clear.pack(side=tk.LEFT)

        # --- File list ---
        list_frame = ttk.LabelFrame(self, text="Uploaded PDFs")
        list_frame.pack(fill=tk.BOTH, expand=True, **pad)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self._listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 10),
        )
        scrollbar.config(command=self._listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- Convert ---
        convert_frame = ttk.Frame(self)
        convert_frame.pack(fill=tk.X, **pad)

        self._btn_convert = ttk.Button(
            convert_frame, text="Convert to Excel",
            command=self._on_convert,
        )
        self._btn_convert.pack(side=tk.LEFT)

        # --- Progress ---
        progress_frame = ttk.Frame(self)
        progress_frame.pack(fill=tk.X, **pad)

        self._progress = ttk.Progressbar(
            progress_frame, orient=tk.HORIZONTAL, mode="determinate",
        )
        self._progress.pack(fill=tk.X, side=tk.LEFT, expand=True)

        self._lbl_status = ttk.Label(
            progress_frame, text="", width=18, anchor=tk.E,
        )
        self._lbl_status.pack(side=tk.RIGHT, padx=(8, 0))

    # ------------------------------------------------------------------
    # Button-state helpers
    # ------------------------------------------------------------------

    def _refresh_button_states(self) -> None:
        has_files = len(self._pdf_paths) > 0
        converting = self._is_converting

        state_normal = "!disabled"
        state_disabled = "disabled"

        self._btn_upload.state([state_disabled if converting else state_normal])
        self._btn_remove.state(
            [state_normal if has_files and not converting else state_disabled]
        )
        self._btn_clear.state(
            [state_normal if has_files and not converting else state_disabled]
        )
        self._btn_convert.state(
            [state_normal if has_files and not converting else state_disabled]
        )

    def _sync_listbox(self) -> None:
        self._listbox.delete(0, tk.END)
        for p in self._pdf_paths:
            self._listbox.insert(tk.END, p.name)
        self._refresh_button_states()

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _on_upload(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not files:
            return
        for f in files:
            p = Path(f)
            if p not in self._pdf_paths:
                self._pdf_paths.append(p)
        self._sync_listbox()

    def _on_remove(self) -> None:
        selected = list(self._listbox.curselection())
        if not selected:
            messagebox.showinfo("Remove", "Select one or more PDFs to remove.")
            return
        for idx in reversed(selected):
            del self._pdf_paths[idx]
        self._sync_listbox()

    def _on_clear(self) -> None:
        self._pdf_paths.clear()
        self._sync_listbox()
        self._progress["value"] = 0
        self._lbl_status.config(text="")

    def _on_convert(self) -> None:
        output_path = filedialog.asksaveasfilename(
            title="Save Excel file as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="cylinder_head_summary.xlsx",
        )
        if not output_path:
            return
        self._start_conversion(Path(output_path))

    # ------------------------------------------------------------------
    # Conversion (runs on a background thread)
    # ------------------------------------------------------------------

    def _start_conversion(self, output_path: Path) -> None:
        self._is_converting = True
        self._refresh_button_states()

        total = len(self._pdf_paths)
        self._progress["maximum"] = total
        self._progress["value"] = 0
        self._lbl_status.config(text=f"0 / {total}")

        thread = threading.Thread(
            target=self._convert_worker,
            args=(list(self._pdf_paths), output_path),
            daemon=True,
        )
        thread.start()

    def _convert_worker(
        self, paths: list[Path], output_path: Path,
    ) -> None:
        """Background worker: parse each PDF, then write Excel."""
        total = len(paths)
        records: list[dict[str, Any]] = []
        errors: list[str] = []

        for idx, pdf_path in enumerate(paths):
            try:
                record = parse_pdf(pdf_path)
                records.append(record)
            except PDFParseError as exc:
                errors.append(f"{pdf_path.name}: {exc}")
            except Exception as exc:
                errors.append(f"{pdf_path.name}: unexpected error — {exc}")

            self.after(0, self._update_progress, idx + 1, total)

        if not records and errors:
            self.after(0, self._conversion_done, output_path, errors, 0)
            return

        try:
            write_excel(records, output_path)
        except Exception as exc:
            errors.append(f"Excel write failed: {exc}")

        self.after(
            0, self._conversion_done, output_path, errors, len(records),
        )

    # ------------------------------------------------------------------
    # UI callbacks from worker (run on main thread via .after)
    # ------------------------------------------------------------------

    def _update_progress(self, current: int, total: int) -> None:
        self._progress["value"] = current
        self._lbl_status.config(text=f"{current} / {total}")

    def _conversion_done(
        self, output_path: Path, errors: list[str], written: int,
    ) -> None:
        self._is_converting = False
        self._refresh_button_states()

        if errors:
            detail = "\n".join(errors)
            messagebox.showwarning(
                "Conversion completed with warnings",
                f"{written} PDF(s) written to Excel.\n\n"
                f"Errors ({len(errors)}):\n{detail}",
            )
        else:
            messagebox.showinfo(
                "Done",
                f"Successfully wrote {written} record(s) to:\n{output_path}",
            )
