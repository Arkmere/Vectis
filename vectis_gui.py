"""Minimal Tkinter GUI for Vectis v0.1."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from triage_engine import open_folder, triage_file


class VectisApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Vectis")
        self.geometry("760x420")
        self.minsize(680, 360)

        base_dir = Path(__file__).resolve().parent
        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar(value=str(base_dir / "output"))
        self.reference_folder = tk.StringVar(value=str(base_dir))

        self._build_widgets()

    def _build_widgets(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(4, weight=1)

        ttk.Label(main, text="Input CSV/XLSX").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.input_file).grid(row=0, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(main, text="Browse...", command=self.pick_input_file).grid(row=0, column=2, pady=4)

        ttk.Label(main, text="Output folder").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.output_folder).grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(main, text="Browse...", command=self.pick_output_folder).grid(row=1, column=2, pady=4)

        ttk.Label(main, text="VKB/reference folder").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.reference_folder).grid(row=2, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(main, text="Browse...", command=self.pick_reference_folder).grid(row=2, column=2, pady=4)

        actions = ttk.Frame(main)
        actions.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 8))
        ttk.Button(actions, text="Process file", command=self.process_file).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(actions, text="Open output folder", command=self.open_output_folder).pack(side=tk.LEFT)

        ttk.Label(main, text="Status").grid(row=4, column=0, sticky="nw", pady=4)
        self.status = tk.Text(main, height=10, wrap=tk.WORD)
        self.status.grid(row=4, column=1, columnspan=2, sticky="nsew", padx=8, pady=4)
        self.status.insert(tk.END, "Ready. Select a daily NM CSV/XLSX file and click Process file.\n")

    def log(self, message: str) -> None:
        self.status.insert(tk.END, message + "\n")
        self.status.see(tk.END)
        self.update_idletasks()

    def pick_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select NM input file",
            filetypes=[("CSV and Excel files", "*.csv *.xlsx"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if path:
            self.input_file.set(path)

    def pick_output_folder(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_folder.set(path)

    def pick_reference_folder(self) -> None:
        path = filedialog.askdirectory(title="Select VKB/reference folder")
        if path:
            self.reference_folder.set(path)

    def process_file(self) -> None:
        if not self.input_file.get():
            messagebox.showerror("Vectis", "Please select an input CSV/XLSX file.")
            return
        try:
            self.log(f"Processing: {self.input_file.get()}")
            result = triage_file(self.input_file.get(), self.output_folder.get(), self.reference_folder.get())
            for warning in result.warnings:
                self.log(f"Warning: {warning}")
            self.log(f"Created: {result.output_path}")
            self.log(
                f"Rows: total={result.total_rows}, extracted={result.extracted_rows}, "
                f"remainder={result.remainder_rows}"
            )
            messagebox.showinfo("Vectis", f"Triaged workbook created:\n{result.output_path}")
        except Exception as exc:
            self.log(f"Error: {exc}")
            messagebox.showerror("Vectis", str(exc))

    def open_output_folder(self) -> None:
        try:
            open_folder(self.output_folder.get())
        except Exception as exc:
            self.log(f"Error opening folder: {exc}")
            messagebox.showerror("Vectis", str(exc))


if __name__ == "__main__":
    VectisApp().mainloop()
