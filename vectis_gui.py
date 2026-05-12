"""Tkinter GUI for the Vectis analyst triage workbook."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from triage_engine import open_folder, triage_file

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "vectis.png"
APP_VERSION = "v1.0.0"


class VectisApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Vectis")
        self.geometry("980x700")
        self.minsize(860, 600)

        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar(value=str(BASE_DIR / "output"))
        self.reference_folder = tk.StringVar(value=str(BASE_DIR))
        self.state_text = tk.StringVar(value="Ready")
        self.logo_image: tk.PhotoImage | None = None

        self._configure_style()
        self._build_widgets()
        self._load_logo()
        self.log("Ready. Select a daily NM CSV/XLSX file and click Process file.")

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#f6f8fb")
        style.configure("Card.TFrame", background="#ffffff", relief="flat")
        style.configure("Header.TLabel", background="#f6f8fb", foreground="#102a43", font=("Segoe UI", 24, "bold"))
        style.configure("Subtitle.TLabel", background="#f6f8fb", foreground="#486581", font=("Segoe UI", 11))
        style.configure("Section.TLabel", background="#ffffff", foreground="#102a43", font=("Segoe UI", 11, "bold"))
        style.configure("TLabel", background="#ffffff", foreground="#243b53", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        style.configure("TButton", padding=(10, 6))
        style.configure("Status.TLabel", background="#e6f0fa", foreground="#102a43", padding=(10, 6))

    def _build_widgets(self) -> None:
        self.configure(background="#f6f8fb")
        main = ttk.Frame(self, padding=18)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(3, weight=1)

        header = ttk.Frame(main)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.columnconfigure(1, weight=1)
        self.logo_label = ttk.Label(header, text="Vectis", style="Subtitle.TLabel")
        self.logo_label.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 16))
        ttk.Label(header, text="Vectis", style="Header.TLabel").grid(row=0, column=1, sticky="sw")
        ttk.Label(header, text=f"Aviation Intelligence Triage • Internal Tool • {APP_VERSION}", style="Subtitle.TLabel").grid(row=1, column=1, sticky="nw")

        paths = ttk.Frame(main, style="Card.TFrame", padding=16)
        paths.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        paths.columnconfigure(1, weight=1)
        ttk.Label(paths, text="Input & Paths", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        self._path_row(paths, 1, "Input CSV/XLSX", self.input_file, self.pick_input_file)
        self._path_row(paths, 2, "Output folder", self.output_folder, self.pick_output_folder)
        self._path_row(paths, 3, "VKB/reference folder", self.reference_folder, self.pick_reference_folder)

        actions = ttk.Frame(main, style="Card.TFrame", padding=16)
        actions.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(actions, text="Actions", style="Section.TLabel").pack(side=tk.LEFT, padx=(0, 18))
        ttk.Button(actions, text="Process file", style="Accent.TButton", command=self.process_file).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(actions, text="Open output folder", command=self.open_output_folder).pack(side=tk.LEFT)

        log_card = ttk.Frame(main, style="Card.TFrame", padding=16)
        log_card.grid(row=3, column=0, sticky="nsew", pady=(0, 12))
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)
        log_header = ttk.Frame(log_card, style="Card.TFrame")
        log_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(log_header, text="Status / Log", style="Section.TLabel").pack(side=tk.LEFT)
        ttk.Button(log_header, text="Clear", command=self.clear_log).pack(side=tk.RIGHT)
        self.status = tk.Text(log_card, height=15, wrap=tk.WORD, relief=tk.FLAT, bg="#fbfdff", fg="#102a43", font=("Consolas", 10), padx=10, pady=10)
        self.status.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_card, command=self.status.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.status.configure(yscrollcommand=scrollbar.set)

        ttk.Label(main, textvariable=self.state_text, style="Status.TLabel").grid(row=4, column=0, sticky="ew")

    def _path_row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=5, padx=(0, 12))
        ttk.Entry(parent, textvariable=variable, width=88).grid(row=row, column=1, sticky="ew", pady=5, padx=(0, 10))
        ttk.Button(parent, text="Browse...", command=command).grid(row=row, column=2, sticky="e", pady=5)

    def _load_logo(self) -> None:
        try:
            self.logo_image = tk.PhotoImage(file=str(LOGO_PATH))
            max_width = 92
            if self.logo_image.width() > max_width:
                factor = max(1, round(self.logo_image.width() / max_width))
                self.logo_image = self.logo_image.subsample(factor, factor)
            self.logo_label.configure(image=self.logo_image, text="")
        except Exception:
            self.logo_image = None
            self.logo_label.configure(text="Vectis")
            self.log(f"Warning: logo asset not found at assets/vectis.png; continuing without logo.")

    def log(self, message: str) -> None:
        self.status.insert(tk.END, message + "\n")
        self.status.see(tk.END)
        self.update_idletasks()

    def clear_log(self) -> None:
        self.status.delete("1.0", tk.END)

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
            self.state_text.set("Processing")
            self.log(f"Processing: {self.input_file.get()}")
            result = triage_file(self.input_file.get(), self.output_folder.get(), self.reference_folder.get())
            for warning in result.warnings:
                self.log(f"Warning: {warning}")
            self.log(f"Created: {result.output_path}")
            self.log(
                f"Rows: total={result.total_rows}, extracted={result.extracted_rows}, "
                f"remainder={result.remainder_rows}"
            )
            self.state_text.set("Complete")
            messagebox.showinfo("Vectis", f"Triaged workbook created:\n{result.output_path}")
        except Exception as exc:
            self.state_text.set("Error")
            self.log(f"Error: {exc}")
            messagebox.showerror("Vectis", str(exc))
        finally:
            if self.state_text.get() == "Processing":
                self.state_text.set("Ready")

    def open_output_folder(self) -> None:
        try:
            self.state_text.set("Opening output folder")
            open_folder(self.output_folder.get())
            self.state_text.set("Ready")
        except Exception as exc:
            self.state_text.set("Error")
            self.log(f"Error opening folder: {exc}")
            messagebox.showerror("Vectis", str(exc))


if __name__ == "__main__":
    VectisApp().mainloop()
