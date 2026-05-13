"""Tkinter GUI for the Vectis analyst triage workbook."""

from __future__ import annotations

import queue
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from triage_engine import TriageResult, open_folder, triage_file

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "vectis.png"
APP_VERSION = "v1.0.0"

VECTAIR_HEADER = "#6F8890"
VECTAIR_HEADER_DARK = "#4F6870"
VECTAIR_ACCENT = "#6B5638"
VECTAIR_BG = "#EEF2F1"
VECTAIR_PANEL = "#F7F8F7"
VECTAIR_BORDER = "#BFC8C8"
VECTAIR_TEXT = "#1D2A2E"
VECTAIR_MUTED = "#5F6D70"
VECTAIR_OK = "#4D8A57"
VECTAIR_WARN = "#B3842F"
VECTAIR_ERROR = "#A64040"

PROGRESS_POLL_MS = 150
HEARTBEAT_MS = 750


class VectisApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Vectis")
        self.geometry("980x720")
        self.minsize(880, 620)

        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar(value=str(BASE_DIR / "output"))
        self.reference_folder = tk.StringVar(value=str(BASE_DIR))
        self.state_text = tk.StringVar(value="Engine: ready")
        self.progress_text = tk.StringVar(value="Idle / ready 0%")
        self.progress_value = tk.IntVar(value=0)
        self.logo_image: tk.PhotoImage | None = None

        self._worker_thread: threading.Thread | None = None
        self._progress_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._processing = False
        self._started_at = 0.0
        self._last_stage = "Idle / ready"
        self._browse_buttons: list[ttk.Button] = []

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
        style.configure("TFrame", background=VECTAIR_BG)
        style.configure("Card.TFrame", background=VECTAIR_PANEL, relief="flat")
        style.configure("Header.TFrame", background=VECTAIR_HEADER)
        style.configure("HeaderTitle.TLabel", background=VECTAIR_HEADER, foreground="white", font=("Segoe UI", 22, "bold"))
        style.configure("HeaderSubtitle.TLabel", background=VECTAIR_HEADER, foreground="#F1F4F3", font=("Segoe UI", 10))
        style.configure("HeaderStatus.TLabel", background=VECTAIR_HEADER, foreground="#F1F4F3", font=("Segoe UI", 10, "bold"))
        style.configure("Section.TLabel", background=VECTAIR_PANEL, foreground=VECTAIR_TEXT, font=("Segoe UI", 11, "bold"))
        style.configure("TLabel", background=VECTAIR_PANEL, foreground=VECTAIR_TEXT, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=VECTAIR_PANEL, foreground=VECTAIR_MUTED, font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=VECTAIR_HEADER_DARK, foreground="white", padding=(10, 6), font=("Segoe UI", 10, "bold"))
        style.configure("TButton", padding=(10, 6), background=VECTAIR_PANEL, foreground=VECTAIR_TEXT, bordercolor=VECTAIR_BORDER)
        style.configure("Accent.TButton", padding=(16, 7), background=VECTAIR_ACCENT, foreground="white", font=("Segoe UI", 10, "bold"), bordercolor=VECTAIR_ACCENT)
        style.map("Accent.TButton", background=[("disabled", VECTAIR_BORDER), ("active", VECTAIR_HEADER_DARK)], foreground=[("disabled", VECTAIR_MUTED), ("active", "white")])
        style.configure("Vectair.Horizontal.TProgressbar", troughcolor="#DDE5E3", background=VECTAIR_ACCENT, bordercolor=VECTAIR_BORDER, lightcolor=VECTAIR_ACCENT, darkcolor=VECTAIR_ACCENT)

    def _build_widgets(self) -> None:
        self.configure(background=VECTAIR_BG)
        main = ttk.Frame(self, padding=18)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)

        header = ttk.Frame(main, style="Header.TFrame", padding=(14, 10))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(1, weight=1)
        self.logo_label = ttk.Label(header, text="Vectis", style="HeaderSubtitle.TLabel")
        self.logo_label.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 14))
        ttk.Label(header, text="VECTIS", style="HeaderTitle.TLabel").grid(row=0, column=1, sticky="sw")
        ttk.Label(header, text="Aviation Intelligence Triage", style="HeaderSubtitle.TLabel").grid(row=1, column=1, sticky="nw")
        ttk.Label(header, text=f"Version: {APP_VERSION}", style="HeaderStatus.TLabel").grid(row=0, column=2, sticky="e", padx=(14, 0))
        ttk.Label(header, textvariable=self.state_text, style="HeaderStatus.TLabel").grid(row=1, column=2, sticky="e", padx=(14, 0))

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
        self.process_button = ttk.Button(actions, text="Process file", style="Accent.TButton", command=self.process_file)
        self.process_button.pack(side=tk.LEFT, padx=(0, 10))
        self.open_button = ttk.Button(actions, text="Open output folder", command=self.open_output_folder)
        self.open_button.pack(side=tk.LEFT)

        progress_card = ttk.Frame(main, style="Card.TFrame", padding=(16, 12))
        progress_card.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        progress_card.columnconfigure(0, weight=1)
        ttk.Label(progress_card, text="Progress", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(progress_card, textvariable=self.progress_text, style="Muted.TLabel").grid(row=0, column=1, sticky="e", pady=(0, 6))
        self.progress_bar = ttk.Progressbar(progress_card, variable=self.progress_value, maximum=100, mode="determinate", style="Vectair.Horizontal.TProgressbar")
        self.progress_bar.grid(row=1, column=0, columnspan=2, sticky="ew")

        log_card = ttk.Frame(main, style="Card.TFrame", padding=16)
        log_card.grid(row=4, column=0, sticky="nsew", pady=(0, 12))
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)
        log_header = ttk.Frame(log_card, style="Card.TFrame")
        log_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(log_header, text="Status / Log", style="Section.TLabel").pack(side=tk.LEFT)
        ttk.Button(log_header, text="Clear", command=self.clear_log).pack(side=tk.RIGHT)
        self.status = tk.Text(log_card, height=14, wrap=tk.WORD, relief=tk.SOLID, bd=1, bg="#FBFCFB", fg=VECTAIR_TEXT, insertbackground=VECTAIR_TEXT, highlightthickness=1, highlightbackground=VECTAIR_BORDER, font=("Consolas", 10), padx=10, pady=10)
        self.status.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_card, command=self.status.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.status.configure(yscrollcommand=scrollbar.set)

        ttk.Label(main, textvariable=self.state_text, style="Status.TLabel").grid(row=5, column=0, sticky="ew")

    def _path_row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=5, padx=(0, 12))
        ttk.Entry(parent, textvariable=variable, width=88).grid(row=row, column=1, sticky="ew", pady=5, padx=(0, 10))
        button = ttk.Button(parent, text="Browse...", command=command)
        button.grid(row=row, column=2, sticky="e", pady=5)
        self._browse_buttons.append(button)

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
            self.log("Warning: logo asset not found at assets/vectis.png; continuing without logo.")

    def log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status.see(tk.END)

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
        if self._processing:
            self.log("Processing is already running; duplicate request ignored.")
            return
        if not self.input_file.get():
            messagebox.showerror("Vectis", "Please select an input CSV/XLSX file.")
            return
        self._set_processing_controls(True)
        self._started_at = time.monotonic()
        self._last_stage = "Reading input file"
        self._set_progress("Reading input file", 5)
        self.state_text.set("Engine: active · 00:00:00")
        self.log(f"Processing: {self.input_file.get()}")

        def progress_callback(stage: str, percent: int) -> None:
            self._progress_queue.put(("progress", (stage, percent)))

        def worker() -> None:
            try:
                result = triage_file(self.input_file.get(), self.output_folder.get(), self.reference_folder.get(), progress_callback=progress_callback)
                self._progress_queue.put(("complete", result))
            except Exception as exc:
                self._progress_queue.put(("error", exc))

        self._worker_thread = threading.Thread(target=worker, name="VectisTriageWorker", daemon=True)
        self._worker_thread.start()
        self.after(PROGRESS_POLL_MS, self._poll_progress_queue)
        self.after(HEARTBEAT_MS, self._heartbeat)

    def _set_progress(self, stage: str, percent: int) -> None:
        self._last_stage = stage
        self.progress_value.set(percent)
        self.progress_text.set(f"{stage}... {percent}%" if percent < 100 else f"{stage} {percent}%")

    def _set_processing_controls(self, processing: bool) -> None:
        self._processing = processing
        state = tk.DISABLED if processing else tk.NORMAL
        self.process_button.configure(state=state)
        for button in self._browse_buttons:
            button.configure(state=state)

    def _poll_progress_queue(self) -> None:
        try:
            while True:
                event, payload = self._progress_queue.get_nowait()
                if event == "progress":
                    stage, percent = payload  # type: ignore[misc]
                    self._set_progress(str(stage), int(percent))
                    self.log(f"{stage}...")
                elif event == "complete":
                    self._handle_complete(payload)  # type: ignore[arg-type]
                    return
                elif event == "error":
                    self._handle_error(payload)  # type: ignore[arg-type]
                    return
        except queue.Empty:
            pass
        if self._processing:
            self.after(PROGRESS_POLL_MS, self._poll_progress_queue)

    def _heartbeat(self) -> None:
        if not self._processing:
            return
        elapsed = int(time.monotonic() - self._started_at)
        hh, remainder = divmod(elapsed, 3600)
        mm, ss = divmod(remainder, 60)
        self.state_text.set(f"Engine: active · {hh:02d}:{mm:02d}:{ss:02d}")
        self.after(HEARTBEAT_MS, self._heartbeat)

    def _handle_complete(self, result: TriageResult) -> None:
        for warning in result.warnings:
            self.log(f"Warning: {warning}")
        self._set_progress("Complete", 100)
        self.log(f"Complete: {result.output_path}")
        self.log(f"Rows: total={result.total_rows}, extracted={result.extracted_rows}, remainder={result.remainder_rows}")
        self.state_text.set("Engine: complete")
        self._set_processing_controls(False)
        messagebox.showinfo("Vectis", f"Triaged workbook created:\n{result.output_path}")

    def _handle_error(self, exc: object) -> None:
        self.state_text.set("Engine: error")
        self.log(f"Error: {exc}")
        self._set_processing_controls(False)
        messagebox.showerror("Vectis", str(exc))

    def open_output_folder(self) -> None:
        try:
            self.state_text.set("Engine: opening output folder")
            open_folder(self.output_folder.get())
            self.state_text.set("Engine: ready")
        except Exception as exc:
            self.state_text.set("Engine: error")
            self.log(f"Error opening folder: {exc}")
            messagebox.showerror("Vectis", str(exc))


if __name__ == "__main__":
    VectisApp().mainloop()
