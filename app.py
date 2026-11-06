from __future__ import annotations

import queue
import threading
import time
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, List, Optional, Set, Tuple
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import pythoncom
import pywintypes
import win32com.client


APP_NAME = "AeroXL to PDF"
GITHUB_URL = "https://github.com/JekoGeek"
SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm"}
XL_TYPE_PDF = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
RETRYABLE_COM_ERRORS = {-2147418111, -2147417846}


@dataclass(frozen=True)
class ConversionItem:
    source: Path
    output: Path


def discover_excel_files(source_root: Path) -> List[Path]:
    files = [
        path
        for path in source_root.rglob("*")
        if path.is_file()
        and not path.name.startswith("~$")
        and path.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    return sorted(files, key=lambda item: str(item).lower())


def build_plan(source_root: Path, target_root: Path, files: List[Path]) -> List[ConversionItem]:
    items: List[ConversionItem] = []
    used: Set[Path] = set()
    for source in files:
        parent = source.relative_to(source_root).parent
        relative_output = parent / (source.stem + ".pdf")
        if relative_output in used:
            label = source.suffix.lower().replace(".", "") or "file"
            relative_output = parent / (source.stem + "__" + label + ".pdf")
            number = 2
            while relative_output in used:
                relative_output = parent / (source.stem + "__" + label + "_" + str(number) + ".pdf")
                number += 1
        used.add(relative_output)
        items.append(ConversionItem(source, target_root / relative_output))
    return items


class ExcelPdfConverter:
    def __init__(self) -> None:
        self.excel = None

    def __enter__(self) -> "ExcelPdfConverter":
        pythoncom.CoInitialize()
        try:
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
            self.excel.EnableEvents = False
            self.excel.AskToUpdateLinks = False
            try:
                self.excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:
                pass
            return self
        except Exception:
            pythoncom.CoUninitialize()
            raise

    def __exit__(self, exc_type, exc, traceback) -> None:
        if self.excel is not None:
            try:
                self._retry(self.excel.Quit, swallow=True)
            finally:
                self.excel = None
        pythoncom.CoUninitialize()

    def convert(self, source: Path, output: Path) -> None:
        output.parent.mkdir(parents=True, exist_ok=True)
        workbook = None
        try:
            workbook = self._retry(
                self.excel.Workbooks.Open,
                str(source),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
            )
        except Exception:
            workbook = self._retry(
                self.excel.Workbooks.Open,
                str(source),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
            )
        try:
            self._retry(
                workbook.ExportAsFixedFormat,
                Type=XL_TYPE_PDF,
                Filename=str(output),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        finally:
            if workbook is not None:
                self._retry(workbook.Close, SaveChanges=False, swallow=True)

    @staticmethod
    def _retry(func: Callable, *args, **kwargs):
        swallow = kwargs.pop("swallow", False)
        last_error = None
        for attempt in range(8):
            try:
                return func(*args, **kwargs)
            except pywintypes.com_error as error:
                last_error = error
                hresult = error.args[0] if error.args else None
                if hresult not in RETRYABLE_COM_ERRORS:
                    raise
                time.sleep(0.35 + attempt * 0.08)
        if not swallow and last_error is not None:
            raise last_error
        return None


class AeroButton(tk.Canvas):
    def __init__(self, parent, text: str, command: Callable, primary: bool = False) -> None:
        self.command = command
        self.primary = primary
        self.normal_bg = "#16b987" if primary else "#ecfbff"
        self.hover_bg = "#0da06f" if primary else "#d5f4ff"
        self.normal_fg = "#ffffff" if primary else "#0c5575"
        self.text = text
        self.enabled = True
        super().__init__(
            parent,
            width=142 if primary else 106,
            height=42,
            bg=parent.cget("bg"),
            highlightthickness=0,
            cursor="hand2",
        )
        self._paint(self.normal_bg, self.normal_fg)
        self.bind("<Button-1>", self._invoke)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)

    def _paint(self, fill: str, foreground: str) -> None:
        self.delete("all")
        width = int(self.cget("width"))
        height = int(self.cget("height"))
        radius = height // 2
        self.create_oval(0, 0, height, height, fill=fill, outline="")
        self.create_oval(width - height, 0, width, height, fill=fill, outline="")
        self.create_rectangle(radius, 0, width - radius, height, fill=fill, outline="")
        self.create_text(width // 2, height // 2, text=self.text, fill=foreground, font=("Segoe UI Semibold", 10))

    def _invoke(self, event=None) -> None:
        if self.enabled:
            self.command()

    def _enter(self, event=None) -> None:
        if self.enabled:
            self._paint(self.hover_bg, self.normal_fg)

    def _leave(self, event=None) -> None:
        if self.enabled:
            self._paint(self.normal_bg, self.normal_fg)

    def set_enabled(self, enabled: bool) -> None:
        if enabled:
            self.enabled = True
            self.configure(cursor="hand2")
            self._paint(self.normal_bg, self.normal_fg)
        else:
            self.enabled = False
            self.configure(cursor="arrow")
            self._paint("#b8d4d6", "#edf8f7")


class AeroSignature(tk.Canvas):
    def __init__(self, parent) -> None:
        super().__init__(parent, width=230, height=42, bg=parent.cget("bg"), highlightthickness=0, cursor="hand2")
        self._paint(False)
        self.bind("<Button-1>", lambda event: webbrowser.open_new_tab(GITHUB_URL))
        self.bind("<Enter>", lambda event: self._paint(True))
        self.bind("<Leave>", lambda event: self._paint(False))

    def _paint(self, active: bool) -> None:
        fill = "#d6f7ff" if active else "#eafcff"
        self.delete("all")
        self.create_oval(0, 0, 42, 42, fill=fill, outline="")
        self.create_oval(188, 0, 230, 42, fill=fill, outline="")
        self.create_rectangle(21, 0, 209, 42, fill=fill, outline="")
        self.create_text(64, 21, text="Made by JEKO", fill="#17435a", font=("Segoe UI Semibold", 9))
        self.create_line(130, 10, 130, 32, fill="#a9deeb", width=1)
        self.create_text(175, 21, text="My GitHub", fill="#0879ad", font=("Segoe UI", 9, "underline"))


class ConverterApp:
    SKY = "#dff7ff"
    WATER = "#69d8ee"
    DEEP_WATER = "#1a94bc"
    WHITE = "#ffffff"
    INK = "#17435a"
    MUTED = "#5d8293"
    MINT = "#16b987"
    CLOUD = "#f4fdff"
    BORDER = "#a9deeb"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("980x720")
        self.root.minsize(860, 640)
        self.root.configure(bg=self.SKY)

        self.events: queue.Queue = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.running = False
        self.total_files = 0
        self.completed_files = 0
        self.failed_files = 0
        self.source_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Choose two folders to start your conversion.")
        self.detail_var = tk.StringVar(value="Ready when you are.")
        self.counter_var = tk.StringVar(value="0 / 0")
        self.summary_var = tk.StringVar(value="Awaiting files")
        self.log_lines: List[str] = []

        self._create_ui()
        self.root.after(120, self._poll_events)

    def _create_ui(self) -> None:
        # A regular frame layout is deliberately used here. Canvas-embedded widgets
        # can receive a zero height on older Tk builds, leaving a blank window.
        self.content = tk.Frame(self.root, bg=self.SKY)
        self.content.pack(fill="both", expand=True)

        top = tk.Frame(self.content, bg=self.SKY, padx=42, pady=8)
        top.pack(fill="x")
        top.columnconfigure(0, weight=1)

        brand = tk.Frame(top, bg=self.SKY)
        brand.grid(row=0, column=0, sticky="w")
        tk.Label(brand, text="AeroXL", bg=self.SKY, fg=self.DEEP_WATER, font=("Segoe UI Semibold", 25)).pack(side="left")
        tk.Label(brand, text="to PDF", bg=self.SKY, fg=self.MINT, font=("Segoe UI Semibold", 25)).pack(side="left", padx=(7, 0))
        tk.Label(top, text="BATCH EXCEL CONVERTER", bg=self.SKY, fg="#34849c", font=("Segoe UI Semibold", 9)).grid(
            row=1, column=0, sticky="w", pady=(5, 0)
        )

        AeroSignature(top).grid(row=0, column=1, rowspan=2, sticky="e")

        card = tk.Frame(self.content, bg=self.WHITE, highlightbackground=self.BORDER, highlightthickness=1, padx=24, pady=10)
        card.pack(fill="both", expand=True, padx=42, pady=(2, 8))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(6, weight=1)

        intro = tk.Frame(card, bg=self.WHITE)
        intro.grid(row=0, column=0, sticky="ew")
        tk.Label(intro, text="A clear path from spreadsheets to PDFs.", bg=self.WHITE, fg=self.INK, font=("Segoe UI Semibold", 15)).pack(anchor="w")
        tk.Label(
            intro,
            text="AeroXL preserves your folder structure and lets Microsoft Excel do the rendering.",
            bg=self.WHITE,
            fg=self.MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(5, 0))

        self._folder_row(card, 1, "1  Source folder", "Where your Excel files live", self.source_var, self.choose_source)
        self._folder_row(card, 2, "2  PDF destination", "Where converted PDFs will be saved", self.target_var, self.choose_target)

        info = tk.Frame(card, bg="#eefdff", highlightbackground="#c0e9f1", highlightthickness=1, padx=16, pady=6)
        info.grid(row=3, column=0, sticky="ew", pady=(2, 8))
        tk.Label(info, text="Works with", bg="#eefdff", fg=self.DEEP_WATER, font=("Segoe UI Semibold", 9)).pack(side="left")
        tk.Label(info, text="  XLS   XLSX   XLSM   XLSB   XLT   XLTX   XLTM", bg="#eefdff", fg=self.MUTED, font=("Segoe UI", 9)).pack(side="left")

        status = tk.Frame(card, bg="#e8fbf5", highlightbackground="#ace4d2", highlightthickness=1, padx=18, pady=10)
        status.grid(row=4, column=0, sticky="ew")
        status.columnconfigure(0, weight=1)
        tk.Label(status, text="CONVERSION STATUS", bg="#e8fbf5", fg="#16805f", font=("Segoe UI Semibold", 9)).grid(row=0, column=0, sticky="w")
        tk.Label(status, textvariable=self.counter_var, bg="#e8fbf5", fg="#138767", font=("Segoe UI Semibold", 18)).grid(row=0, column=1, sticky="e")
        tk.Label(status, textvariable=self.status_var, bg="#e8fbf5", fg=self.INK, font=("Segoe UI Semibold", 10), anchor="w").grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 11))
        self.progress = ttk.Progressbar(status, mode="determinate", maximum=100, value=0, style="Aero.Horizontal.TProgressbar")
        self.progress.grid(row=2, column=0, columnspan=2, sticky="ew")
        tk.Label(status, textvariable=self.detail_var, bg="#e8fbf5", fg=self.MUTED, font=("Segoe UI", 9), anchor="w").grid(row=3, column=0, sticky="w", pady=(7, 0))
        tk.Label(status, textvariable=self.summary_var, bg="#e8fbf5", fg=self.MUTED, font=("Segoe UI", 9), anchor="e").grid(row=3, column=1, sticky="e", pady=(7, 0))

        log_frame = tk.Frame(card, bg=self.WHITE)
        log_frame.grid(row=5, column=0, sticky="nsew", pady=(7, 0))
        tk.Label(log_frame, text="Activity", bg=self.WHITE, fg=self.INK, font=("Segoe UI Semibold", 10)).pack(anchor="w", pady=(0, 7))
        self.log = ScrolledText(log_frame, height=2, bg="#f6fdff", fg="#315d70", relief="flat", bd=1, highlightbackground="#c6eaf1", highlightthickness=1, font=("Consolas", 9), padx=10, pady=7, wrap="word")
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")
        self._render_log()

        actions = tk.Frame(card, bg=self.WHITE)
        actions.grid(row=6, column=0, sticky="ew", pady=(8, 0))
        self.convert_button = AeroButton(actions, "Convert to PDF", self.start_conversion, primary=True)
        self.convert_button.pack(side="left")
        tk.Label(actions, text="No administrator rights needed.", bg=self.WHITE, fg=self.MUTED, font=("Segoe UI", 9)).pack(side="left", padx=(16, 0))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Aero.Horizontal.TProgressbar", troughcolor="#ccefe5", background="#16b987", lightcolor="#16b987", darkcolor="#16b987", bordercolor="#ccefe5", thickness=12)

    def _folder_row(self, parent, row: int, title: str, hint: str, variable: tk.StringVar, command: Callable) -> None:
        block = tk.Frame(parent, bg=self.WHITE)
        block.grid(row=row, column=0, sticky="ew", pady=(10 if row == 1 else 8, 0))
        block.columnconfigure(0, weight=1)
        tk.Label(block, text=title, bg=self.WHITE, fg=self.INK, font=("Segoe UI Semibold", 10)).grid(row=0, column=0, sticky="w")
        tk.Label(block, text=hint, bg=self.WHITE, fg=self.MUTED, font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=(2, 4))
        entry = tk.Entry(block, textvariable=variable, bg="#f9feff", fg=self.INK, relief="flat", highlightbackground=self.BORDER, highlightcolor=self.DEEP_WATER, highlightthickness=1, font=("Segoe UI", 10), insertbackground=self.INK)
        entry.grid(row=2, column=0, sticky="ew", ipady=9)
        button = AeroButton(block, "Browse", command)
        button.grid(row=2, column=1, sticky="e", padx=(12, 0))
        if row == 1:
            self.source_entry, self.source_button = entry, button
        else:
            self.target_entry, self.target_button = entry, button

    def choose_source(self) -> None:
        choice = filedialog.askdirectory(title="Choose a folder with Excel files")
        if choice:
            self.source_var.set(choice)

    def choose_target(self) -> None:
        choice = filedialog.askdirectory(title="Choose a destination for PDF files")
        if choice:
            self.target_var.set(choice)

    def start_conversion(self) -> None:
        if self.running:
            return
        source = Path(self.source_var.get().strip())
        target = Path(self.target_var.get().strip())
        if not source.is_dir():
            messagebox.showerror(APP_NAME, "Choose an existing source folder with Excel files.")
            return
        if not target.is_dir():
            messagebox.showerror(APP_NAME, "Choose an existing folder for PDFs.")
            return
        self.running = True
        self.total_files = self.completed_files = self.failed_files = 0
        self.progress.configure(value=0, maximum=100)
        self.counter_var.set("0 / 0")
        self.status_var.set("Looking for Excel files...")
        self.detail_var.set("Preparing your conversion queue.")
        self.summary_var.set("Scanning")
        self.log_lines = []
        self._render_log()
        self._set_enabled(False)
        self.worker = threading.Thread(target=self._convert_worker, args=(source, target), daemon=True)
        self.worker.start()

    def _convert_worker(self, source: Path, target: Path) -> None:
        try:
            plan = build_plan(source, target, discover_excel_files(source))
            self.events.put(("scan", len(plan)))
            if not plan:
                self.events.put(("done", 0))
                return
            with ExcelPdfConverter() as converter:
                for index, item in enumerate(plan, 1):
                    self.events.put(("current", index, len(plan), item.source.name))
                    try:
                        converter.convert(item.source, item.output)
                        self.events.put(("success", index, len(plan), item.source.name, item.output.name))
                    except Exception as error:
                        self.events.put(("error", index, len(plan), item.source.name, str(error)))
            self.events.put(("done", len(plan)))
        except Exception as error:
            self.events.put(("fatal", str(error)))

    def _poll_events(self) -> None:
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break
            self._handle_event(event)
        self.root.after(120, self._poll_events)

    def _handle_event(self, event: Tuple) -> None:
        kind = event[0]
        if kind == "scan":
            self.total_files = event[1]
            self.counter_var.set("0 / " + str(self.total_files))
            self.summary_var.set("Found: " + str(self.total_files))
            if self.total_files:
                self.progress.configure(maximum=self.total_files, value=0)
                self.status_var.set("Converting your Excel files...")
            else:
                self.status_var.set("No supported Excel files were found.")
                self.detail_var.set("Try another source folder.")
        elif kind == "current":
            index, total, name = event[1:]
            self.counter_var.set(str(min(index - 1, total)) + " / " + str(total))
            self.detail_var.set("Now converting: " + name)
        elif kind == "success":
            index, total, source, output = event[1:]
            self.completed_files += 1
            self.progress.configure(value=index)
            self.counter_var.set(str(index) + " / " + str(total))
            self.detail_var.set("Ready: " + source)
            self.summary_var.set("PDF: " + str(self.completed_files) + "  Errors: " + str(self.failed_files))
            self._append_log("OK   " + source + " -> " + output)
        elif kind == "error":
            index, total, source, error = event[1:]
            self.failed_files += 1
            self.progress.configure(value=index)
            self.counter_var.set(str(index) + " / " + str(total))
            self.status_var.set("Some files could not be converted.")
            self.detail_var.set("Problem file: " + source)
            self.summary_var.set("PDF: " + str(self.completed_files) + "  Errors: " + str(self.failed_files))
            self._append_log("ERR  " + source + " -> " + error)
        elif kind == "done":
            total = event[1]
            self.running = False
            self.progress.configure(value=total if total else 0)
            self.counter_var.set(str(total) + " / " + str(total))
            if not total:
                self.status_var.set("Done. No Excel files to convert.")
                self.detail_var.set("Choose another source folder and try again.")
                self.summary_var.set("0 files")
            elif self.failed_files:
                self.status_var.set("Finished with a few errors.")
                self.detail_var.set("Check Activity for the files that need attention.")
                self.summary_var.set("PDF: " + str(self.completed_files) + "  Errors: " + str(self.failed_files))
                messagebox.showwarning(APP_NAME, "Finished.\n\nConverted: " + str(self.completed_files) + "\nErrors: " + str(self.failed_files))
            else:
                self.status_var.set("All done. Your PDFs are ready.")
                self.detail_var.set("Every Excel file was converted successfully.")
                self.summary_var.set("Converted: " + str(self.completed_files))
                messagebox.showinfo(APP_NAME, "All done.\n\nConverted files: " + str(self.completed_files))
            self._set_enabled(True)
        elif kind == "fatal":
            self.running = False
            self.status_var.set("Conversion could not start.")
            self.detail_var.set("Make sure Microsoft Excel is installed and the selected folders are available.")
            self.summary_var.set("Stopped")
            self._append_log("FATAL  " + event[1])
            self._set_enabled(True)
            messagebox.showerror(APP_NAME, "Could not start conversion.\n\n" + event[1])

    def _set_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.source_entry.configure(state=state)
        self.target_entry.configure(state=state)
        self.source_button.set_enabled(enabled)
        self.target_button.set_enabled(enabled)
        self.convert_button.set_enabled(enabled)

    def _append_log(self, line: str) -> None:
        self.log_lines.append(line)
        self.log_lines = self.log_lines[-120:]
        self._render_log()

    def _render_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("1.0", "\n".join(self.log_lines) if self.log_lines else "Activity will appear here.")
        self.log.configure(state="disabled")


def main() -> None:
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
