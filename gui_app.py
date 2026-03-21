"""Windows minimal-theme GUI for QAM roster sync."""

from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from modules import calendar_api
from modules.sync_service import (
    DEFAULT_ALL_WORKERS_TITLE,
    MODE_ADD_ONLY,
    MODE_DELETE_CURRENT,
    MODE_DELETE_MONTH_ALL,
    MODE_DELETE_ONLY,
    MODE_FULL,
    SyncOptions,
    execute_sync_plan,
    prepare_sync_plan,
    summarize_plan,
)


MODE_LABELS = {
    "Full sync (delete old + add new)": MODE_FULL,
    "Delete previous version only": MODE_DELETE_ONLY,
    "Delete current version (from selected DOCX)": MODE_DELETE_CURRENT,
    "Delete all versions in chosen month": MODE_DELETE_MONTH_ALL,
    "Add new version only": MODE_ADD_ONLY,
}

APP_ICON_FILE = "QAM-Logo-1-2048x1310whiteBGRND.png"


class RosterSyncApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("QAM Roster Automation")
        self.geometry("980x760")
        self.minsize(880, 660)
        self._set_theme()
        self._set_window_icon()

        now = datetime.now()
        self.docx_var = tk.StringVar()
        self.calendar_var = tk.StringVar()
        self.all_workers_enabled_var = tk.BooleanVar(value=False)
        self.all_workers_calendar_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="Full sync (delete old + add new)")
        self.dry_run_var = tk.BooleanVar(value=True)
        self.replace_current_var = tk.BooleanVar(value=False)
        self.timezone_var = tk.StringVar(value="Australia/Brisbane")
        self.month_var = tk.IntVar(value=now.month)
        self.year_var = tk.IntVar(value=now.year)
        self.volunteer_var = tk.StringVar(value="Wayne Freestun")
        self.event_title_var = tk.StringVar(value="Wayne volunteer Queensland Air Museum")
        self.all_workers_event_title_var = tk.StringVar(value=DEFAULT_ALL_WORKERS_TITLE)
        self.location_var = tk.StringVar(value="Queensland Air Museum")

        self.calendars: list[dict] = []
        self.current_plan = None

        self._build_ui()
        self.after(100, self._load_calendars)

    def _set_theme(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")

    def _set_window_icon(self) -> None:
        icon_path = Path(__file__).resolve().parent / APP_ICON_FILE
        if not icon_path.exists():
            return
        try:
            self._icon_image = tk.PhotoImage(file=str(icon_path))
            self.iconphoto(True, self._icon_image)
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=14)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(12, weight=1)

        ttk.Label(frame, text="Roster DOCX").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.docx_entry = ttk.Entry(frame, textvariable=self.docx_var)
        self.docx_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        self.docx_browse_btn = ttk.Button(frame, text="Browse...", command=self._pick_file)
        self.docx_browse_btn.grid(row=0, column=2, padx=(8, 0), pady=(0, 8))

        ttk.Label(frame, text="Wayne Calendar").grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.calendar_combo = ttk.Combobox(frame, textvariable=self.calendar_var, state="readonly")
        self.calendar_combo.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(frame, text="Refresh List", command=self._load_calendars).grid(
            row=1, column=2, padx=(8, 0), pady=(0, 8)
        )

        self.all_workers_check = ttk.Checkbutton(
            frame,
            text="Also sync All Workers calendar",
            variable=self.all_workers_enabled_var,
            command=self._on_all_workers_toggled,
        )
        self.all_workers_check.grid(row=2, column=1, sticky="w", pady=(0, 8))

        ttk.Label(frame, text="All Workers Calendar").grid(row=3, column=0, sticky="w", pady=(0, 8))
        self.all_workers_calendar_combo = ttk.Combobox(
            frame,
            textvariable=self.all_workers_calendar_var,
            state="readonly",
        )
        self.all_workers_calendar_combo.grid(row=3, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Action").grid(row=4, column=0, sticky="w", pady=(0, 8))
        mode_combo = ttk.Combobox(frame, textvariable=self.mode_var, state="readonly", values=list(MODE_LABELS.keys()))
        mode_combo.grid(row=4, column=1, sticky="ew", pady=(0, 8))
        mode_combo.bind("<<ComboboxSelected>>", self._on_mode_changed)

        ttk.Label(frame, text="Delete Month/Year").grid(row=5, column=0, sticky="w", pady=(0, 8))
        month_row = ttk.Frame(frame)
        month_row.grid(row=5, column=1, sticky="w", pady=(0, 8))
        ttk.Spinbox(month_row, from_=1, to=12, width=6, textvariable=self.month_var).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Spinbox(month_row, from_=2000, to=2100, width=8, textvariable=self.year_var).pack(side=tk.LEFT)

        checks = ttk.Frame(frame)
        checks.grid(row=6, column=1, sticky="w", pady=(0, 8))
        ttk.Checkbutton(checks, text="Dry run", variable=self.dry_run_var).pack(side=tk.LEFT, padx=(0, 16))
        self.replace_current_check = ttk.Checkbutton(
            checks,
            text="Include current version in delete (full sync only)",
            variable=self.replace_current_var,
        )
        self.replace_current_check.pack(side=tk.LEFT)

        ttk.Label(frame, text="Timezone").grid(row=7, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(frame, textvariable=self.timezone_var).grid(row=7, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Volunteer").grid(row=8, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(frame, textvariable=self.volunteer_var).grid(row=8, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Wayne Event Title").grid(row=9, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(frame, textvariable=self.event_title_var).grid(row=9, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="All Workers Title").grid(row=10, column=0, sticky="w", pady=(0, 8))
        self.all_workers_title_entry = ttk.Entry(frame, textvariable=self.all_workers_event_title_var)
        self.all_workers_title_entry.grid(row=10, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Location").grid(row=11, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(frame, textvariable=self.location_var).grid(row=11, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Identified Changes").grid(row=12, column=0, sticky="nw")
        self.preview_text = tk.Text(frame, height=14, wrap="word")
        self.preview_text.grid(row=12, column=1, columnspan=2, sticky="nsew")

        btns = ttk.Frame(frame)
        btns.grid(row=13, column=1, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Preview Changes", command=self._preview).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns, text="Proceed", command=self._proceed).pack(side=tk.LEFT)
        ttk.Label(frame, text=f"Build v{self._load_build_version()}").grid(row=13, column=0, sticky="w", pady=(10, 0))

        self._on_mode_changed()
        self._on_all_workers_toggled()

    def _load_build_version(self) -> str:
        project_file = Path(__file__).resolve().parent / "project.json"
        try:
            data = json.loads(project_file.read_text(encoding="utf-8"))
            version = data.get("version")
            if isinstance(version, str) and version.strip():
                return version.strip()
        except Exception:  # noqa: BLE001
            pass
        return "unknown"

    def _pick_file(self) -> None:
        from main import prompt_for_docx_file

        selected = prompt_for_docx_file()
        if selected:
            self.docx_var.set(selected)

    def _load_calendars(self) -> None:
        try:
            self.calendars = calendar_api.list_calendars()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Calendar Load Error", str(exc))
            return

        if not self.calendars:
            self.calendar_combo["values"] = []
            self.all_workers_calendar_combo["values"] = []
            self.calendar_var.set("")
            self.all_workers_calendar_var.set("")
            messagebox.showwarning("No Calendars", "No calendars were returned for this account.")
            return

        values = [self._calendar_label(c) for c in self.calendars]
        current_wayne = self.calendar_var.get()
        current_all_workers = self.all_workers_calendar_var.get()
        self.calendar_combo["values"] = values
        self.all_workers_calendar_combo["values"] = values
        if current_wayne in values:
            self.calendar_combo.current(values.index(current_wayne))
        else:
            self.calendar_combo.current(0)
        if current_all_workers in values:
            self.all_workers_calendar_combo.current(values.index(current_all_workers))
        elif len(values) > 1:
            self.all_workers_calendar_combo.current(1)
        else:
            self.all_workers_calendar_combo.current(0)

    def _calendar_label(self, cal: dict) -> str:
        summary = cal.get("summary", "<no title>")
        calendar_id = cal.get("id", "<no id>")
        primary = " (primary)" if cal.get("primary") else ""
        return f"{summary}{primary} [{calendar_id}]"

    def _selected_calendar_id(self) -> str:
        if not self.calendar_var.get():
            return "primary"
        idx = self.calendar_combo.current()
        if idx < 0 or idx >= len(self.calendars):
            return "primary"
        return str(self.calendars[idx].get("id", "primary"))

    def _selected_all_workers_calendar_id(self) -> str:
        if not self.all_workers_calendar_var.get():
            return "primary"
        idx = self.all_workers_calendar_combo.current()
        if idx < 0 or idx >= len(self.calendars):
            return "primary"
        return str(self.calendars[idx].get("id", "primary"))

    def _selected_mode(self) -> str:
        return MODE_LABELS.get(self.mode_var.get(), MODE_FULL)

    def _requires_docx(self, mode: str) -> bool:
        return mode != MODE_DELETE_MONTH_ALL

    def _on_mode_changed(self, _event=None) -> None:
        mode = self._selected_mode()
        if mode == MODE_FULL:
            self.replace_current_check.state(["!disabled"])
        else:
            self.replace_current_check.state(["disabled"])

        if mode == MODE_DELETE_MONTH_ALL:
            self.docx_entry.state(["disabled"])
            self.docx_browse_btn.state(["disabled"])
            self.all_workers_check.state(["disabled"])
            self.all_workers_enabled_var.set(False)
        else:
            self.docx_entry.state(["!disabled"])
            self.docx_browse_btn.state(["!disabled"])
            self.all_workers_check.state(["!disabled"])

        self._on_all_workers_toggled()

    def _on_all_workers_toggled(self) -> None:
        enabled = self.all_workers_enabled_var.get() and self._selected_mode() != MODE_DELETE_MONTH_ALL
        if enabled:
            self.all_workers_calendar_combo.state(["!disabled"])
            self.all_workers_title_entry.state(["!disabled"])
        else:
            self.all_workers_calendar_combo.state(["disabled"])
            self.all_workers_title_entry.state(["disabled"])

    def _build_options(self) -> SyncOptions:
        mode = self._selected_mode()
        docx_value = self.docx_var.get().strip() if self._requires_docx(mode) else None
        return SyncOptions(
            docx_file=docx_value,
            calendar_id=self._selected_calendar_id(),
            timezone=self.timezone_var.get().strip() or "Australia/Brisbane",
            dry_run=self.dry_run_var.get(),
            replace_current=self.replace_current_var.get(),
            mode=mode,
            month=int(self.month_var.get()),
            year=int(self.year_var.get()),
            volunteer_name=self.volunteer_var.get().strip() or "Wayne Freestun",
            event_title=self.event_title_var.get().strip() or "Wayne volunteer Queensland Air Museum",
            location=self.location_var.get().strip() or "Queensland Air Museum",
            sync_all_workers=self.all_workers_enabled_var.get(),
            all_workers_calendar_id=self._selected_all_workers_calendar_id() if self.all_workers_enabled_var.get() else None,
            all_workers_event_title=self.all_workers_event_title_var.get().strip() or DEFAULT_ALL_WORKERS_TITLE,
        )

    def _preview(self) -> None:
        mode = self._selected_mode()
        if self._requires_docx(mode) and not self.docx_var.get().strip():
            messagebox.showwarning("Missing DOCX", "Select a DOCX file first.")
            return
        if mode == MODE_DELETE_MONTH_ALL:
            month = int(self.month_var.get())
            if month < 1 or month > 12:
                messagebox.showwarning("Invalid Month", "Month must be between 1 and 12.")
                return
        if self.all_workers_enabled_var.get() and self._selected_all_workers_calendar_id() == self._selected_calendar_id():
            messagebox.showwarning(
                "Separate Calendar Required",
                "Choose a different calendar for the optional All Workers sync.",
            )
            return

        try:
            self.current_plan = prepare_sync_plan(self._build_options())
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Preview Error", str(exc))
            return

        summary = summarize_plan(self.current_plan)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", summary)

    def _proceed(self) -> None:
        if self.current_plan is None:
            self._preview()
            if self.current_plan is None:
                return

        confirmation_text = summarize_plan(self.current_plan)
        if not messagebox.askyesno("Confirm Changes", f"Proceed with these changes?\n\n{confirmation_text}"):
            return

        try:
            result = execute_sync_plan(self.current_plan)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Execution Error", str(exc))
            return

        messagebox.showinfo(
            "Done",
            f"Completed.\nDeleted: {result.deleted_count}\nCreated: {result.created_count}\n"
            f"Dry run: {self.current_plan.options.dry_run}",
        )


def run_gui() -> None:
    app = RosterSyncApp()
    app.mainloop()
