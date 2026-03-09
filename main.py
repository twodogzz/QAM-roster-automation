"""Entry point for syncing QAM roster DOCX files to Google Calendar."""

from __future__ import annotations

from argparse import ArgumentParser
from pathlib import Path
import logging
import os
import sys
from tkinter import TclError, Tk, filedialog

from modules import calendar_api
from modules.sync_service import (
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


LOGGER = logging.getLogger("qam_roster_automation")


def main() -> int:
    args = parse_args()
    configure_logging()

    if should_launch_gui(args):
        from gui_app import run_gui

        run_gui()
        return 0

    mode = resolve_mode(args)
    docx_file = resolve_docx_file(args, mode)
    calendar_id = resolve_calendar_id(args)

    options = SyncOptions(
        docx_file=docx_file,
        calendar_id=calendar_id,
        timezone=args.timezone,
        dry_run=args.dry_run,
        replace_current=args.replace_current,
        mode=mode,
        month=args.month,
        year=args.year,
        volunteer_name=args.volunteer_name,
        event_title=args.event_title,
        location=args.location,
    )

    plan = prepare_sync_plan(options)
    LOGGER.info("\n%s", summarize_plan(plan))
    result = execute_sync_plan(plan)
    LOGGER.info("Run complete. Deleted=%d Created=%d", result.deleted_count, result.created_count)
    return 0


def parse_args():
    parser = ArgumentParser(description="Sync QAM roster DOCX to Google Calendar")
    parser.add_argument(
        "docx_file",
        nargs="?",
        help="Path to the source roster DOCX file (or drag and drop onto the script)",
    )
    parser.add_argument("--gui", action="store_true", help="Launch the Windows GUI")
    parser.add_argument("--cli", action="store_true", help="Force CLI mode (disable auto GUI launch)")
    parser.add_argument("--select-file", action="store_true", help="Open a file-picker UI for the DOCX file")
    parser.add_argument(
        "--no-file-picker",
        action="store_true",
        help="Disable automatic file-picker fallback when no DOCX path is provided",
    )
    parser.add_argument("--calendar-id", default=None, help="Google Calendar ID (default: prompt/primary)")
    parser.add_argument("--select-calendar", action="store_true", help="Show interactive calendar picker")
    parser.add_argument(
        "--no-calendar-prompt",
        action="store_true",
        help="Disable interactive calendar picker and fallback to primary when not configured",
    )
    parser.add_argument("--timezone", default="Australia/Brisbane", help="IANA timezone for event timestamps")
    parser.add_argument("--dry-run", action="store_true", help="Preview only; do not write to Google Calendar")
    parser.add_argument("--replace-current", action="store_true", help="Delete current version before add")
    parser.add_argument("--delete-only", action="store_true", help="Delete previous versions only")
    parser.add_argument("--delete-current", action="store_true", help="Delete only the version matched from DOCX")
    parser.add_argument("--delete-month-all", action="store_true", help="Delete all versions in a chosen month")
    parser.add_argument("--month", type=int, default=None, help="Month number for --delete-month-all (1-12)")
    parser.add_argument("--year", type=int, default=None, help="Year for --delete-month-all")
    parser.add_argument("--add-only", action="store_true", help="Add new version only")
    parser.add_argument("--volunteer-name", default="Wayne Freestun", help="Volunteer name to match in roster")
    parser.add_argument("--event-title", default="Wayne volunteer Queensland Air Museum", help="Event title prefix")
    parser.add_argument("--location", default="Queensland Air Museum", help="Google Calendar event location")
    return parser.parse_args()


def should_launch_gui(args) -> bool:
    if args.cli:
        return False
    if args.gui:
        return True
    # Default behavior for desktop use: open GUI on Windows when launched without workflow args.
    if os.name != "nt":
        return False
    has_cli_work = bool(
        args.docx_file
        or args.select_file
        or args.calendar_id
        or args.select_calendar
        or args.delete_only
        or args.delete_current
        or args.delete_month_all
        or args.add_only
        or args.dry_run
    )
    return not has_cli_work


def resolve_mode(args) -> str:
    selected = [args.delete_only, args.add_only, args.delete_current, args.delete_month_all]
    if sum(1 for x in selected if x) > 1:
        raise ValueError("Choose only one action mode flag.")
    if args.delete_only:
        return MODE_DELETE_ONLY
    if args.delete_current:
        return MODE_DELETE_CURRENT
    if args.delete_month_all:
        if not args.month or not args.year:
            raise ValueError("--delete-month-all requires --month and --year.")
        if args.month < 1 or args.month > 12:
            raise ValueError("--month must be in range 1..12.")
        return MODE_DELETE_MONTH_ALL
    if args.add_only:
        return MODE_ADD_ONLY
    return MODE_FULL


def resolve_docx_file(args, mode: str) -> str | None:
    if mode == MODE_DELETE_MONTH_ALL:
        return None

    if args.docx_file:
        return validate_docx_file(args.docx_file)

    if args.select_file:
        selected = prompt_for_docx_file()
        if not selected:
            raise RuntimeError("No DOCX file selected.")
        return validate_docx_file(selected)

    if not args.no_file_picker and sys.stdin.isatty():
        selected = prompt_for_docx_file()
        if selected:
            return validate_docx_file(selected)
        raise RuntimeError("No DOCX file selected.")

    raise RuntimeError("No DOCX file provided. Use path, drag-and-drop, or --select-file.")


def resolve_calendar_id(args) -> str:
    if args.calendar_id:
        return args.calendar_id

    env_calendar = os.getenv("QAM_CALENDAR_ID")
    if env_calendar:
        return env_calendar

    if args.select_calendar or (not args.no_calendar_prompt and sys.stdin.isatty()):
        return prompt_for_calendar_selection()

    return "primary"


def validate_docx_file(docx_path: str) -> str:
    path = Path(docx_path).expanduser()
    if not path.exists():
        raise FileNotFoundError(f"DOCX file not found: {path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Expected a .docx file, got: {path}")
    return str(path.resolve())


def prompt_for_docx_file() -> str | None:
    try:
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            selected = filedialog.askopenfilename(
                title="Select roster DOCX file",
                filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            )
        finally:
            root.destroy()
        return selected or None
    except TclError as exc:
        raise RuntimeError("File-picker UI is unavailable in this environment.") from exc


def prompt_for_calendar_selection() -> str:
    calendars = calendar_api.list_calendars()
    if not calendars:
        raise RuntimeError("No calendars returned by Google Calendar API for this account.")

    print("")
    print("Select target Google Calendar:")
    for idx, cal in enumerate(calendars, start=1):
        summary = cal.get("summary", "<no title>")
        calendar_id = cal.get("id", "<no id>")
        primary = " (primary)" if cal.get("primary") else ""
        print(f"{idx}. {summary}{primary} [{calendar_id}]")

    while True:
        raw = input("Enter number: ").strip()
        if not raw.isdigit():
            print("Invalid input. Enter a number from the list.")
            continue
        selected_index = int(raw)
        if selected_index < 1 or selected_index > len(calendars):
            print("Selection out of range. Try again.")
            continue
        selected = calendars[selected_index - 1]
        selected_id = selected.get("id")
        if not selected_id:
            print("Selected calendar has no id. Choose another entry.")
            continue
        return selected_id


def configure_logging() -> None:
    project_root = Path(__file__).resolve().parent
    log_dir = project_root / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "creation.log"

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logging.basicConfig(level=logging.INFO, handlers=[file_handler, stream_handler], force=True)
    LOGGER.info("Logging configured at %s", log_file)


if __name__ == "__main__":
    raise SystemExit(main())
