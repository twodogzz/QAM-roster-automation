"""Shared sync planning and execution logic for CLI and GUI flows."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import re
from typing import Any

from modules import calendar_api
from modules.roster_parser import MONTH_SHORT, parse_roster_docx, to_google_event_payload


VERSION_PATTERN = re.compile(r"\bv(\d+)\b", re.IGNORECASE)

MODE_FULL = "full"
MODE_DELETE_ONLY = "delete_only"
MODE_ADD_ONLY = "add_only"
MODE_DELETE_CURRENT = "delete_current"
MODE_DELETE_MONTH_ALL = "delete_month_all"
VALID_MODES = {MODE_FULL, MODE_DELETE_ONLY, MODE_ADD_ONLY, MODE_DELETE_CURRENT, MODE_DELETE_MONTH_ALL}


@dataclass(frozen=True)
class SyncOptions:
    docx_file: str | None = None
    calendar_id: str = "primary"
    timezone: str = "Australia/Brisbane"
    dry_run: bool = False
    replace_current: bool = False
    mode: str = MODE_FULL
    month: int | None = None
    year: int | None = None
    volunteer_name: str = "Wayne Freestun"
    event_title: str = "Wayne volunteer Queensland Air Museum"
    location: str = "Queensland Air Museum"


@dataclass(frozen=True)
class SyncPlan:
    options: SyncOptions
    summary_prefix: str
    month: int
    year: int
    version: str | None
    existing_count: int
    delete_events: list[dict[str, str]]
    create_payloads: list[dict[str, Any]]


@dataclass(frozen=True)
class SyncResult:
    deleted_count: int
    created_count: int


def prepare_sync_plan(options: SyncOptions) -> SyncPlan:
    if options.mode not in VALID_MODES:
        raise ValueError(f"Invalid mode '{options.mode}'. Expected one of: {sorted(VALID_MODES)}")

    calendar_api.set_calendar_id(options.calendar_id)

    parsed = None
    if options.mode != MODE_DELETE_MONTH_ALL:
        if not options.docx_file:
            raise ValueError("DOCX file is required for the selected action.")
        parsed = parse_roster_docx(
            options.docx_file,
            volunteer_name=options.volunteer_name,
            event_title=options.event_title,
            location=options.location,
        )
        month = parsed.month
        year = parsed.year
        version = parsed.version
        summary_prefix = parsed.summary_prefix
    else:
        if not options.month or not options.year:
            raise ValueError("Month and year are required for 'delete all versions in chosen month'.")
        month = options.month
        year = options.year
        version = None
        summary_prefix = f"{options.event_title} - {MONTH_SHORT[month - 1]}"

    existing = calendar_api.list_events(summary_prefix)
    month_events = [event for event in existing if _matches_month_year(event, year, month)]

    delete_events: list[dict[str, str]] = []
    if options.mode == MODE_DELETE_MONTH_ALL:
        delete_events = _event_delete_rows(month_events)
    elif options.mode == MODE_DELETE_CURRENT:
        if not version:
            raise ValueError("Current version is unavailable; DOCX parse failed.")
        delete_events = _event_delete_rows(
            [event for event in month_events if _extract_version(event.get("summary", "")) == version]
        )
    elif options.mode in {MODE_FULL, MODE_DELETE_ONLY}:
        if not version:
            raise ValueError("Current version is unavailable; DOCX parse failed.")
        prior_events = [
            event for event in month_events if _extract_version(event.get("summary", "")) not in {None, version}
        ]
        delete_events = _event_delete_rows(prior_events)
        if options.replace_current:
            current_events = [
                event for event in month_events if _extract_version(event.get("summary", "")) == version
            ]
            delete_events.extend(_event_delete_rows(current_events))

    create_payloads: list[dict[str, Any]] = []
    if options.mode in {MODE_FULL, MODE_ADD_ONLY} and parsed:
        create_payloads = [
            to_google_event_payload(
                event,
                year=parsed.year,
                month=parsed.month,
                timezone=options.timezone,
            )
            for event in parsed.roster_events
        ]

    return SyncPlan(
        options=options,
        summary_prefix=summary_prefix,
        month=month,
        year=year,
        version=version,
        existing_count=len(existing),
        delete_events=delete_events,
        create_payloads=create_payloads,
    )


def execute_sync_plan(plan: SyncPlan) -> SyncResult:
    deleted = 0
    created = 0

    if not plan.options.dry_run:
        for row in plan.delete_events:
            event_id = row.get("id")
            if not event_id:
                continue
            calendar_api.delete_event_by_id(event_id)
            deleted += 1
        for payload in plan.create_payloads:
            calendar_api.create_event(payload)
            created += 1

    return SyncResult(deleted_count=deleted, created_count=created)


def summarize_plan(plan: SyncPlan) -> str:
    roster_label = f"{plan.month:02d}/{plan.year}"
    if plan.version:
        roster_label = f"{roster_label} v{plan.version}"
    lines = [
        f"Mode: {plan.options.mode}",
        f"Dry run: {plan.options.dry_run}",
        f"Calendar ID: {plan.options.calendar_id}",
        f"Roster: {roster_label}",
        f"Existing matching events: {plan.existing_count}",
        f"Events to delete: {len(plan.delete_events)}",
        f"Events to create: {len(plan.create_payloads)}",
    ]

    if plan.delete_events:
        lines.append("")
        lines.append("Events to delete:")
        for row in plan.delete_events:
            lines.append(f"- {row.get('start', '?')} | {row.get('summary', '<no summary>')}")

    if plan.create_payloads:
        lines.append("")
        lines.append("Events to create:")
        for payload in plan.create_payloads:
            start = payload.get("start", {}).get("dateTime", "?")
            summary = payload.get("summary", "<no summary>")
            lines.append(f"- {start} | {summary}")

    return "\n".join(lines)


def _event_delete_rows(events: list[dict[str, Any]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for event in events:
        event_id = event.get("id")
        if not event_id:
            continue
        start = event.get("start", {})
        start_text = start.get("dateTime") or start.get("date") or "?"
        rows.append(
            {
                "id": str(event_id),
                "summary": str(event.get("summary", "")),
                "start": str(start_text),
            }
        )
    return rows


def _extract_version(summary: str) -> str | None:
    match = VERSION_PATTERN.search(summary or "")
    return match.group(1) if match else None


def _matches_month_year(event: dict[str, Any], year: int, month: int) -> bool:
    start = event.get("start", {})
    raw_start = start.get("dateTime") or start.get("date")
    if not raw_start:
        return False
    try:
        event_dt = datetime.fromisoformat(str(raw_start).replace("Z", "+00:00"))
    except ValueError:
        return False
    return event_dt.year == year and event_dt.month == month

