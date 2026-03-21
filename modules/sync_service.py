"""Shared sync planning and execution logic for CLI and GUI flows."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import logging
import os
import re
from typing import Any

from modules import calendar_api
from modules.roster_parser import (
    MONTH_SHORT,
    RosterEvent,
    RosterParseResult,
    build_all_workers_events,
    parse_roster_docx,
    to_google_event_payload,
)


VERSION_PATTERN = re.compile(r"\bv(\d+)\b", re.IGNORECASE)

MODE_FULL = "full"
MODE_DELETE_ONLY = "delete_only"
MODE_ADD_ONLY = "add_only"
MODE_DELETE_CURRENT = "delete_current"
MODE_DELETE_MONTH_ALL = "delete_month_all"
VALID_MODES = {MODE_FULL, MODE_DELETE_ONLY, MODE_ADD_ONLY, MODE_DELETE_CURRENT, MODE_DELETE_MONTH_ALL}

DEFAULT_ALL_WORKERS_TITLE = "QAM Front Counter Roster"
ALL_WORKERS_CALENDAR_ENV = "QAM_ALL_WORKERS_CALENDAR_ID"

LOGGER = logging.getLogger(__name__)


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
    sync_all_workers: bool = False
    all_workers_calendar_id: str | None = None
    all_workers_event_title: str = DEFAULT_ALL_WORKERS_TITLE


@dataclass(frozen=True)
class SyncPlan:
    plan_name: str
    options: SyncOptions
    calendar_id: str
    summary_prefix: str
    roster_month: int
    roster_year: int
    covered_months: list[tuple[int, int]]
    version: str | None
    existing_count: int
    delete_events: list[dict[str, str]]
    create_payloads: list[dict[str, Any]]


@dataclass(frozen=True)
class AutomationPlan:
    options: SyncOptions
    primary_plan: SyncPlan
    all_workers_plan: SyncPlan | None = None


@dataclass(frozen=True)
class SyncResult:
    deleted_count: int
    created_count: int


@dataclass(frozen=True)
class AutomationResult:
    primary_result: SyncResult
    all_workers_result: SyncResult | None = None

    @property
    def deleted_count(self) -> int:
        total = self.primary_result.deleted_count
        if self.all_workers_result:
            total += self.all_workers_result.deleted_count
        return total

    @property
    def created_count(self) -> int:
        total = self.primary_result.created_count
        if self.all_workers_result:
            total += self.all_workers_result.created_count
        return total


def prepare_sync_plan(options: SyncOptions) -> AutomationPlan:
    if options.mode not in VALID_MODES:
        raise ValueError(f"Invalid mode '{options.mode}'. Expected one of: {sorted(VALID_MODES)}")

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

    primary_plan = _prepare_single_sync_plan(
        options,
        plan_name="Wayne Calendar",
        calendar_id=options.calendar_id,
        parsed=parsed,
        events=parsed.roster_events if parsed else None,
        summary_prefix=parsed.summary_prefix if parsed else None,
        roster_month=parsed.month if parsed else options.month,
        roster_year=parsed.year if parsed else options.year,
        covered_months=parsed.covered_months if parsed else None,
        version=parsed.version if parsed else None,
    )

    all_workers_plan = None
    if options.sync_all_workers:
        if options.mode == MODE_DELETE_MONTH_ALL:
            raise ValueError("All workers sync requires a DOCX-backed mode; it is not available with delete_month_all.")
        if parsed is None:
            raise ValueError("All workers sync requires a parsed DOCX roster.")

        all_workers_calendar_id = resolve_all_workers_calendar_id(options)
        if all_workers_calendar_id == options.calendar_id:
            raise ValueError("The all workers calendar must be different from the Wayne calendar.")

        # This optional flow writes one event per roster day to a separate
        # calendar, while the default Wayne flow only writes Wayne's shifts.
        all_workers_events = build_all_workers_events(
            parsed,
            event_title=options.all_workers_event_title,
            location=options.location,
        )
        all_workers_plan = _prepare_single_sync_plan(
            options,
            plan_name="All Workers Calendar",
            calendar_id=all_workers_calendar_id,
            parsed=parsed,
            events=all_workers_events,
            summary_prefix=f"{options.all_workers_event_title} - {MONTH_SHORT[parsed.month - 1]}",
            roster_month=parsed.month,
            roster_year=parsed.year,
            covered_months=parsed.covered_months,
            version=parsed.version,
        )

    return AutomationPlan(options=options, primary_plan=primary_plan, all_workers_plan=all_workers_plan)


def execute_sync_plan(plan: AutomationPlan) -> AutomationResult:
    primary_result = _execute_single_sync_plan(plan.primary_plan)
    all_workers_result = None
    if plan.all_workers_plan is not None:
        all_workers_result = _execute_single_sync_plan(plan.all_workers_plan)
    return AutomationResult(primary_result=primary_result, all_workers_result=all_workers_result)


def summarize_plan(plan: AutomationPlan) -> str:
    sections = [_summarize_single_plan(plan.primary_plan)]
    if plan.all_workers_plan is not None:
        sections.append("")
        sections.append(_summarize_single_plan(plan.all_workers_plan))
    return "\n".join(sections)


def resolve_all_workers_calendar_id(options: SyncOptions) -> str:
    if options.all_workers_calendar_id:
        return options.all_workers_calendar_id

    env_calendar = os.getenv(ALL_WORKERS_CALENDAR_ENV, "").strip()
    if env_calendar:
        return env_calendar

    raise ValueError(
        "All workers sync was requested but no separate calendar was configured. "
        f"Provide all_workers_calendar_id or set {ALL_WORKERS_CALENDAR_ENV}."
    )


def _prepare_single_sync_plan(
    options: SyncOptions,
    *,
    plan_name: str,
    calendar_id: str,
    parsed: RosterParseResult | None,
    events: list[RosterEvent] | None,
    summary_prefix: str | None,
    roster_month: int | None,
    roster_year: int | None,
    covered_months: list[tuple[int, int]] | None,
    version: str | None,
) -> SyncPlan:
    if summary_prefix is None or roster_month is None or roster_year is None:
        raise ValueError("Sync plan is missing roster metadata.")

    if covered_months is None:
        covered_months = [(roster_year, roster_month)]

    calendar_api.set_calendar_id(calendar_id)
    existing = calendar_api.list_events(summary_prefix)
    scoped_events = [event for event in existing if _matches_any_month_year(event, covered_months)]

    LOGGER.info(
        "Prepared %s scope for roster %02d/%04d covering %s on calendar %s",
        plan_name,
        roster_month,
        roster_year,
        ", ".join(f"{month:02d}/{year}" for year, month in covered_months),
        calendar_id,
    )
    LOGGER.info("Found %d matching existing events, %d inside scope", len(existing), len(scoped_events))

    delete_events: list[dict[str, str]] = []
    if options.mode == MODE_DELETE_MONTH_ALL:
        delete_events = _event_delete_rows(scoped_events)
    elif options.mode == MODE_DELETE_CURRENT:
        if not version:
            raise ValueError("Current version is unavailable; DOCX parse failed.")
        delete_events = _event_delete_rows(
            [event for event in scoped_events if _extract_version(event.get("summary", "")) == version]
        )
    elif options.mode in {MODE_FULL, MODE_DELETE_ONLY}:
        if not version:
            raise ValueError("Current version is unavailable; DOCX parse failed.")
        prior_events = [
            event for event in scoped_events if _extract_version(event.get("summary", "")) not in {None, version}
        ]
        delete_events = _event_delete_rows(prior_events)
        if options.replace_current:
            current_events = [
                event for event in scoped_events if _extract_version(event.get("summary", "")) == version
            ]
            delete_events.extend(_event_delete_rows(current_events))

    create_payloads: list[dict[str, Any]] = []
    if options.mode in {MODE_FULL, MODE_ADD_ONLY} and events:
        create_payloads = [to_google_event_payload(event, timezone=options.timezone) for event in events]

    return SyncPlan(
        plan_name=plan_name,
        options=options,
        calendar_id=calendar_id,
        summary_prefix=summary_prefix,
        roster_month=roster_month,
        roster_year=roster_year,
        covered_months=covered_months,
        version=version,
        existing_count=len(existing),
        delete_events=delete_events,
        create_payloads=create_payloads,
    )


def _execute_single_sync_plan(plan: SyncPlan) -> SyncResult:
    deleted = 0
    created = 0

    if not plan.options.dry_run:
        calendar_api.set_calendar_id(plan.calendar_id)
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


def _summarize_single_plan(plan: SyncPlan) -> str:
    roster_label = f"{plan.roster_month:02d}/{plan.roster_year}"
    if plan.version:
        roster_label = f"{roster_label} v{plan.version}"
    covered_label = ", ".join(f"{month:02d}/{year}" for year, month in plan.covered_months)
    lines = [
        f"{plan.plan_name}:",
        f"Mode: {plan.options.mode}",
        f"Dry run: {plan.options.dry_run}",
        f"Calendar ID: {plan.calendar_id}",
        f"Roster label: {roster_label}",
        f"Covered months: {covered_label}",
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


def _matches_any_month_year(event: dict[str, Any], covered_months: list[tuple[int, int]]) -> bool:
    event_period = _event_month_year(event)
    if event_period is None:
        return False
    return event_period in set(covered_months)


def _event_month_year(event: dict[str, Any]) -> tuple[int, int] | None:
    start = event.get("start", {})
    raw_start = start.get("dateTime") or start.get("date")
    if not raw_start:
        return None
    try:
        event_dt = datetime.fromisoformat(str(raw_start).replace("Z", "+00:00"))
    except ValueError:
        return None
    return event_dt.year, event_dt.month
