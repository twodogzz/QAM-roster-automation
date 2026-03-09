"""Parse QAM roster DOCX files into structured event objects."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time
from pathlib import Path
import re
from typing import Iterable

from docx import Document


MONTH_MAP = {
    "JANUARY": 1,
    "JAN": 1,
    "FEBRUARY": 2,
    "FEB": 2,
    "MARCH": 3,
    "MAR": 3,
    "APRIL": 4,
    "APR": 4,
    "MAY": 5,
    "JUNE": 6,
    "JUN": 6,
    "JULY": 7,
    "JUL": 7,
    "AUGUST": 8,
    "AUG": 8,
    "SEPTEMBER": 9,
    "SEP": 9,
    "SEPT": 9,
    "OCTOBER": 10,
    "OCT": 10,
    "NOVEMBER": 11,
    "NOV": 11,
    "DECEMBER": 12,
    "DEC": 12,
}
MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


@dataclass(frozen=True)
class RosterEvent:
    day: int
    workers_raw: str
    summary: str
    description: str
    location: str
    start_time: time
    end_time: time


@dataclass(frozen=True)
class RosterParseResult:
    source_file: Path
    year: int
    month: int
    version: str
    summary_prefix: str
    roster_events: list[RosterEvent]
    all_day_workers: list[tuple[int, str]]


def extract_version_from_filename(path: Path) -> str | None:
    match = re.search(r"\bv(\d+)\b", path.name, re.IGNORECASE)
    return match.group(1) if match else None


def extract_version_from_doc(doc: Document) -> str | None:  # pyright: ignore[reportGeneralTypeIssues]
    pattern = re.compile(r"\bv(\d+)\b", re.IGNORECASE)
    for para in doc.paragraphs:
        match = pattern.search(para.text)
        if match:
            return match.group(1)
    for text in _iter_table_cell_text(doc):
        match = pattern.search(text)
        if match:
            return match.group(1)
    return None


def extract_month_year(doc: Document) -> tuple[int, int]:  # pyright: ignore[reportGeneralTypeIssues]
    pattern = re.compile(r"\b([A-Za-z]+)\s+(\d{4})\b")
    for para in doc.paragraphs:
        match = pattern.search(para.text.strip())
        if match:
            month = MONTH_MAP.get(match.group(1).upper())
            if month:
                return month, int(match.group(2))
    for text in _iter_table_cell_text(doc):
        match = pattern.search(text.strip())
        if match:
            month = MONTH_MAP.get(match.group(1).upper())
            if month:
                return month, int(match.group(2))
    raise ValueError("Unable to detect month/year from roster document.")


def parse_roster_docx(
    docx_path: str | Path,
    *,
    volunteer_name: str = "Wayne Freestun",
    event_title: str = "Wayne volunteer Queensland Air Museum",
    location: str = "Queensland Air Museum",
    shift_start: time = time(9, 0),
    shift_end: time = time(16, 30),
) -> RosterParseResult:
    path = Path(docx_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Roster DOCX not found: {path}")

    doc = Document(str(path))
    month, year = extract_month_year(doc)
    version = extract_version_from_filename(path) or extract_version_from_doc(doc) or "0"
    month_label = MONTH_SHORT[month - 1]
    summary_with_version = f"{event_title} - {month_label} v{version}"
    summary_prefix = f"{event_title} - {month_label}"

    all_workers_map: dict[int, str] = {}
    volunteer_events: dict[int, str] = {}
    for text in _iter_table_cell_text(doc):
        normalized = normalize_whitespace(text)
        if not normalized:
            continue
        day = parse_day_number(normalized)
        if day is None:
            continue
        workers_raw = re.sub(r"^\s*\d{1,2}\.\s*", "", normalized).strip()
        all_workers_map[day] = workers_raw
        if volunteer_name.lower() in normalized.lower():
            volunteer_events[day] = workers_raw

    roster_events = [
        RosterEvent(
            day=day,
            workers_raw=workers,
            summary=summary_with_version,
            description=f"Workers: {workers} (Roster v{version})",
            location=location,
            start_time=shift_start,
            end_time=shift_end,
        )
        for day, workers in sorted(volunteer_events.items())
    ]

    return RosterParseResult(
        source_file=path,
        year=year,
        month=month,
        version=version,
        summary_prefix=summary_prefix,
        roster_events=roster_events,
        all_day_workers=sorted(all_workers_map.items()),
    )


def to_google_event_payload(event: RosterEvent, *, year: int, month: int, timezone: str) -> dict:
    start_dt = datetime(year, month, event.day, event.start_time.hour, event.start_time.minute)
    end_dt = datetime(year, month, event.day, event.end_time.hour, event.end_time.minute)
    return {
        "summary": event.summary,
        "location": event.location,
        "description": event.description,
        "start": {"dateTime": start_dt.isoformat(), "timeZone": timezone},
        "end": {"dateTime": end_dt.isoformat(), "timeZone": timezone},
    }


def parse_day_number(text: str) -> int | None:
    match = re.match(r"\s*(\d{1,2})\.", text)
    return int(match.group(1)) if match else None


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _iter_table_cell_text(doc: Document) -> Iterable[str]:  # pyright: ignore[reportGeneralTypeIssues]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield cell.text

