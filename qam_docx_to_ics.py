"""Compatibility wrapper around the DOCX roster parser."""

from __future__ import annotations

from argparse import ArgumentParser
import json

from modules.roster_parser import parse_roster_docx


def extract_events_from_docx(docx_path: str):
    """Return structured roster data for compatibility with legacy calls."""
    parsed = parse_roster_docx(docx_path)
    return {
        "source_file": str(parsed.source_file),
        "roster_year": parsed.year,
        "roster_month": parsed.month,
        "version": parsed.version,
        "summary_prefix": parsed.summary_prefix,
        "covered_months": [
            {
                "year": year,
                "month": month,
                "label": f"{year:04d}-{month:02d}",
            }
            for year, month in parsed.covered_months
        ],
        "events": [
            {
                "date": event.event_date.isoformat(),
                "year": event.event_date.year,
                "month": event.event_date.month,
                "day": event.day,
                "workers_raw": event.workers_raw,
                "summary": event.summary,
                "description": event.description,
                "location": event.location,
                "start_time": event.start_time.isoformat(),
                "end_time": event.end_time.isoformat(),
            }
            for event in parsed.roster_events
        ],
        "all_day_workers": [
            {
                "date": assignment.event_date.isoformat(),
                "year": assignment.event_date.year,
                "month": assignment.event_date.month,
                "day": assignment.event_date.day,
                "workers_raw": assignment.workers_raw,
            }
            for assignment in parsed.all_day_workers
        ],
    }


def main() -> int:
    parser = ArgumentParser(description="Parse QAM DOCX roster into structured events")
    parser.add_argument("docx_file", help="Path to roster DOCX")
    args = parser.parse_args()
    print(json.dumps(extract_events_from_docx(args.docx_file), indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
