"""Compatibility wrapper around the new DOCX roster parser."""

from __future__ import annotations

from argparse import ArgumentParser
import json

from modules.roster_parser import parse_roster_docx


def extract_events_from_docx(docx_path: str):
    """Return structured roster data for compatibility with legacy calls."""
    parsed = parse_roster_docx(docx_path)
    return {
        "source_file": str(parsed.source_file),
        "year": parsed.year,
        "month": parsed.month,
        "version": parsed.version,
        "summary_prefix": parsed.summary_prefix,
        "events": [
            {
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
        "all_day_workers": parsed.all_day_workers,
    }


def main() -> int:
    parser = ArgumentParser(description="Parse QAM DOCX roster into structured events")
    parser.add_argument("docx_file", help="Path to roster DOCX")
    args = parser.parse_args()
    print(json.dumps(extract_events_from_docx(args.docx_file), indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

