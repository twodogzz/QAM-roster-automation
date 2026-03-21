# QAM Roster Automation

Version: 0.1.3  
Audience: End users and developers  
Last updated: 2026-03-21

## 1. Purpose

QAM Roster Automation syncs volunteer shifts from a roster `.docx` file into Google Calendar.

It supports two flows:
- Wayne-only sync to the main calendar
- Optional all-workers sync to a separate calendar

The all-workers flow is off by default and only runs when you explicitly enable it.

## 2. System Requirements

- Python 3.10+ recommended
- Internet access for Google OAuth and Google Calendar API
- Google account with calendar access
- Local OAuth client file (`modules/credentials.json` recommended)

Install:

```bash
python -m pip install -r requirements.txt
```

## 3. First-Time Setup

1. Copy OAuth template:

```bash
copy modules\credentials.example.json modules\credentials.json
```

2. Replace placeholder values in `modules/credentials.json` with your real Google OAuth desktop client credentials.
3. Run the app once. A browser sign-in flow will open.
4. After successful sign-in, `token.json` is created in the project root.

## 4. Multi-Month Roster Support

The parser supports roster files that include:
- the roster month shown in the title
- the previous month
- the next month

Examples:
- `March 2026` roster with `1. April ...` creates an event on `2026-04-01`
- `January 2026` roster with `31. December ...` creates an event on `2025-12-31`

Important:
- Event summaries still use the roster month/version label
- A March roster event in April still uses `... - Mar v<version>`

## 5. Date Detection Rules

The parser keeps the logic simple:

1. It reads the roster month and year from document text such as `March 2026`.
2. A normal day cell like `17. Wayne Freestun ...` stays in the roster month.
3. A day cell with an explicit month name like `1. April ...` uses that month instead.
4. Explicit month names are assumed to be only the previous month, current month, or next month.
5. The year is adjusted automatically for December/January rollovers.

## 6. Wayne-Only Sync

Default behavior:
- Only Wayne-matching roster rows are turned into calendar events
- The target calendar is the normal selected calendar

CLI example:

```bash
python main.py path\to\Roster_March_2026_v9.docx --dry-run
```

GUI example:
- Leave `Also sync All Workers calendar` unchecked
- Preview and then proceed as normal

## 7. Optional All-Workers Sync

The new all-workers flow creates:
- one event per roster day
- with the full worker list in the event description
- on a separate Google Calendar

It does not run automatically.

### 7.1 CLI Trigger

Enable it explicitly with:

```bash
python main.py path\to\Roster_March_2026_v9.docx --sync-all-workers --all-workers-calendar-id your_other_calendar_id
```

You can also let the app prompt for the separate calendar:

```bash
python main.py path\to\Roster_March_2026_v9.docx --sync-all-workers --select-all-workers-calendar
```

Environment fallback:
- `QAM_ALL_WORKERS_CALENDAR_ID`

### 7.2 GUI Trigger

In the GUI:
1. Select the roster DOCX
2. Choose the normal Wayne calendar
3. Tick `Also sync All Workers calendar`
4. Choose a different calendar in `All Workers Calendar`
5. Preview changes
6. Click `Proceed`

### 7.3 Important Rules

- The all-workers calendar must be different from the Wayne calendar
- The all-workers flow uses the same roster parsing and multi-month logic as the Wayne flow
- The same action mode applies to both flows

## 8. Action Modes

- `full`: delete old versions, then create new events
- `delete_only`: delete prior versions only
- `delete_current`: delete only the current DOCX version
- `delete_month_all`: delete all matching events in a selected month/year
- `add_only`: create events only

Notes:
- `full`, `delete_only`, and `delete_current` now operate across every month found in the DOCX
- The optional all-workers sync is not available for `delete_month_all` because that mode has no DOCX context

## 9. Calendar Selection

Wayne calendar priority:
1. `--calendar-id`
2. `QAM_CALENDAR_ID`
3. Interactive picker
4. `primary`

All-workers calendar priority:
1. `--all-workers-calendar-id`
2. `QAM_ALL_WORKERS_CALENDAR_ID`
3. `--select-all-workers-calendar` or interactive picker when enabled

## 10. Logging and Preview

Preview output now shows:
- Wayne calendar plan
- All-workers calendar plan, when enabled
- covered months
- delete/create counts

Nothing is written until you proceed.

If `Dry run` is enabled, no calendar changes are made even after proceeding.

## 11. Troubleshooting

`Unable to detect roster month/year from roster document`
- Confirm the DOCX still includes text such as `March 2026`

`No roster day entries were found in the DOCX table`
- Confirm day cells still begin with a number and a period, such as `17.`

`Roster cell month must be the roster month, previous month, or next month`
- Confirm explicit month labels in day cells only refer to the adjacent months

`All workers sync requires a separate calendar`
- Provide `--all-workers-calendar-id`
- or set `QAM_ALL_WORKERS_CALENDAR_ID`
- or use `--select-all-workers-calendar`

`The all workers calendar must be different from the Wayne calendar`
- Choose a different calendar for the optional all-workers flow

## 12. Developer Notes

Implementation shape:
- Wayne sync remains the default path
- All-workers sync is a second optional plan layered on top of the same DOCX parse result
- One all-workers event is created per roster day using the full worker list

Key assumptions:
- Worker identification for Wayne remains a case-insensitive name match
- All-workers events are day-based, not person-based
- Explicit month names inside day cells are reliable when present
