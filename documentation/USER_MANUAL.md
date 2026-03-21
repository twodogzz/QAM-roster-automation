# QAM Roster Automation

Version: 0.1.2  
Audience: End users and developers  
Last updated: 2026-03-21

## 1. Purpose

QAM Roster Automation syncs volunteer shifts from a roster `.docx` file into Google Calendar.

It can:
- Create events for rostered days
- Delete old roster versions
- Delete current roster version
- Delete all roster versions for a selected month
- Preview all changes before writing

## 2. System Requirements

- Python 3.10+ recommended
- Internet access for Google OAuth and Google Calendar API
- Google account with calendar access
- Local OAuth client file (`modules/credentials.json` recommended)

Python dependencies:
- `google-api-python-client`
- `google-auth`
- `google-auth-oauthlib`
- `python-docx`

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

Notes:
- Keep `token.json` local and do not commit it.
- If token refresh fails due to invalid credentials, the app removes stale `token.json` and re-authenticates.

## 4. Running the Application

### 4.1 GUI Mode

```bash
python main.py --gui
```

On Windows, running `python main.py` with no workflow args opens GUI automatically.

### 4.2 CLI Mode

```bash
python main.py path\to\Roster_March_2026_v9.docx --dry-run
python main.py path\to\Roster_March_2026_v9.docx --replace-current
python main.py --cli path\to\Roster_March_2026_v9.docx --delete-only
python main.py --cli --delete-month-all --month 3 --year 2026 --calendar-id your_calendar_id
python main.py --cli path\to\Roster_March_2026_v9.docx --add-only
```

File picker options:

```bash
python main.py --select-file
python main.py --no-file-picker --calendar-id your_calendar_id path\to\Roster.docx
```

Calendar picker options:

```bash
python main.py --select-calendar path\to\Roster.docx
python main.py --no-calendar-prompt --calendar-id your_calendar_id path\to\Roster.docx
```

## 5. Multi-Month Roster Behavior

The parser now supports roster files that include:
- The roster month shown in the title
- Days from the previous month
- Days from the next month

Example:
- A DOCX titled `Queensland Air Museum – Front Counter Roster March 2026 v9` may contain March shifts and also entries such as `1. April ...`, `2. April ...`, and `3. April ...`
- Those entries are created as April calendar events, but their summary still stays tied to the roster label, for example `Wayne volunteer Queensland Air Museum - Mar v9`

This keeps every event identifiable by the roster version it came from, even when the actual event date is in another month.

## 6. How Date Detection Works

The parser determines dates using a simple set of rules:

1. The roster month and year are read from the document content, such as `March 2026`.
2. A normal cell like `31. Wayne Freestun ...` is treated as a day in the roster month.
3. A cell with an explicit month name, such as `1. April ...`, is treated as that named month.
4. Explicit month names are assumed to be only:
   - the previous month
   - the roster month
   - the next month
5. The year is adjusted automatically for January/December rollovers.

Examples:
- `January 2026` roster with `31. December ...` becomes `2025-12-31`
- `March 2026` roster with `1. April ...` becomes `2026-04-01`

## 7. Action Modes

- `full` (default): delete old versions in every month found in the DOCX, then create new events
- `delete_only`: delete prior versions in every month found in the DOCX
- `delete_current`: delete only events matching the current DOCX version across every month found in the DOCX
- `delete_month_all`: delete all matching events in the selected month/year, no DOCX required
- `add_only`: create events only, no deletes

`--replace-current` behavior:
- Only affects `full` mode
- Adds current-version events to the deletion list before recreating them

## 8. Matching and Event Labels

The parser extracts:
- The roster month and year from document text or table cells
- The roster version from the filename or document content
- Volunteer shifts by finding rows containing the configured volunteer name

Generated event summary format:
- `<Event Title> - <RosterMon> v<version>`
- Example: `Wayne volunteer Queensland Air Museum - Mar v9`

Important:
- Even if an event date is in April, an event from the March roster still uses `Mar v9`
- Delete logic for `full`, `delete_only`, and `delete_current` now covers every month found in that DOCX, not only the roster-title month

## 9. Logging and Feedback

The app now logs useful parser progress, including:
- The detected roster month and version
- How many day entries were found
- Which months were covered by the DOCX
- How many volunteer matches were found

Logs are written to `logs/creation.log`.

## 10. Troubleshooting

`Unable to detect roster month/year from roster document`
- Confirm the DOCX still includes text such as `March 2026` in a paragraph or table cell

`No roster day entries were found in the DOCX table`
- Confirm day cells still begin with a number and a period, such as `17.`

`Roster cell month must be the roster month, previous month, or next month`
- Confirm any explicit month names inside day cells only refer to the previous, current, or next month

`Invalid roster date`
- Check for malformed cells such as `31. April ...`

`Expected a .docx file`
- Ensure the input file extension is `.docx`

`No calendars returned by Google Calendar API for this account`
- Confirm the authenticated Google account has at least one accessible calendar

`--delete-month-all requires --month and --year`
- Supply both numeric values and keep month in `1..12`

## 11. Developer Notes

Key parser assumptions:
- Day rows begin with `^\s*\d{1,2}\.`
- Explicit month names in day cells are reliable when present
- Explicit month names only refer to the previous, current, or next month
- Event summaries always keep the roster month label for version tracking

Key parser outputs:
- `RosterParseResult.year` and `RosterParseResult.month` still represent the roster label month
- Each parsed event now carries its own real calendar date
- `covered_months` lists every `(year, month)` found in the DOCX

Compatibility script:
- `qam_docx_to_ics.py` now returns per-event dates and covered months so downstream tooling can handle multi-month rosters cleanly
