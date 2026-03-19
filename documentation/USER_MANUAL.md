# QAM Roster Automation

Version: 0.1.1  
Audience: End users and developers  
Last updated: 2026-03-09

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
- If token refresh fails due invalid credentials, the app can remove stale `token.json` and re-authenticate.

## 4. Running the Application

### 4.1 GUI Mode (recommended on Windows)

```bash
python main.py --gui
```

On Windows, running `python main.py` with no workflow args opens GUI automatically.

Main GUI workflow:
1. Select a roster DOCX (`Browse...`)
2. Choose target Google Calendar
3. Choose action mode
4. Configure Dry Run and other options
5. Click `Preview Changes`
6. Review "Identified Changes"
7. Click `Proceed` and confirm

GUI fields:
- `Roster DOCX`: source file for roster parsing
- `Google Calendar`: destination calendar
- `Action`: sync mode
- `Delete Month/Year`: used for "Delete all versions in chosen month"
- `Dry run`: preview-only execution (no writes)
- `Include current version in delete`: enabled only for Full sync
- `Timezone`: event timezone for new events
- `Volunteer`, `Event Title`, `Location`: parser and event metadata defaults

### 4.2 CLI Mode

Basic examples:

```bash
python main.py path\to\Roster_March_2026_v7.docx --dry-run
python main.py path\to\Roster_March_2026_v7.docx --replace-current
python main.py --cli path\to\Roster_March_2026_v7.docx --delete-only
python main.py --cli --delete-month-all --month 3 --year 2026 --calendar-id your_calendar_id
python main.py --cli path\to\Roster_March_2026_v7.docx --add-only
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

## 5. Action Modes

- `full` (default): delete old versions, then create events from current DOCX
- `delete_only`: delete prior versions only, no creates
- `delete_current`: delete only events matching current DOCX version
- `delete_month_all`: delete all matching events in selected month/year
- `add_only`: create events only, no deletes

`--replace-current` behavior:
- Only affects `full` mode
- Adds current version events to deletion list before creates

## 6. How Matching Works

The parser extracts:
- Month and year from document text/table cells (e.g., "March 2026")
- Version from filename (`v7`) or document content (`v7`)
- Volunteer shifts by finding lines containing configured volunteer name

Generated event summary format:
- `<Event Title> - <Mon> v<version>`
- Example: `Wayne volunteer Queensland Air Museum - Mar v7`

Delete scope is controlled by:
- Matching summary prefix
- Filtering to selected/parsed month and year
- Version logic based on selected mode

## 7. Configuration Reference

Command-line flags:
- `--gui`, `--cli`
- `--select-file`, `--no-file-picker`
- `--calendar-id`, `--select-calendar`, `--no-calendar-prompt`
- `--timezone`
- `--dry-run`
- `--replace-current`
- `--delete-only`, `--delete-current`, `--delete-month-all`, `--add-only`
- `--month`, `--year` (required with `--delete-month-all`)
- `--volunteer-name`, `--event-title`, `--location`

Environment variables:
- `QAM_CALENDAR_ID`: default calendar id when `--calendar-id` not provided
- `QAM_GOOGLE_CLIENT_SECRET_PATH`: preferred explicit OAuth client secret path
- `GOOGLE_OAUTH_CLIENT_SECRETS`: alternate OAuth client secret path

OAuth client secret discovery order:
1. `QAM_GOOGLE_CLIENT_SECRET_PATH`
2. `GOOGLE_OAUTH_CLIENT_SECRETS`
3. `modules/credentials.json`
4. `credentials.json`
5. One legacy file matching `modules/client_secret_*.apps.googleusercontent.com.json`

## 8. Output, Logging, and Safety

- Plan preview includes counts and detailed event lists.
- In dry run mode, no calendar writes are made.
- Runtime logs are written to `logs/creation.log`.
- Console logs include created/deleted counts and selected options.

## 9. Troubleshooting

`Missing Google API dependencies`
- Run `python -m pip install -r requirements.txt`

`Google OAuth client JSON not found`
- Ensure `modules/credentials.json` exists or set `QAM_GOOGLE_CLIENT_SECRET_PATH`

`No DOCX file selected` or `No DOCX file provided`
- Provide positional DOCX path or use `--select-file`

`Expected a .docx file`
- Ensure input file extension is `.docx`

`Unable to detect month/year from roster document`
- Confirm month/year text is present in paragraph or table cells (e.g., "March 2026")

`No calendars returned by Google Calendar API for this account`
- Confirm authenticated Google account has at least one accessible calendar

`--delete-month-all requires --month and --year`
- Supply both numeric values and keep month in `1..12`

## 10. Developer Technical Reference

### 10.1 Architecture Overview

Entry points:
- `main.py`: CLI parsing, mode selection, app startup routing
- `gui_app.py`: Tkinter desktop GUI wrapper around shared sync service
- `qam_docx_to_ics.py`: compatibility parser utility outputting JSON

Core modules:
- `modules/roster_parser.py`: DOCX parsing and event payload generation
- `modules/sync_service.py`: sync planning, delete/create decision logic, execution
- `modules/calendar_api.py`: Google Calendar API integration and OAuth credential handling

High-level flow:
1. Build `SyncOptions` from CLI/GUI input
2. `prepare_sync_plan(options)` creates deterministic delete/create plan
3. `summarize_plan(plan)` renders human-readable preview
4. `execute_sync_plan(plan)` applies writes unless `dry_run=True`

### 10.2 Data Contracts

`SyncOptions`:
- Inputs from GUI/CLI. Contains mode, calendar id, timezone, behavior flags, and metadata defaults.

`SyncPlan`:
- Immutable plan object containing:
  - parsed month/year/version
  - existing event count
  - `delete_events` rows (`id`, `summary`, `start`)
  - `create_payloads` for Google Calendar API

`SyncResult`:
- Operation result counters (`deleted_count`, `created_count`)

`RosterParseResult`:
- Structured parser output including source file, period, version, summary prefix, parsed events

### 10.3 Mode Logic (sync_service)

- `MODE_FULL`: delete prior versions; optionally current version; then create parsed events
- `MODE_DELETE_ONLY`: delete prior versions only
- `MODE_ADD_ONLY`: create parsed events only
- `MODE_DELETE_CURRENT`: delete events matching parsed version only
- `MODE_DELETE_MONTH_ALL`: delete all matching month events, no DOCX required

Version extraction for delete matching uses regex `\bv(\d+)\b` against event summaries.

### 10.4 Parser Assumptions (roster_parser)

- Day rows are recognized by leading pattern `^\s*\d{1,2}\.`
- Volunteer match is case-insensitive substring on normalized row text
- Month/year are discovered from free text in paragraphs and table cells
- Shift times default to `09:00` to `16:30`
- Event description embeds raw worker text and roster version

Potential parser fragility points:
- Unexpected document formatting or changed day row notation
- Month/year absent from parseable text
- Volunteer name spelling mismatch

### 10.5 Google API Layer (calendar_api)

- Singleton-like cached service (`_SERVICE`) and mutable calendar id (`_CALENDAR_ID`)
- Full calendar scope: `https://www.googleapis.com/auth/calendar`
- `list_events(q=...)` uses free-text search and client-side month/year filtering later
- Token refresh failures for invalid auth trigger stale token removal and fresh login

### 10.6 GUI Notes (gui_app)

- GUI is a thin layer over `sync_service`
- Uses shared planning and execution functions to keep CLI and GUI behavior aligned
- Calendar list is loaded on startup and can be refreshed
- Action-specific controls are enabled/disabled dynamically
- Window icon loads from `QAM-Logo-1-2048x1310whiteBGRND.png` at startup

### 10.7 Extension Guidelines

Adding a new sync mode:
1. Add mode constant in `modules/sync_service.py`
2. Add mode selection argument in `main.py`
3. Add GUI label mapping in `gui_app.py`
4. Implement delete/create behavior in `prepare_sync_plan`
5. Ensure `summarize_plan` still communicates behavior clearly

Changing parser behavior:
1. Update parsing helpers in `modules/roster_parser.py`
2. Keep `RosterParseResult` contract stable or update all consumers
3. Validate against representative DOCX samples

Changing event schema:
1. Modify `to_google_event_payload`
2. Ensure downstream Google API fields remain valid
3. Re-test all modes, especially create paths

### 10.8 Operational Recommendations

- Add automated tests for:
  - parser edge cases
  - mode decision matrix in `prepare_sync_plan`
  - summary rendering for change review
- Add sample roster fixtures in `tests/fixtures/`
- Keep credentials and tokens out of version control
- Consider structured logging fields for easier production diagnostics

### 10.9 Quick Maintenance Checklist

Before release:
1. Verify `project.json` version
2. Run dry-run checks on a known roster DOCX
3. Run full sync against a test calendar
4. Validate delete-month-all safety on non-production data
5. Review `logs/creation.log` for unexpected warnings

## 11. Building Windows Executable (PyInstaller)

This project includes a Windows batch file to build a single-file executable.

### 11.1 Prerequisites

```bash
python -m pip install pyinstaller
```

### 11.2 Build Command (Batch File)

Run the batch file from the project root:

```bash
scripts\pyinstaller_command.bat
```

You can also double-click `scripts\pyinstaller_command.bat` in File Explorer.

The batch file uses `python -m PyInstaller`, so PyInstaller does not need to be on PATH (but it must be installed in the active Python environment).

### 11.3 Output Details

The build outputs are written into the `scripts` folder:

- `scripts\qamRoster.exe` (the Windows 10/11 executable)
- `scripts\qamRoster.spec` (PyInstaller spec file)
- `scripts\build\` (temporary build artifacts)

If you want a different executable name, edit `scripts\pyinstaller_command.bat` and change the `--name` value.
