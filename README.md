# QAM-roster-automation

Create, delete, and modify Google Calendar events programmatically

---

## Project Metadata
- **Version:** 0.1.3
- **Created:** 2026-03-08
- **Author:** Wayne Freestun

---

## Overview
This project was created using the New Project Wizard and follows the standard ecosystem structure.

## Python Roster Sync
The project now includes a modular Python workflow for syncing roster DOCX files to Google Calendar.

### Install
```bash
python -m pip install -r requirements.txt
```

Copy local OAuth client credentials (do not commit real credentials):
```bash
copy modules\credentials.example.json modules\credentials.json
```
Then edit `modules/credentials.json` with your real Google OAuth values.

### Run
```bash
python main.py path\to\Roster_March_2026_v7.docx --dry-run
python main.py path\to\Roster_March_2026_v7.docx --replace-current
python main.py --gui
python main.py --cli path\to\Roster_March_2026_v7.docx --dry-run
python main.py --cli path\to\Roster_March_2026_v7.docx --sync-all-workers --all-workers-calendar-id your_other_calendar_id
```

When launched on Windows with no workflow arguments, the GUI opens by default.

### Source DOCX Selection
- Command-line: pass the file path as the first argument.
- Windows drag-and-drop: drag a `.docx` file onto the script/launcher; the dropped path is used as `docx_file`.
- UI file-picker: run without a DOCX argument in an interactive terminal, or force it with:
```bash
python main.py --select-file
```
- Disable picker fallback for automation:
```bash
python main.py --no-file-picker --calendar-id your_calendar_id path\to\Roster.docx
```

### GUI Features
- Minimal Windows GUI (`python main.py --gui`)
- DOCX file picker
- Google Calendar list auto-loads on startup
- `Refresh List` button reloads calendars from Google
- Change preview before execution
- Confirmation prompt before proceeding
- Dry run toggle
- Action modes:
  - Full sync (delete old + add new)
  - Delete previous version only
  - Delete current version (from selected DOCX)
  - Delete all versions in chosen month
  - Add new version only

### Notes
- OAuth client JSON lookup order:
  1. `QAM_GOOGLE_CLIENT_SECRET_PATH`
  2. `modules/credentials.json`
  3. `credentials.json`
  4. Legacy `modules/client_secret_*.apps.googleusercontent.com.json` (if exactly one file exists)
- `token.json` is created in the project root after first authentication and should remain local only.
- Logs are written to `logs/creation.log`.
- Calendar targeting priority:
  1. `--calendar-id`
  2. `QAM_CALENDAR_ID`
  3. Interactive picker UI (default in terminal sessions)
  4. `primary` (only when `--no-calendar-prompt` is used or non-interactive session)
- Use `--select-calendar` to force the picker UI.
- Use `--no-calendar-prompt` for unattended runs.
- Multi-month rosters are supported. A roster labeled for one month can now create events in the previous, current, and next month when day cells include an explicit month name such as `1. April ...`.
- Event summaries keep the roster month/version label so cross-month events are still traceable back to the source roster version.
- The optional all-workers flow creates one event per roster day on a separate Google Calendar and only runs when `--sync-all-workers` is enabled in CLI or the matching checkbox is enabled in the GUI.

