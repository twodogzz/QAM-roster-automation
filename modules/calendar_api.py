"""Google Calendar API wrapper for QAM roster automation."""

from __future__ import annotations

from pathlib import Path
import logging
import os
from typing import Any, cast

try:
    from google.auth.credentials import Credentials as GoogleCredentials
    from google.auth.exceptions import RefreshError
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials as OAuthUserCredentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
except ModuleNotFoundError as exc:
    raise ModuleNotFoundError(
        "Missing Google API dependencies. Install them with: "
        "python -m pip install -r requirements.txt"
    ) from exc


SCOPES = ["https://www.googleapis.com/auth/calendar"]
CLIENT_SECRET_ENV_VARS = (
    "QAM_GOOGLE_CLIENT_SECRET_PATH",
    "GOOGLE_OAUTH_CLIENT_SECRETS",
)
DEFAULT_CLIENT_SECRET_CANDIDATES = (
    "modules/credentials.json",
    "credentials.json",
)

LOGGER = logging.getLogger(__name__)
_SERVICE: Any | None = None
_CALENDAR_ID: str | None = None


def create_event(event_data: dict[str, Any]) -> dict[str, Any]:
    """Insert an event into the configured Google Calendar."""
    service = _get_service()
    created = (
        service.events()
        .insert(calendarId=_get_calendar_id(), body=event_data)
        .execute()
    )
    LOGGER.info("Created event id=%s summary=%s", created.get("id"), created.get("summary"))
    return created


def delete_events_by_query(query_text: str) -> int:
    """Delete all events matching a free-text query from the configured calendar."""
    service = _get_service()
    events = list_events(query_text)
    deleted = 0
    for event in events:
        event_id = event.get("id")
        if not event_id:
            continue
        service.events().delete(calendarId=_get_calendar_id(), eventId=event_id).execute()
        deleted += 1
        LOGGER.info("Deleted event id=%s summary=%s", event_id, event.get("summary"))
    return deleted


def delete_event_by_id(event_id: str) -> None:
    """Delete a single event by id from the configured calendar."""
    service = _get_service()
    service.events().delete(calendarId=_get_calendar_id(), eventId=event_id).execute()
    LOGGER.info("Deleted event id=%s", event_id)


def update_event(event_id: str, updated_fields: dict[str, Any]) -> dict[str, Any]:
    """Update an existing event using partial fields."""
    service = _get_service()
    updated = (
        service.events()
        .patch(calendarId=_get_calendar_id(), eventId=event_id, body=updated_fields)
        .execute()
    )
    LOGGER.info("Updated event id=%s summary=%s", event_id, updated.get("summary"))
    return updated


def list_events(query_text: str) -> list[dict[str, Any]]:
    """List events matching a free-text query from the configured calendar."""
    service = _get_service()
    response = (
        service.events()
        .list(
            calendarId=_get_calendar_id(),
            q=query_text,
            singleEvents=True,
            orderBy="startTime",
            maxResults=2500,
        )
        .execute()
    )
    return cast(list[dict[str, Any]], response.get("items", []))


def list_calendars() -> list[dict[str, Any]]:
    """Return available calendars from the authenticated account."""
    service = _get_service()
    calendars: list[dict[str, Any]] = []
    page_token = None
    while True:
        response = service.calendarList().list(pageToken=page_token).execute()
        calendars.extend(response.get("items", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    return calendars


def set_calendar_id(calendar_id: str) -> None:
    """Set the target calendar ID for API calls in this process."""
    global _CALENDAR_ID
    _CALENDAR_ID = calendar_id
    LOGGER.info("Target calendar set to id=%s", calendar_id)


def _get_service() -> Any:
    global _SERVICE
    if _SERVICE is not None:
        return _SERVICE

    creds = _load_credentials()
    _SERVICE = build("calendar", "v3", credentials=creds)
    return _SERVICE


def _load_credentials() -> GoogleCredentials:
    project_root = Path(__file__).resolve().parents[1]
    token_path = project_root / "token.json"
    client_secret_path = _resolve_client_secret_path(project_root)

    creds: GoogleCredentials | None = None
    if token_path.exists():
        creds = cast(GoogleCredentials, OAuthUserCredentials.from_authorized_user_file(str(token_path), SCOPES))

    if creds and creds.valid:
        return creds

    if creds and creds.expired and getattr(creds, "refresh_token", None):
        try:
            creds.refresh(Request())
        except RefreshError as exc:
            if not _should_force_reauth(exc):
                raise
            LOGGER.warning("OAuth refresh failed (%s). Re-authenticating with a new browser login.", exc)
            _delete_stale_token(token_path)
            creds = None
    else:
        creds = None

    if creds is None or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(client_secret_path), SCOPES)
        creds = cast(GoogleCredentials, flow.run_local_server(port=0))

    if creds is None:
        raise RuntimeError("OAuth credentials could not be created.")

    to_json = getattr(creds, "to_json", None)
    if callable(to_json):
        json_text = to_json()
        if isinstance(json_text, str):
            token_path.write_text(json_text, encoding="utf-8")
            LOGGER.info("OAuth token updated: %s", token_path)
        else:
            LOGGER.warning("Credential to_json() did not return str; token file was not updated.")
    else:
        LOGGER.warning("Credential type does not support to_json(); token file was not updated.")
    return creds


def _get_calendar_id() -> str:
    if _CALENDAR_ID:
        return _CALENDAR_ID
    return os.getenv("QAM_CALENDAR_ID", "primary")


def _resolve_client_secret_path(project_root: Path) -> Path:
    for env_var in CLIENT_SECRET_ENV_VARS:
        env_value = os.getenv(env_var, "").strip()
        if not env_value:
            continue
        candidate = Path(env_value).expanduser()
        if not candidate.is_absolute():
            candidate = project_root / candidate
        candidate = candidate.resolve()
        if candidate.exists():
            return candidate
        raise FileNotFoundError(f"{env_var} points to a missing OAuth client JSON file: {candidate}")

    for relative_path in DEFAULT_CLIENT_SECRET_CANDIDATES:
        candidate = (project_root / relative_path).resolve()
        if candidate.exists():
            return candidate

    legacy_candidates = sorted((project_root / "modules").glob("client_secret_*.apps.googleusercontent.com.json"))
    if len(legacy_candidates) == 1:
        return legacy_candidates[0].resolve()
    if len(legacy_candidates) > 1:
        raise RuntimeError(
            "Multiple legacy Google OAuth JSON files found in modules/. "
            "Keep only one, or set QAM_GOOGLE_CLIENT_SECRET_PATH explicitly."
        )

    raise FileNotFoundError(
        "Google OAuth client JSON not found. "
        "Create modules/credentials.json (recommended), "
        "or set QAM_GOOGLE_CLIENT_SECRET_PATH to your local client secret JSON file."
    )


def _should_force_reauth(exc: RefreshError) -> bool:
    text = str(exc).lower()
    return (
        "invalid_client" in text
        or "invalid_grant" in text
        or "unauthorized" in text
    )


def _delete_stale_token(token_path: Path) -> None:
    try:
        if token_path.exists():
            token_path.unlink()
            LOGGER.info("Removed stale OAuth token file: %s", token_path)
    except OSError as exc:
        LOGGER.warning("Could not remove stale OAuth token file %s: %s", token_path, exc)
