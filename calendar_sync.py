from __future__ import annotations

import hashlib
import re
from datetime import date, datetime
from pathlib import Path


GOOGLE_SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]
DATE_PATTERN = re.compile(
    r"(?P<start>\d{2}\.\d{2}\.\d{2})"
    r"(?:\s*~\s*(?P<end>\d{2}\.\d{2}\.\d{2}))?"
)


def _google_service(app_dir: Path, config: dict):
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Google 연결 구성요소가 없습니다. install_windows.bat을 다시 실행해주세요."
        ) from exc

    token_path = app_dir / config.get(
        "google_token_path", "app_data/google_token.json"
    )
    if not token_path.exists():
        raise RuntimeError(
            "Google 로그인 정보가 없습니다. 네이버 가격 최신화에서 Google 로그인을 먼저 완료해주세요."
        )
    credentials = Credentials.from_authorized_user_file(
        str(token_path), GOOGLE_SCOPE
    )
    if credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
        token_path.write_text(credentials.to_json(), encoding="utf-8")
    if not credentials.valid:
        raise RuntimeError(
            "Google 로그인이 만료되었습니다. 네이버 가격 최신화에서 다시 로그인해주세요."
        )
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def _parse_date(value: str) -> date:
    return datetime.strptime(value, "%y.%m.%d").date()


def fetch_sheet_events(app_dir: Path, config: dict) -> list[dict]:
    service = _google_service(app_dir, config)
    spreadsheet_id = str(config["spreadsheet_id"])
    sheet_name = str(config["sheet_name"]).replace("'", "''")
    read_range = str(config.get("range", "A1:V1030"))
    response = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!{read_range}",
            valueRenderOption="FORMATTED_VALUE",
        )
        .execute()
    )
    parsed_events: list[dict] = []
    seen: set[tuple[str, str, str]] = set()
    for row in response.get("values", []):
        for cell in row:
            if not isinstance(cell, str):
                continue
            match = DATE_PATTERN.search(cell)
            if match is None:
                continue
            event_name = cell[: match.start()].strip(" _")
            if not event_name:
                continue
            start = _parse_date(match.group("start"))
            end = _parse_date(match.group("end") or match.group("start"))
            if end < start:
                continue
            dedupe_key = (
                event_name.casefold(),
                start.isoformat(),
                end.isoformat(),
            )
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            source_text = (
                f"{event_name}|{start.isoformat()}|{end.isoformat()}"
            )
            source_key = hashlib.sha1(
                source_text.encode("utf-8")
            ).hexdigest()
            if "팝업" in event_name:
                event_type = "popup"
            elif "공구" in event_name or "공동구매" in event_name:
                event_type = "group_buy"
            else:
                event_type = "special"
            parsed_events.append(
                {
                    "source_key": source_key,
                    "date": start.isoformat(),
                    "end_date": end.isoformat(),
                    "text": event_name,
                    "event_type": event_type,
                }
            )
    parsed_events.sort(
        key=lambda item: (item["date"], item["end_date"], item["text"])
    )
    return parsed_events


def merge_sheet_events(
    tasks: list[dict],
    sheet_events: list[dict],
    spreadsheet_id: str,
) -> tuple[list[dict], dict]:
    """Replace every calendar event with the current sheet snapshot."""
    preserved = [item for item in tasks if item.get("kind") != "event"]
    removed = len(tasks) - len(preserved)
    now = datetime.now().isoformat(timespec="seconds")
    synced = []
    for event in sheet_events:
        event_id = hashlib.sha1(
            f"{spreadsheet_id}|{event['source_key']}".encode("utf-8")
        ).hexdigest()
        synced.append(
            {
                "id": event_id,
                "date": event["date"],
                "end_date": event["end_date"],
                "text": event["text"],
                "kind": "event",
                "event_type": event["event_type"],
                "source": "google_sheet",
                "source_spreadsheet_id": spreadsheet_id,
                "source_key": event["source_key"],
                "synced_at": now,
                "created_at": now,
                "completions": {},
                "statuses": {},
            }
        )
    merged = preserved + synced
    merged.sort(
        key=lambda item: (
            str(item.get("date", "")),
            str(item.get("created_at", "")),
        )
    )
    return merged, {
        "total": len(synced),
        "added": len(synced),
        "updated": 0,
        "removed": removed,
        "synced_at": now,
    }
