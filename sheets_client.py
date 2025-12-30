from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any

from google.oauth2 import service_account
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


@dataclass
class SheetInfo:
    sheet_id: int
    title: str


def build_sheets_service(credentials_path: str):
    credentials = service_account.Credentials.from_service_account_file(
        credentials_path, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=credentials)


def fetch_sheet_infos(service, spreadsheet_id: str) -> list[SheetInfo]:
    response = (
        service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
    )
    sheet_infos = []
    for sheet in response.get("sheets", []):
        props = sheet.get("properties", {})
        sheet_infos.append(SheetInfo(sheet_id=props.get("sheetId"), title=props.get("title")))
    return sheet_infos


def find_mp_sheet(sheet_infos: list[SheetInfo], store_name: str) -> SheetInfo | None:
    store_lower = store_name.lower()
    candidates = [
        info
        for info in sheet_infos
        if info.title and "мп" in info.title.lower()
    ]

    for info in candidates:
        title_lower = info.title.lower()
        if store_lower == "цум":
            if "цум" in title_lower:
                return info
        if store_lower == "диамант":
            if "цитрус" in title_lower:
                return info
        if store_lower in title_lower:
            return info
    return None


def get_last_filled_row(service, spreadsheet_id: str, sheet_title: str) -> int:
    range_name = f"'{sheet_title}'!A:H"
    response = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_name)
        .execute()
    )
    values = response.get("values", [])
    last_row = 0
    for idx, row in enumerate(values, start=1):
        if any(cell not in (None, "") for cell in row):
            last_row = idx
    return last_row


def apply_green_fill(service, spreadsheet_id: str, sheet_id: int, row_index: int) -> None:
    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_index - 1,
                    "endRowIndex": row_index,
                    "startColumnIndex": 0,
                    "endColumnIndex": 8,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.76, "green": 0.87, "blue": 0.78}
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        }
    ]
    body = {"requests": requests}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def update_values(
    service,
    spreadsheet_id: str,
    sheet_title: str,
    start_row: int,
    rows: list[list[Any]],
) -> None:
    end_row = start_row + len(rows) - 1
    range_name = f"'{sheet_title}'!A{start_row}:H{end_row}"
    body = {"values": rows}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="RAW",
        body=body,
    ).execute()


def update_formulas(
    service,
    spreadsheet_id: str,
    sheet_title: str,
    start_row: int,
    end_row: int,
) -> None:
    formulas = [[f"=C{row}/D{row}"] for row in range(start_row, end_row + 1)]
    range_name = f"'{sheet_title}'!E{start_row}:E{end_row}"
    body = {"values": formulas}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
