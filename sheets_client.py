from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any, Optional

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
        if info.title and info.title.lower().strip().startswith("мп")
    ]
    keywords_map = {
        "цум": ["цум", "советница"],
        "диамант": ["диамант", "цитрус"],
        "козловская": ["козловская", "санвэй", "санвей"],
        "парк хаус": ["парк хаус", "паркхаус"],
        "стройград": ["стройград", "строй град"],
    }
    keywords = keywords_map.get(store_lower, [store_lower])
    for info in candidates:
        title_lower = info.title.lower()
        if any(keyword in title_lower for keyword in keywords):
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


def insert_row(service, spreadsheet_id: str, sheet_id: int, row_index: int) -> None:
    requests = [
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row_index - 1,
                    "endIndex": row_index,
                },
                "inheritFromBefore": False,
            }
        }
    ]
    body = {"requests": requests}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def update_summary_row(
    service,
    spreadsheet_id: str,
    sheet_title: str,
    summary_row: int,
    period_label: str,
    data_start: int,
    data_end: int,
) -> None:
    values = [
        [
            period_label,
            "",
            f"=SUM(C{data_start}:C{data_end})",
            f"=SUM(D{data_start}:D{data_end})",
            f"=AVERAGE(E{data_start}:E{data_end})",
            f"=SUM(F{data_start}:F{data_end})",
            f"=SUM(G{data_start}:G{data_end})",
            f"=SUM(H{data_start}:H{data_end})",
        ]
    ]
    range_name = f"'{sheet_title}'!A{summary_row}:H{summary_row}"
    body = {"values": values}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()


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


def fetch_row_values(
    service,
    spreadsheet_id: str,
    sheet_title: str,
    row_index: int,
) -> list[Any]:
    range_name = f"'{sheet_title}'!A{row_index}:H{row_index}"
    response = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
    )
    values = response.get("values", [])
    return values[0] if values else []


def update_summary_sheet(
    service,
    spreadsheet_id: str,
    sheet_infos: list[SheetInfo],
    source_sheet_title: str,
    source_row: list[Any],
    period_label: str,
) -> None:
    summary_info = _find_summary_sheet(sheet_infos)
    if not summary_info:
        logger.warning("Не найден лист 'Сводная'")
        return

    keyword = _extract_store_keyword(source_sheet_title)
    if not keyword:
        logger.warning("Не удалось определить магазин для листа '%s'", source_sheet_title)
        return

    block_start = _find_summary_block_start(service, spreadsheet_id, summary_info, keyword)
    if block_start is None:
        logger.warning("Не найдён блок '%s' в листе 'Сводная'", keyword)
        return

    block_end = block_start + 7
    target_row = _find_first_empty_row(
        service,
        spreadsheet_id,
        summary_info.title,
        block_start,
        block_end,
    )
    if target_row is None:
        logger.warning("Не удалось найти свободную строку для листа 'Сводная'")
        return

    values = [
        [
            period_label,
            _get_cell_value(source_row, 2),
            _get_cell_value(source_row, 3),
            _get_cell_value(source_row, 4),
            _get_cell_value(source_row, 5),
            _get_cell_value(source_row, 6),
            _get_cell_value(source_row, 7),
        ]
    ]
    range_name = (
        f"'{summary_info.title}'!"
        f"{_column_to_letter(block_start + 1)}{target_row}:"
        f"{_column_to_letter(block_start + 7)}{target_row}"
    )
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()


def _find_summary_sheet(sheet_infos: list[SheetInfo]) -> SheetInfo | None:
    for info in sheet_infos:
        if info.title and info.title.strip().lower() == "сводная":
            return info
    return None


def _extract_store_keyword(sheet_title: str) -> Optional[str]:
    title_lower = sheet_title.lower()
    candidates = [
        "авиаторов",
        "козловская",
        "цитрус",
        "привоз",
        "простор",
        "бахтурова",
        "ахтубинск",
        "стройград",
        "цум",
        "европа",
        "парк хаус",
    ]
    for keyword in candidates:
        if keyword in title_lower:
            return keyword
    return None


def _find_summary_block_start(
    service,
    spreadsheet_id: str,
    summary_info: SheetInfo,
    keyword: str,
) -> Optional[int]:
    response = (
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=f"'{summary_info.title}'!1:1",
            includeGridData=True,
            fields="sheets(properties,merges,data.rowData.values.effectiveValue,data.rowData.values.formattedValue)",
        )
        .execute()
    )
    sheets = response.get("sheets", [])
    if not sheets:
        return None
    sheet_data = sheets[0]
    merges = sheet_data.get("merges", [])
    row_data = sheet_data.get("data", [])
    row_values = []
    if row_data and row_data[0].get("rowData"):
        row_values = row_data[0]["rowData"][0].get("values", [])

    for merged in merges:
        if merged.get("startRowIndex") != 0:
            continue
        start_col = merged.get("startColumnIndex", 0)
        cell_text = _get_row_value_text(row_values, start_col)
        if keyword in cell_text:
            return start_col
    return None


def _get_row_value_text(values: list[dict], col_index: int) -> str:
    if col_index >= len(values):
        return ""
    value = values[col_index].get("effectiveValue") or {}
    text = (
        value.get("stringValue")
        or values[col_index].get("formattedValue")
        or ""
    )
    return _normalize_text(str(text))


def _find_first_empty_row(
    service,
    spreadsheet_id: str,
    sheet_title: str,
    start_col: int,
    end_col: int,
) -> Optional[int]:
    range_name = (
        f"'{sheet_title}'!"
        f"{_column_to_letter(start_col + 1)}2:{_column_to_letter(end_col)}"
    )
    response = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_name)
        .execute()
    )
    values = response.get("values", [])
    for idx, row in enumerate(values, start=2):
        if not any(cell not in (None, "") for cell in row):
            return idx
    return len(values) + 2


def _get_cell_value(row: list[Any], index: int) -> Any:
    if index < len(row):
        return row[index]
    return None


def _column_to_letter(column_index: int) -> str:
    result = ""
    col = column_index
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _normalize_text(text: str) -> str:
    return " ".join(text.replace("\u00a0", " ").lower().split())
