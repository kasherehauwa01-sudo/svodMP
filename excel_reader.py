from __future__ import annotations

import dataclasses
import datetime
import logging
from pathlib import Path
from typing import Any, Iterable, Optional

import openpyxl
import xlrd

logger = logging.getLogger(__name__)


@dataclasses.dataclass
class ExcelData:
    rows: list[list[Any]]
    header_row_index: int
    data_start_row: int
    data_end_row: int
    column_map: dict[str, int]
    date_col: int
    day_col: int


KEYWORDS = {
    "checks": "Чеки",
    "goods": "Товары",
    "gift_cert": "Подарочные сертификаты",
}
KEYWORD_ALIASES = {
    "goods": ["штуки"],
}
HEADER_ROWS = [2, 3, 4, 5]


class ExcelReadError(Exception):
    pass


def read_excel(file_path: Path) -> ExcelData:
    if file_path.suffix.lower() == ".xlsx":
        return _read_xlsx(file_path)
    if file_path.suffix.lower() == ".xls":
        return _read_xls(file_path)
    raise ExcelReadError(f"Неподдерживаемое расширение: {file_path.suffix}")


def _read_xlsx(file_path: Path) -> ExcelData:
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    sheet = workbook.active

    data_start_row, date_col, day_col = _find_data_start_row_xlsx(sheet)
    header_rows = _build_header_rows(data_start_row)
    header_row_index = header_rows[0] if header_rows else HEADER_ROWS[0]
    column_map = _find_keyword_columns_xlsx(sheet, header_rows or HEADER_ROWS)
    data_end_row = _find_data_end_row_xlsx(sheet, data_start_row)

    rows = _extract_rows_xlsx(sheet, data_start_row, data_end_row, column_map, date_col, day_col)

    return ExcelData(
        rows=rows,
        header_row_index=header_row_index,
        data_start_row=data_start_row,
        data_end_row=data_end_row,
        column_map=column_map,
        date_col=date_col,
        day_col=day_col,
    )


def _read_xls(file_path: Path) -> ExcelData:
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)

    data_start_row, date_col, day_col = _find_data_start_row_xls(sheet)
    header_rows = _build_header_rows(data_start_row)
    header_row_index = header_rows[0] if header_rows else HEADER_ROWS[0]
    column_map = _find_keyword_columns_xls(sheet, header_rows or HEADER_ROWS)
    data_end_row = _find_data_end_row_xls(sheet, data_start_row)

    rows = _extract_rows_xls(sheet, data_start_row, data_end_row, column_map, date_col, day_col)

    return ExcelData(
        rows=rows,
        header_row_index=header_row_index,
        data_start_row=data_start_row,
        data_end_row=data_end_row,
        column_map=column_map,
        date_col=date_col,
        day_col=day_col,
    )


def _find_keyword_columns_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    header_rows: list[int],
) -> dict[str, int]:
    max_col = sheet.max_column
    column_map: dict[str, int] = {}

    for header_row in header_rows:
        for col in range(1, max_col + 1):
            text = _get_header_text_xlsx(sheet, header_row, col)
            if not text:
                continue
            for key, keyword in KEYWORDS.items():
                if key in column_map:
                    continue
                if _keyword_in_text(key, keyword, text):
                    left_col = _get_merge_left_col_xlsx(sheet, header_row, col)
                    column_map[key] = left_col - 1
        if len(column_map) == len(KEYWORDS):
            break

    _validate_column_map(column_map, header_rows)
    return column_map


def _get_merge_left_col_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet, row: int, col: int
) -> int:
    for cell_range in sheet.merged_cells.ranges:
        if cell_range.min_row <= row <= cell_range.max_row and cell_range.min_col <= col <= cell_range.max_col:
            return cell_range.min_col
    return col


def _find_keyword_columns_xls(sheet: xlrd.sheet.Sheet, header_rows: list[int]) -> dict[str, int]:
    column_map: dict[str, int] = {}

    for header_row in header_rows:
        row_index = header_row - 1
        for col in range(sheet.ncols):
            text = _get_header_text_xls(sheet, row_index, col)
            if not text:
                continue
            for key, keyword in KEYWORDS.items():
                if key in column_map:
                    continue
                if _keyword_in_text(key, keyword, text):
                    left_col = _get_merge_left_col_xls(sheet, row_index, col)
                    column_map[key] = left_col
        if len(column_map) == len(KEYWORDS):
            break

    _validate_column_map(column_map, header_rows)
    return column_map


def _get_merge_left_col_xls(sheet: xlrd.sheet.Sheet, row: int, col: int) -> int:
    for rlo, rhi, clo, chi in sheet.merged_cells:
        if rlo <= row < rhi and clo <= col < chi:
            return clo
    return col


def _find_data_start_row_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
) -> tuple[int, int, int]:
    max_row = sheet.max_row
    date_col = 0
    day_col = 1
    for row in range(1, max_row + 1):
        value = _get_header_text_xlsx(sheet, row, 1)
        if _is_date_header(value):
            return row + 1, date_col, day_col
    for row in range(1, max_row + 1):
        for col in range(1, min(sheet.max_column, 26) + 1):
            value = _get_header_text_xlsx(sheet, row, col)
            if _is_date_header(value):
                return row + 1, col - 1, col
    day_header = _find_day_header_xlsx(sheet)
    if day_header:
        header_row, day_header_col = day_header
        date_col = max(day_header_col - 2, 0)
        day_col = day_header_col - 1
        return header_row + 1, date_col, day_col
    fallback_row = _find_date_like_row_xlsx(sheet)
    if fallback_row:
        return fallback_row, date_col, day_col
    raise ExcelReadError("Не найдена строка с заголовком 'Дата' в диапазоне A:H")


def _find_data_start_row_xls(sheet: xlrd.sheet.Sheet) -> tuple[int, int, int]:
    date_col = 0
    day_col = 1
    for row in range(sheet.nrows):
        value = _get_header_text_xls(sheet, row, 0)
        if _is_date_header(value):
            return row + 2, date_col, day_col
    max_col = min(sheet.ncols, 26)
    for row in range(sheet.nrows):
        for col in range(max_col):
            value = _get_header_text_xls(sheet, row, col)
            if _is_date_header(value):
                return row + 2, col, col + 1
    day_header = _find_day_header_xls(sheet)
    if day_header:
        header_row, day_header_col = day_header
        date_col = max(day_header_col - 1, 0)
        day_col = day_header_col
        return header_row + 2, date_col, day_col
    fallback_row = _find_date_like_row_xls(sheet)
    if fallback_row:
        return fallback_row, date_col, day_col
    raise ExcelReadError("Не найдена строка с заголовком 'Дата' в диапазоне A:H")


def _find_data_end_row_xlsx(sheet: openpyxl.worksheet.worksheet.Worksheet, start_row: int) -> int:
    last_row = sheet.max_row
    return _trim_trailing_empty_rows(
        range(start_row, last_row + 1),
        lambda r: _row_has_data_xlsx(sheet, r),
    )


def _find_data_end_row_xls(sheet: xlrd.sheet.Sheet, start_row: int) -> int:
    last_row = sheet.nrows
    return _trim_trailing_empty_rows(
        range(start_row, last_row + 1),
        lambda r: _row_has_data_xls(sheet, r - 1),
    )


def _trim_trailing_empty_rows(rows: Iterable[int], has_data) -> int:
    rows = list(rows)
    for row in reversed(rows):
        if has_data(row):
            return row
    return rows[0] - 1


def _row_has_data_xlsx(sheet: openpyxl.worksheet.worksheet.Worksheet, row: int) -> bool:
    for col in range(1, 9):
        value = sheet.cell(row=row, column=col).value
        if not _is_empty_value(value):
            return True
    return False


def _row_has_data_xls(sheet: xlrd.sheet.Sheet, row: int) -> bool:
    for col in range(8):
        value = sheet.cell_value(row, col)
        if not _is_empty_value(value):
            return True
    return False


def _extract_rows_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    start_row: int,
    end_row: int,
    column_map: dict[str, int],
    date_col: int,
    day_col: int,
) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for row in range(start_row, end_row + 1):
        rows.append(_build_row_xlsx(sheet, row, column_map, date_col, day_col))
    return rows


def _extract_rows_xls(
    sheet: xlrd.sheet.Sheet,
    start_row: int,
    end_row: int,
    column_map: dict[str, int],
    date_col: int,
    day_col: int,
) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for row in range(start_row - 1, end_row):
        rows.append(_build_row_xls(sheet, row, column_map, date_col, day_col))
    return rows


def _build_row_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    row: int,
    column_map: dict[str, int],
    date_col: int,
    day_col: int,
) -> list[Any]:
    values = [
        sheet.cell(row=row, column=date_col + 1).value,
        sheet.cell(row=row, column=day_col + 1).value,
        sheet.cell(row=row, column=3).value,
        sheet.cell(row=row, column=column_map["checks"] + 1).value,
        None,
        sheet.cell(row=row, column=column_map["goods"] + 1).value,
        sheet.cell(row=row, column=5).value,
        sheet.cell(row=row, column=column_map["gift_cert"] + 1).value,
    ]
    return values


def _build_row_xls(
    sheet: xlrd.sheet.Sheet,
    row: int,
    column_map: dict[str, int],
    date_col: int,
    day_col: int,
) -> list[Any]:
    values = [
        sheet.cell_value(row, date_col),
        sheet.cell_value(row, day_col),
        sheet.cell_value(row, 2),
        sheet.cell_value(row, column_map["checks"]),
        None,
        sheet.cell_value(row, column_map["goods"]),
        sheet.cell_value(row, 4),
        sheet.cell_value(row, column_map["gift_cert"]),
    ]
    return values


def _validate_column_map(column_map: dict[str, int], header_rows: list[int]) -> None:
    missing = [keyword for keyword in KEYWORDS if keyword not in column_map]
    if missing:
        raise ExcelReadError(
            "Не найдены заголовки в строках "
            f"{', '.join(str(row) for row in header_rows)}: "
            + ", ".join(KEYWORDS[key] for key in missing)
        )


def _build_header_rows(data_start_row: int) -> list[int]:
    if data_start_row <= 1:
        return []
    return list(range(1, data_start_row))


def _normalize_header_value(value: Any) -> Optional[str]:
    if value is None:
        return None
    text = str(value).replace("\u00a0", " ").strip().lower()
    text = " ".join(text.split())
    return text if text else None


def _keyword_in_text(key: str, keyword: str, text: str) -> bool:
    if keyword.lower() in text:
        return True
    aliases = KEYWORD_ALIASES.get(key, [])
    return any(alias.lower() in text for alias in aliases)


def _is_date_header(value: Any) -> bool:
    if value is None:
        return False
    text = str(value).strip().lower()
    return "дата" in text


def _is_day_header(value: Any) -> bool:
    if value is None:
        return False
    text = str(value).strip().lower()
    return "день нед" in text


def _find_day_header_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
) -> Optional[tuple[int, int]]:
    for row in range(1, sheet.max_row + 1):
        for col in range(1, min(sheet.max_column, 26) + 1):
            value = _get_header_text_xlsx(sheet, row, col)
            if _is_day_header(value):
                return row, col
    return None


def _find_day_header_xls(sheet: xlrd.sheet.Sheet) -> Optional[tuple[int, int]]:
    for row in range(sheet.nrows):
        for col in range(min(sheet.ncols, 26)):
            value = _get_header_text_xls(sheet, row, col)
            if _is_day_header(value):
                return row, col
    return None


def _is_empty_value(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value == "":
        return True
    return False


def _get_header_text_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    row: int,
    col: int,
) -> Optional[str]:
    value = sheet.cell(row=row, column=col).value
    text = _normalize_header_value(value)
    if text:
        return text
    for cell_range in sheet.merged_cells.ranges:
        if cell_range.min_row <= row <= cell_range.max_row and cell_range.min_col <= col <= cell_range.max_col:
            merged_value = sheet.cell(row=cell_range.min_row, column=cell_range.min_col).value
            return _normalize_header_value(merged_value)
    return None


def _get_header_text_xls(sheet: xlrd.sheet.Sheet, row: int, col: int) -> Optional[str]:
    value = sheet.cell_value(row, col)
    text = _normalize_header_value(value)
    if text:
        return text
    for rlo, rhi, clo, chi in sheet.merged_cells:
        if rlo <= row < rhi and clo <= col < chi:
            merged_value = sheet.cell_value(rlo, clo)
            return _normalize_header_value(merged_value)
    return None


def _find_date_like_row_xlsx(sheet: openpyxl.worksheet.worksheet.Worksheet) -> Optional[int]:
    max_row = sheet.max_row
    max_col = sheet.max_column
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            value = sheet.cell(row=row, column=col).value
            if _is_date_like_value(value):
                return row + 1
    return None


def _find_date_like_row_xls(sheet: xlrd.sheet.Sheet) -> Optional[int]:
    max_col = sheet.ncols
    for row in range(sheet.nrows):
        for col in range(max_col):
            value = sheet.cell_value(row, col)
            if _is_date_like_value(value):
                return row + 2
    return None


def _is_date_like_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, datetime.datetime):
        return True
    if isinstance(value, datetime.date):
        return True
    text = str(value).strip()
    if not text:
        return False
    for fmt in ("%d.%m.%y", "%d.%m.%Y"):
        try:
            datetime.datetime.strptime(text, fmt)
            return True
        except ValueError:
            continue
    return False
