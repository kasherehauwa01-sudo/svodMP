from __future__ import annotations

import dataclasses
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


KEYWORDS = {
    "checks": "Чеки",
    "goods": "Товары",
    "gift_cert": "Подарочные сертификаты",
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

    header_row_index = HEADER_ROWS[0]
    column_map = _find_keyword_columns_xlsx(sheet, HEADER_ROWS)
    data_start_row = _find_data_start_row_xlsx(sheet)
    data_end_row = _find_data_end_row_xlsx(sheet, data_start_row)

    rows = _extract_rows_xlsx(sheet, data_start_row, data_end_row, column_map)

    return ExcelData(
        rows=rows,
        header_row_index=header_row_index,
        data_start_row=data_start_row,
        data_end_row=data_end_row,
        column_map=column_map,
    )


def _read_xls(file_path: Path) -> ExcelData:
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)

    header_row_index = HEADER_ROWS[0]
    column_map = _find_keyword_columns_xls(sheet, HEADER_ROWS)
    data_start_row = _find_data_start_row_xls(sheet)
    data_end_row = _find_data_end_row_xls(sheet, data_start_row)

    rows = _extract_rows_xls(sheet, data_start_row, data_end_row, column_map)

    return ExcelData(
        rows=rows,
        header_row_index=header_row_index,
        data_start_row=data_start_row,
        data_end_row=data_end_row,
        column_map=column_map,
    )


def _find_keyword_columns_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    header_rows: list[int],
) -> dict[str, int]:
    max_col = sheet.max_column
    column_map: dict[str, int] = {}

    for header_row in header_rows:
        for col in range(1, max_col + 1):
            value = sheet.cell(row=header_row, column=col).value
            text = _normalize_header_value(value)
            if not text:
                continue
            for key, keyword in KEYWORDS.items():
                if key in column_map:
                    continue
                if keyword.lower() in text:
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
            value = sheet.cell_value(row_index, col)
            text = _normalize_header_value(value)
            if not text:
                continue
            for key, keyword in KEYWORDS.items():
                if key in column_map:
                    continue
                if keyword.lower() in text:
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


def _find_data_start_row_xlsx(sheet: openpyxl.worksheet.worksheet.Worksheet) -> int:
    for row in range(1, sheet.max_row + 1):
        value = sheet.cell(row=row, column=1).value
        if _is_date_header(value):
            return row + 1
    raise ExcelReadError("Не найдена строка с заголовком 'Дата' в колонке A")


def _find_data_start_row_xls(sheet: xlrd.sheet.Sheet) -> int:
    for row in range(sheet.nrows):
        value = sheet.cell_value(row, 0)
        if _is_date_header(value):
            return row + 2
    raise ExcelReadError("Не найдена строка с заголовком 'Дата' в колонке A")


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
) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for row in range(start_row, end_row + 1):
        rows.append(_build_row_xlsx(sheet, row, column_map))
    return rows


def _extract_rows_xls(
    sheet: xlrd.sheet.Sheet,
    start_row: int,
    end_row: int,
    column_map: dict[str, int],
) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for row in range(start_row - 1, end_row):
        rows.append(_build_row_xls(sheet, row, column_map))
    return rows


def _build_row_xlsx(
    sheet: openpyxl.worksheet.worksheet.Worksheet, row: int, column_map: dict[str, int]
) -> list[Any]:
    values = [
        sheet.cell(row=row, column=1).value,
        sheet.cell(row=row, column=2).value,
        sheet.cell(row=row, column=3).value,
        sheet.cell(row=row, column=column_map["checks"] + 1).value,
        None,
        sheet.cell(row=row, column=column_map["goods"] + 1).value,
        sheet.cell(row=row, column=5).value,
        sheet.cell(row=row, column=column_map["gift_cert"] + 1).value,
    ]
    return values


def _build_row_xls(sheet: xlrd.sheet.Sheet, row: int, column_map: dict[str, int]) -> list[Any]:
    values = [
        sheet.cell_value(row, 0),
        sheet.cell_value(row, 1),
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


def _normalize_header_value(value: Any) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip().lower()
    return text if text else None


def _is_date_header(value: Any) -> bool:
    if value is None:
        return False
    text = str(value).strip().lower()
    return "дата" in text


def _is_empty_value(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value == "":
        return True
    return False
