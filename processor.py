from __future__ import annotations

import calendar
import datetime
import logging
import re
from dataclasses import dataclass
from pathlib import Path

from googleapiclient.errors import HttpError
from google.auth.exceptions import RefreshError

from excel_reader import ExcelReadError, read_excel
from sheets_client import (
    apply_green_fill,
    build_sheets_service,
    fetch_row_values,
    fetch_sheet_infos,
    find_mp_sheet,
    get_last_filled_row,
    group_imported_rows,
    insert_row,
    update_summary_sheet,
    update_summary_row,
    update_formulas,
    update_values,
)

logger = logging.getLogger(__name__)

MONTHS = [
    "январь",
    "февраль",
    "март",
    "апрель",
    "май",
    "июнь",
    "июль",
    "август",
    "сентябрь",
    "октябрь",
    "ноябрь",
    "декабрь",
]
MONTH_REGEX = re.compile(rf"({'|'.join(MONTHS)})\s+(\d{{4}})", re.IGNORECASE)

STORE_ALIASES = {
    "Авиаторов": ["авиаторов"],
    "Козловская": ["козловская", "санвэй", "санвей"],
    "Диамант": ["диамант", "цитрус"],
    "Привоз": ["привоз"],
    "Бахтурова": ["бахтурова"],
    "Ахтубинск": ["ахтубинск"],
    "СтройГрад": ["стройград", "строй град"],
    "Европа": ["европа"],
    "Парк Хаус": ["парк хаус", "паркхаус", "пх"],
    "ЦУМ": ["цум", "советница"],
    "Простор": ["простор"],
}


@dataclass
class FileContext:
    path: Path
    store: str
    period: str


def process_directory(
    input_dir: str,
    period: str | None,
    spreadsheet_id: str,
    credentials: str,
    dry_run: bool,
    progress_callback=None,
) -> None:
    directory = Path(input_dir)
    if not directory.exists():
        logger.error("Папка не найдена: %s", directory)
        return

    files = sorted(directory.glob("*.xls")) + sorted(directory.glob("*.xlsx"))
    if not files:
        logger.warning("В папке нет файлов .xls или .xlsx")
        return

    service = None
    sheet_infos = []
    if not dry_run:
        if not Path(credentials).exists():
            logger.error("Файл credentials не найден: %s", credentials)
            return
        try:
            service = build_sheets_service(credentials)
            sheet_infos = fetch_sheet_infos(service, spreadsheet_id)
        except RefreshError as exc:
            logger.error(
                "Ошибка авторизации Google (RefreshError). Проверьте credentials и доступ к таблице: %s",
                exc,
            )
            return
        except HttpError as exc:
            logger.error("Ошибка доступа к Google Sheets API: %s", exc)
            return

    total_files = len(files)
    for index, file_path in enumerate(files, start=1):
        try:
            context = _build_context(file_path, period, dry_run)
        except ValueError as exc:
            logger.error("%s: %s", file_path.name, exc)
            continue

        logger.info(
            "Файл: %s | Магазин: %s | Период: %s",
            file_path.name,
            context.store,
            context.period,
        )

        try:
            excel_data = read_excel(context.path)
        except ExcelReadError as exc:
            logger.error("%s: %s", file_path.name, exc)
            continue

        logger.info(
            "Найдены колонки: Чеки=%s, Товары=%s, Подарочные сертификаты=%s",
            excel_data.column_map["checks"],
            excel_data.column_map["goods"],
            excel_data.column_map["gift_cert"],
        )

        if not excel_data.rows:
            logger.warning("%s: нет данных для переноса", file_path.name)
            continue

        rows_to_write = _prepare_rows(excel_data.rows, context.period)
        if not rows_to_write:
            logger.warning("%s: после фильтрации нет данных для переноса", file_path.name)
            continue

        if dry_run:
            logger.info("[DRY RUN] Перенесли бы %s строк", len(rows_to_write))
            continue

        sheet_info = find_mp_sheet(sheet_infos, context.store)
        if not sheet_info:
            logger.error("%s: не найден лист МП для магазина '%s'", file_path.name, context.store)
            continue

        last_row = get_last_filled_row(service, spreadsheet_id, sheet_info.title)
        summary_row = last_row + 1
        data_start = summary_row + 1
        data_end = summary_row + len(rows_to_write)

        logger.info(
            "Запись в лист '%s': строки %s-%s",
            sheet_info.title,
            data_start,
            data_end,
        )

        insert_row(service, spreadsheet_id, sheet_info.sheet_id, summary_row)
        apply_green_fill(service, spreadsheet_id, sheet_info.sheet_id, summary_row)
        period_label = context.period.split()[0]
        try:
            update_summary_row(
                service,
                spreadsheet_id,
                sheet_info.title,
                summary_row,
                period_label,
                data_start,
                data_end,
            )
        except HttpError as exc:
            logger.error("Ошибка обновления сводной строки в '%s': %s", sheet_info.title, exc)
            continue
        update_values(service, spreadsheet_id, sheet_info.title, data_start, rows_to_write)
        update_formulas(service, spreadsheet_id, sheet_info.title, data_start, data_end)
        group_imported_rows(
            service,
            spreadsheet_id,
            sheet_info.sheet_id,
            start_row_1based=data_start,
            end_row_1based=data_end,
            excluded_row_1based=summary_row,
        )
        summary_values = fetch_row_values(service, spreadsheet_id, sheet_info.title, summary_row)
        update_summary_sheet(
            service,
            spreadsheet_id,
            sheet_infos,
            sheet_info.title,
            summary_values,
            _format_period_label(period),
        )

        logger.info("%s: успешно перенесено строк: %s", file_path.name, len(rows_to_write))
        if progress_callback:
            progress_callback(index, total_files, file_path.name)


def _build_context(file_path: Path, fallback_period: str | None, dry_run: bool) -> FileContext:
    store = _detect_store(file_path.stem)
    if not store:
        raise ValueError("Не удалось определить магазин по названию")

    detected_period = _detect_period(file_path.stem)
    period = detected_period or fallback_period
    if not period:
        raise ValueError("Не найден период в названии и не указан период вручную")

    new_path = _maybe_rename(file_path, period, detected_period, dry_run)
    return FileContext(path=new_path, store=store, period=period)


def _detect_store(filename: str) -> str | None:
    lower_name = filename.lower()
    for store, aliases in STORE_ALIASES.items():
        if any(alias in lower_name for alias in aliases):
            return store
    return None


def _detect_period(filename: str) -> str | None:
    match = MONTH_REGEX.search(filename)
    if not match:
        return None
    month = match.group(1)
    year = match.group(2)
    return f"{_capitalize_month(month)} {year}"


def _capitalize_month(month: str) -> str:
    month = month.strip()
    return month[:1].upper() + month[1:]


def _maybe_rename(
    file_path: Path,
    period: str,
    detected_period: str | None,
    dry_run: bool,
) -> Path:
    if detected_period:
        return file_path

    new_name = f"{file_path.stem} {period}{file_path.suffix}"
    new_path = file_path.with_name(new_name)

    if dry_run:
        logger.info("[DRY RUN] Переименовали бы файл в %s", new_name)
        return file_path

    logger.info("Переименование файла: %s -> %s", file_path.name, new_name)
    file_path.rename(new_path)
    return new_path


def _prepare_rows(rows: list[list], period: str) -> list[list]:
    """Форматирует дату и ограничивает количество строк по числу дней в месяце."""
    year, month = _parse_period(period)
    days_in_month = calendar.monthrange(year, month)[1]
    prepared: list[list] = []
    for row in rows:
        if len(prepared) >= days_in_month:
            break
        new_row = list(row)
        new_row[0] = _format_date_value(new_row[0])
        prepared.append(new_row)
    return prepared


def _parse_period(period: str) -> tuple[int, int]:
    """Парсит период в формате 'Месяц ГГГГ'."""
    parts = period.split()
    if len(parts) < 2:
        raise ValueError(f"Некорректный период: {period}")
    month_name = parts[0].lower()
    year = int(parts[1])
    month_map = {
        "январь": 1,
        "февраль": 2,
        "март": 3,
        "апрель": 4,
        "май": 5,
        "июнь": 6,
        "июль": 7,
        "август": 8,
        "сентябрь": 9,
        "октябрь": 10,
        "ноябрь": 11,
        "декабрь": 12,
    }
    if month_name not in month_map:
        raise ValueError(f"Неизвестный месяц в периоде: {period}")
    return year, month_map[month_name]


def _format_period_label(period: str) -> str:
    """Возвращает период в формате ММ-ГГГГ."""
    year, month = _parse_period(period)
    return f"{month:02d}-{year}"


def _format_date_value(value) -> str | None:
    """Приводит дату к формату ДД.ММ.ГГГГ."""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, datetime.date):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, (int, float)):
        try:
            base = datetime.datetime(1899, 12, 30)
            converted = base + datetime.timedelta(days=float(value))
            return converted.strftime("%d.%m.%Y")
        except (OverflowError, ValueError):
            return str(value)
    text = str(value).strip()
    if not text:
        return None
    if _is_number(text):
        try:
            base = datetime.datetime(1899, 12, 30)
            converted = base + datetime.timedelta(days=float(text))
            return converted.strftime("%d.%m.%Y")
        except (OverflowError, ValueError):
            return text
    for fmt in ("%d.%m.%Y", "%d.%m.%y"):
        try:
            parsed = datetime.datetime.strptime(text, fmt)
            return parsed.strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def _is_number(value: str) -> bool:
    try:
        float(value.replace(",", "."))
        return True
    except ValueError:
        return False
