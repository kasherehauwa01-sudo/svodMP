from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path

from googleapiclient.errors import HttpError

from excel_reader import ExcelReadError, read_excel
from sheets_client import (
    apply_green_fill,
    build_sheets_service,
    fetch_sheet_infos,
    find_mp_sheet,
    get_last_filled_row,
    insert_row,
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
    "Козловская": ["козловская"],
    "Диамант": ["диамант", "цитрус"],
    "Привоз": ["привоз"],
    "Бахтурова": ["бахтурова"],
    "Ахтубинск": ["ахтубинск"],
    "СтройГрад": ["стройград", "строй град", "стройград"],
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
        except HttpError as exc:
            logger.error("Ошибка доступа к Google Sheets API: %s", exc)
            return

    for file_path in files:
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

        if dry_run:
            logger.info("[DRY RUN] Перенесли бы %s строк", len(excel_data.rows))
            continue

        sheet_info = find_mp_sheet(sheet_infos, context.store)
        if not sheet_info:
            logger.error("%s: не найден лист МП для магазина '%s'", file_path.name, context.store)
            continue

        last_row = get_last_filled_row(service, spreadsheet_id, sheet_info.title)
        summary_row = last_row + 1
        data_start = summary_row + 1
        data_end = summary_row + len(excel_data.rows)

        logger.info(
            "Запись в лист '%s': строки %s-%s",
            sheet_info.title,
            data_start,
            data_end,
        )

        insert_row(service, spreadsheet_id, sheet_info.sheet_id, summary_row)
        apply_green_fill(service, spreadsheet_id, sheet_info.sheet_id, summary_row)
        period_label = context.period.split()[0]
        update_summary_row(
            service,
            spreadsheet_id,
            sheet_info.title,
            summary_row,
            period_label,
            data_start,
            data_end,
        )
        update_values(service, spreadsheet_id, sheet_info.title, data_start, excel_data.rows)
        update_formulas(service, spreadsheet_id, sheet_info.title, data_start, data_end)

        logger.info("%s: успешно перенесено строк: %s", file_path.name, len(excel_data.rows))


def _build_context(file_path: Path, fallback_period: str | None, dry_run: bool) -> FileContext:
    store = _detect_store(file_path.stem)
    if not store:
        raise ValueError("Не удалось определить магазин по названию")

    detected_period = _detect_period(file_path.stem)
    period = detected_period or fallback_period
    if not period:
        raise ValueError("Не найден период в названии и не указан --period")

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
