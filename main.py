from __future__ import annotations

import argparse
import logging

from config_loader import extract_spreadsheet_id, load_config
from processor import process_directory


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Импорт Excel данных в Google Sheets")
    parser.add_argument("--input_dir", required=True, help="Папка с Excel файлами")
    parser.add_argument("--period", help="Период в формате 'Месяц ГГГГ'")
    parser.add_argument("--spreadsheet_id", help="ID или ссылка Google таблицы")
    parser.add_argument("--credentials", required=True, help="Путь к service account JSON")
    parser.add_argument("--dry_run", action="store_true", help="Только логирование, без записи")
    parser.add_argument("--config", default="./config.json", help="Путь к config.json")
    return parser


def main() -> None:
    setup_logging()
    parser = build_parser()
    args = parser.parse_args()

    config = load_config(args.config)
    spreadsheet_id = extract_spreadsheet_id(args.spreadsheet_id) or extract_spreadsheet_id(
        config.get("spreadsheet_id")
    )

    if not spreadsheet_id:
        raise SystemExit("Не указан spreadsheet_id и он не найден в config.json")

    process_directory(
        input_dir=args.input_dir,
        period=args.period,
        spreadsheet_id=spreadsheet_id,
        credentials=args.credentials,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
