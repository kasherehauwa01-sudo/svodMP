from __future__ import annotations

import argparse
import logging

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
    parser.add_argument("--spreadsheet_id", required=True, help="ID Google таблицы")
    parser.add_argument("--credentials", required=True, help="Путь к service account JSON")
    parser.add_argument("--dry_run", action="store_true", help="Только логирование, без записи")
    return parser


def main() -> None:
    setup_logging()
    parser = build_parser()
    args = parser.parse_args()

    process_directory(
        input_dir=args.input_dir,
        period=args.period,
        spreadsheet_id=args.spreadsheet_id,
        credentials=args.credentials,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
