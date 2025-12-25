from __future__ import annotations

import logging
from pathlib import Path

import streamlit as st

from processor import process_directory


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


def main() -> None:
    setup_logging()
    st.title("Импорт данных из Excel в Google Sheets")

    st.markdown("""
    **Описание**
    - Выберите папку с Excel файлами (.xls/.xlsx)
    - Укажите период, если его нет в названии файла
    - При необходимости включите режим dry-run
    """)

    input_dir = st.text_input("Папка с файлами", value="./input")
    period = st.text_input("Период (например, Декабрь 2025)")
    spreadsheet_id = st.text_input("Spreadsheet ID")
    credentials_path = st.text_input("Путь к service account JSON", value="./service_account.json")
    dry_run = st.checkbox("Dry run (без записи)", value=True)

    if st.button("Запустить импорт"):
        if not input_dir or not spreadsheet_id or not credentials_path:
            st.error("Заполните обязательные поля: папка, spreadsheet_id, credentials")
            return

        if period.strip() == "":
            period_value = None
        else:
            period_value = period.strip()

        if not Path(input_dir).exists():
            st.error("Папка не найдена")
            return

        st.info("Запуск обработки. Логи смотрите в консоли приложения.")
        process_directory(
            input_dir=input_dir,
            period=period_value,
            spreadsheet_id=spreadsheet_id,
            credentials=credentials_path,
            dry_run=dry_run,
        )
        st.success("Готово")


if __name__ == "__main__":
    main()
