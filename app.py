from __future__ import annotations

import datetime
import logging
import tempfile
from pathlib import Path

import streamlit as st

from config_loader import extract_spreadsheet_id, load_config
from processor import process_directory


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


def _render_period_picker() -> str:
    """Возвращает выбранный период в формате 'Месяц ГГГГ'."""
    months = [
        "Январь",
        "Февраль",
        "Март",
        "Апрель",
        "Май",
        "Июнь",
        "Июль",
        "Август",
        "Сентябрь",
        "Октябрь",
        "Ноябрь",
        "Декабрь",
    ]
    today = datetime.date.today()
    first_day = today.replace(day=1)
    previous_month = first_day - datetime.timedelta(days=1)
    default_month = previous_month.strftime("%B").capitalize()
    default_year = previous_month.year

    # Если локаль не настроена, fallback на индекс
    if default_month not in months:
        default_month = months[previous_month.month - 1]

    month = st.selectbox("Месяц периода", months, index=months.index(default_month))
    year = st.number_input("Год периода", min_value=2000, max_value=2100, value=default_year, step=1)
    return f"{month} {int(year)}"


def _save_uploaded_files(uploaded_files: list, target_dir: Path) -> None:
    """Сохраняет загруженные файлы во временную папку."""
    target_dir.mkdir(parents=True, exist_ok=True)
    for uploaded_file in uploaded_files:
        file_path = target_dir / uploaded_file.name
        file_path.write_bytes(uploaded_file.getbuffer())


def main() -> None:
    setup_logging()
    st.title("Импорт данных из Excel в Google Sheets")

    st.markdown("""
    **Описание**
    - Выберите Excel файлы (.xls/.xlsx)
    - Укажите период, если его нет в названии файла
    - При необходимости включите режим dry-run
    """)

    uploaded_files = st.file_uploader(
        "Excel файлы",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )
    period = _render_period_picker()
    config = load_config("./config.json")
    spreadsheet_id = extract_spreadsheet_id(config.get("spreadsheet_id", ""))
    credentials_path = config.get("credentials_path", "")
    dry_run = st.checkbox("Dry run (без записи)", value=True)

    if st.button("Запустить импорт"):
        if not uploaded_files or not spreadsheet_id or not credentials_path:
            st.error("Заполните обязательные поля: файлы и config.json")
            return

        st.info("Запуск обработки. Логи смотрите в консоли приложения.")
        with tempfile.TemporaryDirectory() as temp_dir:
            _save_uploaded_files(uploaded_files, Path(temp_dir))
            process_directory(
                input_dir=temp_dir,
                period=period,
                spreadsheet_id=spreadsheet_id,
                credentials=credentials_path,
                dry_run=dry_run,
            )
        st.success("Готово")


if __name__ == "__main__":
    main()
