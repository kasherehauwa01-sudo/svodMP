from __future__ import annotations

import datetime
import json
import logging
import tempfile
from pathlib import Path
from typing import List

import streamlit as st

from config_loader import extract_spreadsheet_id, load_config
from processor import process_directory


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


class StreamlitLogHandler(logging.Handler):
    """Логгер, который сохраняет сообщения в session_state для вывода в UI."""

    def __init__(self, log_store: List[str]) -> None:
        super().__init__()
        self.log_store = log_store

    def emit(self, record: logging.LogRecord) -> None:
        message = self.format(record)
        self.log_store.append(message)


def setup_streamlit_logger() -> None:
    """Инициализирует логирование в UI один раз за сессию."""
    if "log_lines" not in st.session_state:
        st.session_state["log_lines"] = []

    if st.session_state.get("log_handler_attached"):
        return

    handler = StreamlitLogHandler(st.session_state["log_lines"])
    handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
    logging.getLogger().addHandler(handler)
    st.session_state["log_handler_attached"] = True


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

    default_month = months[previous_month.month - 1]
    default_year = previous_month.year

    month = st.selectbox("Месяц периода", months, index=months.index(default_month))
    year = st.number_input(
        "Год периода",
        min_value=2000,
        max_value=2100,
        value=default_year,
        step=1,
    )
    return f"{month} {int(year)}"


def _save_uploaded_files(uploaded_files: list, target_dir: Path) -> None:
    """Сохраняет загруженные файлы во временную папку."""
    target_dir.mkdir(parents=True, exist_ok=True)
    for uploaded_file in uploaded_files:
        file_path = target_dir / uploaded_file.name
        file_path.write_bytes(uploaded_file.getbuffer())


def _resolve_credentials_path(
    temp_path: Path,
    credentials_path: str,
    credentials_upload,
) -> str | None:
    """
    Возвращает путь к credentials:
    1) Streamlit Secrets: [gcp_service_account] (рекомендуется)
    2) Secrets: credentials_json или [google] (на случай старых настроек)
    3) Загрузка JSON через UI
    4) Локальный путь к файлу (для локального запуска)
    """
    secrets = st.secrets

    # Рекомендуемый вариант: secrets.toml содержит таблицу [gcp_service_account]
    if "gcp_service_account" in secrets:
        credentials_file = temp_path / "credentials.json"
        credentials_file.write_text(
            json.dumps(dict(secrets["gcp_service_account"]), ensure_ascii=False),
            encoding="utf-8",
        )
        return str(credentials_file)

    # Backward-compatible варианты
    if "credentials_json" in secrets:
        credentials_file = temp_path / "credentials.json"
        credentials_file.write_text(secrets["credentials_json"], encoding="utf-8")
        return str(credentials_file)

    if "google" in secrets and isinstance(secrets["google"], dict):
        credentials_file = temp_path / "credentials.json"
        credentials_file.write_text(
            json.dumps(secrets["google"], ensure_ascii=False),
            encoding="utf-8",
        )
        return str(credentials_file)

    # Загрузка из UI
    if credentials_upload:
        credentials_file = temp_path / credentials_upload.name
        credentials_file.write_bytes(credentials_upload.getbuffer())
        return str(credentials_file)

    # Локальный путь
    if not credentials_path:
        st.error("Укажите credentials в Secrets, загрузите JSON или введите путь")
        return None

    if not Path(credentials_path).exists():
        st.error("Файл credentials не найден. Проверьте путь или config.json")
        return None

    return credentials_path


def main() -> None:
    setup_logging()
    setup_streamlit_logger()

    st.title("Импорт данных из Excel в Google Sheets")

    st.markdown(
        """
**Описание**
- Выберите Excel файлы (.xls/.xlsx)
- Укажите период, если его нет в названии файла (можно отключить)
- Укажите Spreadsheet ID (или оставьте пустым, если он есть в config.json)
- Credentials можно задать через Streamlit Secrets (рекомендуется), загрузкой JSON или путём к файлу
- При необходимости включите режим dry-run
"""
    )

    uploaded_files = st.file_uploader(
        "Excel файлы",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )

    use_manual_period = st.checkbox("Указать период вручную", value=True)
    period_value = _render_period_picker() if use_manual_period else None

    config = load_config("./config.json")
    config_sheet = extract_spreadsheet_id(config.get("spreadsheet_id"))
    config_credentials = config.get("credentials_path") or config.get("credentials") or ""

    spreadsheet_input = st.text_input(
        "Spreadsheet ID или ссылка (если пусто — возьмём из config.json)",
        value=config_sheet or "",
    )
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_input) or config_sheet

    credentials_path = st.text_input(
        "Путь к service account JSON (для локального запуска)",
        value=config_credentials,
    )

    credentials_upload = st.file_uploader(
        "Или загрузите service account JSON (если не используете Secrets)",
        type=["json"],
        accept_multiple_files=False,
    )

    dry_run = st.checkbox("Dry run (без записи)", value=True)

    if st.button("Запустить импорт"):
        st.session_state["log_lines"].clear()

        if not uploaded_files:
            st.error("Выберите файлы Excel")
            return

        if not spreadsheet_id:
            st.error("Не найден Spreadsheet ID. Укажите его или заполните config.json")
            return

        st.info("Запуск обработки. Логи смотрите ниже.")

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            _save_uploaded_files(uploaded_files, temp_path)

            credentials_to_use = _resolve_credentials_path(
                temp_path=temp_path,
                credentials_path=credentials_path,
                credentials_upload=credentials_upload,
            )
            if not credentials_to_use:
                return

            process_directory(
                input_dir=temp_dir,
                period=period_value,
                spreadsheet_id=spreadsheet_id,
                credentials=credentials_to_use,
                dry_run=dry_run,
            )

        st.success("Готово")

    st.subheader("Логи")
    log_text = "\n".join(st.session_state.get("log_lines", []))
    st.text_area("Журнал выполнения", value=log_text, height=300)
    st.download_button(
        "Скачать логи",
        data=log_text,
        file_name="logs.txt",
        mime="text/plain",
    )


if __name__ == "__main__":
    main()
