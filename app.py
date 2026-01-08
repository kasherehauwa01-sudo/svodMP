from __future__ import annotations

import datetime
import json
import logging
import os
from pathlib import Path
from typing import List

import streamlit as st
import streamlit.components.v1 as components

from config_loader import load_config

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

    # Не зависим от локали
    default_month = months[previous_month.month - 1]
    default_year = previous_month.year

    month = st.selectbox("Месяц периода", months, index=months.index(default_month))
    year = st.number_input("Год периода", min_value=2000, max_value=2100, value=default_year, step=1)
    return f"{month} {int(year)}"


def _save_uploaded_files(uploaded_files: list, target_dir: Path) -> None:
    """Сохраняет загруженные файлы в указанную папку."""
    target_dir.mkdir(parents=True, exist_ok=True)
    for uploaded_file in uploaded_files:
        file_path = target_dir / uploaded_file.name
        file_path.write_bytes(uploaded_file.getbuffer())


def _copy_to_clipboard(text: str) -> None:
    """Копирует текст в буфер обмена через компонент HTML."""
    escaped_text = (
        text.replace("\\", "\\\\")
        .replace("`", "\\`")
        .replace("$", "\\$")
        .replace("\n", "\\n")
    )
    html = (
        "<script>"
        f"navigator.clipboard.writeText(`{escaped_text}`);"
        "</script>"
    )
    components.html(html, height=0)


def _resolve_credentials_path(
    temp_path: Path,
    credentials_path: str,
    credentials_upload,
) -> str | None:
    """Возвращает путь к credentials (из GitHub Secrets, Streamlit secrets, загрузки или пути)."""
    github_secret = os.getenv("SVODMP")
    if github_secret:
        credentials_file = temp_path / "credentials.json"
        credentials_file.write_text(github_secret, encoding="utf-8")
        return str(credentials_file)

    secrets = st.secrets
    if "credentials_json" in secrets and isinstance(secrets["credentials_json"], str):
        credentials_file = temp_path / "credentials.json"
        credentials_file.write_text(secrets["credentials_json"], encoding="utf-8")
        return str(credentials_file)

    if "credentials" in secrets and isinstance(secrets["credentials"], str):
        credentials_file = temp_path / "credentials.json"
        credentials_file.write_text(secrets["credentials"], encoding="utf-8")
        return str(credentials_file)

    if "google" in secrets and isinstance(secrets["google"], dict):
        credentials_file = temp_path / "credentials.json"
        credentials_file.write_text(
            json.dumps(secrets["google"], ensure_ascii=False),
            encoding="utf-8",
        )
        return str(credentials_file)

    if "gcp_service_account" in secrets and isinstance(secrets["gcp_service_account"], dict):
        credentials_file = temp_path / "credentials.json"
        credentials_file.write_text(
            json.dumps(secrets["gcp_service_account"], ensure_ascii=False),
            encoding="utf-8",
        )
        return str(credentials_file)

    if credentials_upload:
        credentials_file = temp_path / credentials_upload.name
        credentials_file.write_bytes(credentials_upload.getbuffer())
        return str(credentials_file)

    if credentials_path and Path(credentials_path).exists():
        return credentials_path

    st.error("Не найдены credentials в GitHub Secrets (SVODMP), Secrets или по пути credentials_path из config.json.")
    return None


def _validate_credentials_json(credentials_path: str) -> bool:
    """Проверяет, что credentials файл содержит валидный JSON."""
    try:
        raw_content = Path(credentials_path).read_text(encoding="utf-8")
        stripped_content = raw_content.strip()
        if not stripped_content:
            st.error("Файл credentials пустой. Загрузите полный JSON service account.")
            return False
        if not stripped_content.startswith("{"):
            st.error("Файл credentials должен быть JSON-объектом. Проверьте формат файла.")
            return False
        json.loads(raw_content)
    except (OSError, json.JSONDecodeError) as exc:
        st.error(f"Некорректный JSON в credentials файле: {exc}")
        st.info("Проверьте, что вы загрузили JSON service account, а не пустой файл.")
        return False
    return True


def main() -> None:
    setup_logging()
    setup_streamlit_logger()

    st.title("Импорт данных из Excel в Google Sheets")

    st.markdown(
        """
**Описание**
- Выберите Excel файлы (.xls/.xlsx)
- Укажите период, если его нет в названии файла (можно отключить)
- При необходимости включите режим dry-run
"""
    )

    uploaded_files = st.file_uploader(
        "Excel файлы",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )

    use_manual_period = st.checkbox("Указать период вручную", value=True)
    period_value = None
    if use_manual_period:
        period_value = _render_period_picker()

    config = load_config("./config.json")
    spreadsheet_id = config.get("spreadsheet_id")
    credentials_path = config.get("credentials_path") or config.get("credentials") or ""

    credentials_upload = st.file_uploader(
        "Загрузите service account JSON (если не используете Secrets)",
        type=["json"],
        accept_multiple_files=False,
    )

    dry_run = st.checkbox("Dry run (без записи)", value=True)

    if uploaded_files:
        st.markdown("**Загруженные файлы:**")
        st.write([uploaded_file.name for uploaded_file in uploaded_files])

    if st.button("Запустить импорт"):
        st.session_state["log_lines"].clear()

        if not uploaded_files:
            st.error("Выберите файлы Excel")
            return
        if not spreadsheet_id:
            st.error("Не найден Spreadsheet ID в config.json")
            return

        st.info("Запуск обработки. Логи смотрите ниже, в журнале выполнения.")
        upload_dir = Path("./uploads")
        _save_uploaded_files(uploaded_files, upload_dir)

        credentials_to_use = _resolve_credentials_path(
            upload_dir,
            credentials_path=credentials_path,
            credentials_upload=credentials_upload,
        )
        if not credentials_to_use:
            return
        if not _validate_credentials_json(credentials_to_use):
            return

        process_directory(
            input_dir=str(upload_dir),
            period=period_value,
            spreadsheet_id=spreadsheet_id,
            credentials=credentials_to_use,
            dry_run=dry_run,
        )

        st.success("Готово")

    st.subheader("Логи")
    log_text = "\n".join(st.session_state.get("log_lines", []))
    st.text_area(
        "Журнал выполнения",
        value=log_text,
        height=300,
    )
    if st.button("Копировать логи"):
        _copy_to_clipboard(log_text)
        st.success("Логи скопированы в буфер обмена")


if __name__ == "__main__":
    main()
