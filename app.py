from __future__ import annotations

import datetime
import io
import json
import logging
import shutil
import zipfile
from pathlib import Path
from typing import List

import streamlit as st
import streamlit.components.v1 as components

from config_loader import load_config
from excel_reader import read_excel_smart

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


def _convert_xls_to_xlsx(source_path: Path, target_path: Path) -> None:
    """Конвертирует .xls в .xlsx через pandas."""
    dataframe = read_excel_smart(source_path)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    dataframe.to_excel(target_path, index=False, engine="openpyxl")


def _prepare_xlsx_folder(source_dir: Path, xlsx_dir: Path) -> list[Path]:
    """Конвертирует .xls и копирует .xlsx в папку xlsx."""
    xlsx_dir.mkdir(parents=True, exist_ok=True)
    converted_files: list[Path] = []
    for file_path in source_dir.glob("*"):
        if file_path.suffix.lower() == ".xls":
            target_path = xlsx_dir / f"{file_path.stem}.xlsx"
            _convert_xls_to_xlsx(file_path, target_path)
            converted_files.append(target_path)
        elif file_path.suffix.lower() == ".xlsx":
            target_path = xlsx_dir / file_path.name
            shutil.copy(file_path, target_path)
            converted_files.append(target_path)
    return converted_files


def _build_xlsx_zip(files: list[Path]) -> bytes:
    """Собирает zip-архив из конвертированных xlsx файлов."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file_path in files:
            archive.write(file_path, arcname=file_path.name)
    buffer.seek(0)
    return buffer.getvalue()


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


def _resolve_credentials_path(temp_path: Path) -> str | None:
    """Возвращает путь к credentials (из Streamlit Secrets)."""
    secrets = st.secrets

    # Вариант 1: строковый JSON целиком
    if "credentials_json" in secrets:
        value = secrets["credentials_json"]
        if isinstance(value, str):
            credentials_file = temp_path / "credentials.json"
            credentials_file.write_text(value, encoding="utf-8")
            return str(credentials_file)

    # Вариант 2: строка с JSON в другом ключе
    if "credentials" in secrets:
        value = secrets["credentials"]
        if isinstance(value, str):
            credentials_file = temp_path / "credentials.json"
            credentials_file.write_text(value, encoding="utf-8")
            return str(credentials_file)

    # Вариант 3: секция-объект (как у тебя [google])
    for key in ("google", "gcp_service_account", "SVODMP"):
        if key in secrets:
            obj = secrets[key]  # это Secrets / mapping-подобный объект
            credentials_file = temp_path / "credentials.json"
            # превращаем его в обычный JSON
            credentials_file.write_text(
                json.dumps(dict(obj), ensure_ascii=False),
                encoding="utf-8",
            )
            return str(credentials_file)

    st.error("Не найдены credentials в Streamlit Secrets.")
    st.info(
        "Добавьте JSON service account в Secrets (секция [google]/[gcp_service_account]/[SVODMP] "
        "или ключ credentials_json / credentials)."
    )
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
        payload = json.loads(raw_content)
        private_key = payload.get("private_key")
        if isinstance(private_key, str):
            if "BEGIN PRIVATE KEY" not in private_key or "END PRIVATE KEY" not in private_key:
                st.error(
                    "В credentials отсутствует корректный private_key в формате PEM. "
                    "Проверьте, что ключ из service account не обрезан."
                )
                return False
        else:
            st.error("В credentials нет поля private_key или оно некорректного типа.")
            return False
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
    dry_run = False

    if uploaded_files:
        st.markdown("**Загруженные файлы:**")
        st.write([uploaded_file.name for uploaded_file in uploaded_files])

    current_files = [uploaded_file.name for uploaded_file in uploaded_files or []]
    if st.session_state.get("uploaded_files") != current_files:
        st.session_state["uploaded_files"] = current_files
        st.session_state["conversion_done"] = False
        st.session_state["xlsx_dir"] = None
        st.session_state["zip_name"] = None

    if st.button("Конвертировать в xlsx"):
        if not uploaded_files:
            st.error("Выберите файлы Excel для конвертации")
            return
        upload_dir = Path("./uploads")
        _save_uploaded_files(uploaded_files, upload_dir)
        xlsx_dir = upload_dir / "xlsx"
        converted_files = _prepare_xlsx_folder(upload_dir, xlsx_dir)
        st.session_state["conversion_done"] = True
        st.session_state["xlsx_dir"] = str(xlsx_dir)
        st.session_state["zip_name"] = None
        st.success(f"Готово. Файлы сохранены в {xlsx_dir}")
        if converted_files:
            st.write([path.name for path in converted_files])
            zip_buffer = _build_xlsx_zip(converted_files)
            st.download_button(
                "Скачать xlsx",
                data=zip_buffer,
                file_name="converted_xlsx.zip",
                mime="application/zip",
            )

    zip_upload = st.file_uploader(
        "Загрузить ZIP с xlsx для импорта",
        type=["zip"],
        accept_multiple_files=False,
    )
    if zip_upload and st.session_state.get("zip_name") != zip_upload.name:
        st.session_state["zip_name"] = zip_upload.name
        st.session_state["conversion_done"] = False
        st.session_state["xlsx_dir"] = None

    if st.button("Подготовить ZIP для импорта"):
        if not zip_upload:
            st.error("Загрузите ZIP с xlsx")
            return
        upload_dir = Path("./uploads")
        upload_dir.mkdir(parents=True, exist_ok=True)
        extracted_dir = upload_dir / "xlsx_from_zip"
        if extracted_dir.exists():
            shutil.rmtree(extracted_dir)
        extracted_dir.mkdir(parents=True, exist_ok=True)
        zip_bytes = zip_upload.getvalue()
        with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as archive:
            archive.extractall(extracted_dir)
        st.session_state["conversion_done"] = True
        st.session_state["xlsx_dir"] = str(extracted_dir)
        st.success(f"ZIP подготовлен. Файлы извлечены в {extracted_dir}")

    if st.button("Запустить импорт", disabled=not st.session_state.get("conversion_done")):
        st.session_state["log_lines"].clear()

        if not uploaded_files:
            st.error("Выберите файлы Excel")
            return
        if not spreadsheet_id:
            st.error("Не найден Spreadsheet ID в config.json")
            return
        xlsx_dir = st.session_state.get("xlsx_dir")
        if not xlsx_dir:
            st.error("Сначала выполните конвертацию в xlsx")
            return

        st.info("Запуск обработки. Логи смотрите ниже, в журнале выполнения.")
        progress_bar = st.progress(0)
        progress_text = st.empty()
        upload_dir = Path(xlsx_dir)

        credentials_to_use = _resolve_credentials_path(upload_dir)
        if not credentials_to_use:
            return
        if not _validate_credentials_json(credentials_to_use):
            return

        def _update_progress(current: int, total: int, filename: str) -> None:
            progress = int((current / total) * 100) if total else 100
            progress_bar.progress(progress)
            progress_text.info(f"Обрабатывается файл {current}/{total}: {filename}")

        process_directory(
            input_dir=str(upload_dir),
            period=period_value,
            spreadsheet_id=spreadsheet_id,
            credentials=credentials_to_use,
            dry_run=dry_run,
            progress_callback=_update_progress,
        )

        progress_bar.progress(100)
        progress_text.success("Обработка завершена.")

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
