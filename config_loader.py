from __future__ import annotations

import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def load_config(config_path: str) -> dict:
    path = Path(config_path)
    if not path.exists():
        logger.warning("Конфиг не найден: %s", config_path)
        return {}

    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def extract_spreadsheet_id(value: str | None) -> str | None:
    if not value:
        return None

    if "/spreadsheets/d/" in value:
        part = value.split("/spreadsheets/d/")[-1]
        return part.split("/")[0]

    if "/edit" in value:
        return value.split("/edit")[0]

    return value
