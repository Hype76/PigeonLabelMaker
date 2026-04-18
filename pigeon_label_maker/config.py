from __future__ import annotations

import json
import logging
import os
from pathlib import Path

from .models import AppSettings


APP_NAME = "PigeonLabelMaker"
APP_DIR = Path(os.getenv("LOCALAPPDATA", Path.home() / ".pigeon_label_maker")) / APP_NAME
LOG_DIR = APP_DIR / "logs"
SETTINGS_FILE = APP_DIR / "settings.json"
USER_PRESETS_FILE = APP_DIR / "user_presets.json"
LOG_FILE = LOG_DIR / "app.log"


def ensure_app_dirs() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)


def load_json(path: Path, default: object) -> object:
    if not path.exists():
        return default

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return default


def save_json(path: Path, payload: object) -> None:
    ensure_app_dirs()
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_settings() -> AppSettings:
    return AppSettings.from_dict(load_json(SETTINGS_FILE, {}))


def save_settings(settings: AppSettings) -> None:
    save_json(SETTINGS_FILE, settings.to_dict())


def load_user_presets() -> dict[str, dict]:
    presets = load_json(USER_PRESETS_FILE, {})
    return presets if isinstance(presets, dict) else {}


def save_user_presets(presets: dict[str, dict]) -> None:
    save_json(USER_PRESETS_FILE, presets)


def setup_logging() -> logging.Logger:
    ensure_app_dirs()
    logger = logging.getLogger(APP_NAME)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    logger.propagate = False
    return logger
