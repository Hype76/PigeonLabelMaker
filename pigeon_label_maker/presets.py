from __future__ import annotations

from copy import deepcopy

from .models import AppSettings


DEFAULT_PRESETS: dict[str, dict] = {
    "Dual Text": {
        "font_name": "Arial Bold",
        "line_mode": "Auto Wrap",
        "layer1": {"text": "LEFT", "mode": "Text", "align": "Left"},
        "layer2": {"text": "RIGHT", "mode": "Text", "align": "Right"},
    },
    "Text + Barcode": {
        "font_name": "Arial Bold",
        "line_mode": "Auto Wrap",
        "layer1": {"text": "PIGEON 101", "mode": "Text", "align": "Left"},
        "layer2": {"text": "PIGEON-101", "mode": "Barcode", "align": "Right"},
    },
    "Text + QR": {
        "font_name": "Arial Bold",
        "line_mode": "2 Lines",
        "layer1": {"text": "Scan Me", "mode": "Text", "align": "Left"},
        "layer2": {"text": "https://example.com", "mode": "QR", "align": "Right"},
    },
    "Logo + Text": {
        "font_name": "Calibri",
        "line_mode": "Auto Wrap",
        "layer1": {"text": "Club Loft", "mode": "Text", "align": "Center"},
        "layer2": {"text": "", "mode": "Off", "align": "Right"},
    },
}

PRESET_KEYS = {
    "font_name",
    "line_mode",
    "layer1",
    "layer2",
    "profile_id",
    "label_width_mm",
    "label_height_mm",
    "gap_mm",
    "render_dpi",
    "print_dpi",
    "density",
    "contrast",
    "threshold",
    "invert",
}


def get_builtin_presets() -> dict[str, dict]:
    return deepcopy(DEFAULT_PRESETS)


def serialize_preset(settings: AppSettings) -> dict:
    payload = settings.to_dict()
    return {key: deepcopy(value) for key, value in payload.items() if key in PRESET_KEYS}
