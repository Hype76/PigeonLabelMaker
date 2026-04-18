from __future__ import annotations

import base64
from dataclasses import asdict, is_dataclass
from io import BytesIO
import json
from pathlib import Path
import sys
import traceback
from typing import Any

from PIL import Image, ImageDraw

from .config import load_settings, load_user_presets, save_settings
from .models import ALIGNMENTS, LAYER_MODES, LINE_MODES, OUTPUT_MODES, AppSettings
from .presets import get_builtin_presets
from .printing import (
    PROFILE_TEMPLATES,
    apply_print_processing,
    ble_connection_state,
    connect_ble_async,
    disconnect_ble,
    discover_ble_devices,
    list_ble_notify_characteristics,
    list_ble_writable_characteristics,
    list_serial_ports,
    run_ble_coro_in_worker,
    send_to_ble_printer,
    send_to_printer,
    validate_settings,
)
from .rendering import list_font_names, render_label


def image_to_data_url(image: Image.Image, fmt: str = "PNG") -> str:
    buffer = BytesIO()
    image.save(buffer, format=fmt)
    payload = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/{fmt.lower()};base64,{payload}"


def load_optional_image(image_path: str | None) -> Image.Image | None:
    if not image_path:
        return None

    if str(image_path).startswith("data:image"):
        try:
            _header, encoded = str(image_path).split(",", 1)
            data = base64.b64decode(encoded)
            return Image.open(BytesIO(data)).convert("RGB")
        except Exception as exc:
            raise ValueError("Could not decode image data") from exc

    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(f"Image file not found: {path}")
    with Image.open(path) as image:
        return image.copy()


def serialize_payload(value: Any) -> Any:
    if isinstance(value, AppSettings):
        return value.to_dict()
    if is_dataclass(value):
        return asdict(value)
    if isinstance(value, tuple):
        return [serialize_payload(item) for item in value]
    if isinstance(value, list):
        return [serialize_payload(item) for item in value]
    if isinstance(value, dict):
        return {key: serialize_payload(item) for key, item in value.items()}
    return value


def build_settings(raw: dict[str, Any] | None) -> AppSettings:
    return AppSettings.from_dict(raw or {})


def command_init(_params: dict[str, Any]) -> dict[str, Any]:
    settings = load_settings()
    return {
        "settings": settings.to_dict(),
        "fonts": list_font_names(),
        "profiles": [asdict(profile) for profile in PROFILE_TEMPLATES.values()],
        "builtInPresets": get_builtin_presets(),
        "userPresets": load_user_presets(),
        "lineModes": list(LINE_MODES),
        "layerModes": list(LAYER_MODES),
        "alignments": list(ALIGNMENTS),
        "outputModes": list(OUTPUT_MODES),
    }


def command_save_settings(params: dict[str, Any]) -> dict[str, Any]:
    settings = build_settings(params.get("settings"))
    save_settings(settings)
    return {"saved": True}


def command_preview(params: dict[str, Any]) -> dict[str, Any]:
    settings = build_settings(params.get("settings"))
    current_image = load_optional_image(params.get("imagePath"))
    if current_image is not None:
        rendered = current_image
    else:
        rendered = render_label(settings, current_image)
    design = rendered
    print_ready = apply_print_processing(rendered, settings)
    return {
        "designImage": image_to_data_url(design),
        "printImage": image_to_data_url(print_ready.convert("L")),
    }


def command_export_png(params: dict[str, Any]) -> dict[str, Any]:
    settings = build_settings(params.get("settings"))
    current_image = load_optional_image(params.get("imagePath"))
    output_path = Path(params["outputPath"])
    rendered = render_label(settings, current_image)
    rendered.save(output_path, format="PNG")
    return {"path": str(output_path)}


def command_list_serial_ports(_params: dict[str, Any]) -> dict[str, Any]:
    return {"ports": list_serial_ports()}


def command_scan_ble(params: dict[str, Any]) -> dict[str, Any]:
    timeout = float(params.get("timeout", 5.0))
    return {"devices": serialize_payload(discover_ble_devices(timeout=timeout))}


def command_connect_ble(params: dict[str, Any]) -> dict[str, Any]:
    address = str(params.get("address", "")).strip()
    pair = bool(params.get("pair", False))
    state = run_ble_coro_in_worker(connect_ble_async(address=address, pair=pair))
    return {"state": serialize_payload(state)}


def command_disconnect_ble(_params: dict[str, Any]) -> dict[str, Any]:
    disconnect_ble()
    return {"state": serialize_payload(ble_connection_state())}


def command_ble_state(_params: dict[str, Any]) -> dict[str, Any]:
    return {"state": serialize_payload(ble_connection_state())}


async def get_ble_battery_async(address: str, pair: bool = False) -> int | None:
    try:
        from .printing import ensure_ble_connection_async

        client = await ensure_ble_connection_async(address=address, pair=pair)
        services = client.services

        for service in services:
            if "180f" in str(service.uuid).lower():
                for characteristic in service.characteristics:
                    if "2a19" in str(characteristic.uuid).lower():
                        data = await client.read_gatt_char(characteristic.uuid)
                        if data:
                            return int(data[0])
    except Exception:
        return None

    return None


def command_ble_battery(params: dict[str, Any]) -> dict[str, Any]:
    address = str(params.get("address", "")).strip()
    pair = bool(params.get("pair", False))

    if not address:
        return {"battery": None}

    level = run_ble_coro_in_worker(get_ble_battery_async(address, pair))
    return {"battery": level}


def command_inspect_ble(params: dict[str, Any]) -> dict[str, Any]:
    address = str(params.get("address", "")).strip()
    pair = bool(params.get("pair", False))
    return {
        "writable": serialize_payload(list_ble_writable_characteristics(address=address, pair=pair)),
        "notify": serialize_payload(list_ble_notify_characteristics(address=address, pair=pair)),
    }


def command_print(params: dict[str, Any]) -> dict[str, Any]:
    settings = build_settings(params.get("settings"))
    errors = validate_settings(settings)
    if errors:
        raise ValueError("\n".join(errors))

    current_image = load_optional_image(params.get("imagePath"))
    if current_image is not None:
        rendered = current_image
    else:
        rendered = render_label(settings, current_image)
    print_ready = apply_print_processing(rendered, settings)

    from .printing import build_print_command

    command = build_print_command(print_ready, settings)
    if settings.output_mode == "BLE":
        result = send_to_ble_printer(command, settings, rendered)
        save_settings(settings)
        return {
            "mode": "BLE",
            "result": serialize_payload(result),
        }

    send_to_printer(command, settings)
    save_settings(settings)
    return {
        "mode": "Printer",
        "port": settings.port,
        "copies": settings.copies,
    }


def command_test_print(params: dict[str, Any]) -> dict[str, Any]:
    settings = build_settings(params.get("settings"))
    errors = validate_settings(settings)
    if errors:
        raise ValueError("\n".join(errors))

    rendered = Image.new("RGB", (384, 120), "white")
    draw = ImageDraw.Draw(rendered)
    draw.text((12, 42), "TEST PRINT", fill="black")

    print_ready = apply_print_processing(rendered, settings)

    from .printing import build_print_command

    command = build_print_command(print_ready, settings)
    if settings.output_mode == "BLE":
        result = send_to_ble_printer(command, settings, rendered)
        return {
            "mode": "BLE",
            "result": serialize_payload(result),
        }

    send_to_printer(command, settings)
    return {
        "mode": "Printer",
        "port": settings.port,
    }


COMMANDS = {
    "init": command_init,
    "saveSettings": command_save_settings,
    "preview": command_preview,
    "exportPng": command_export_png,
    "listSerialPorts": command_list_serial_ports,
    "scanBle": command_scan_ble,
    "connectBle": command_connect_ble,
    "disconnectBle": command_disconnect_ble,
    "bleState": command_ble_state,
    "bleBattery": command_ble_battery,
    "inspectBle": command_inspect_ble,
    "testPrint": command_test_print,
    "print": command_print,
}


def write_message(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload) + "\n")
    sys.stdout.flush()


def main() -> None:
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue
        request_id: Any = None
        try:
            request = json.loads(line)
            request_id = request.get("id")
            command_name = request.get("command")
            params = request.get("params") or {}
            if command_name not in COMMANDS:
                raise ValueError(f"Unknown command: {command_name}")
            result = COMMANDS[command_name](params)
            write_message({"id": request_id, "ok": True, "result": serialize_payload(result)})
        except Exception as exc:
            write_message(
                {
                    "id": request_id,
                    "ok": False,
                    "error": str(exc),
                    "traceback": traceback.format_exc(),
                }
            )


if __name__ == "__main__":
    main()
