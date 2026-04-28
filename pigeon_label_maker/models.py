from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


LAYER_MODES = ("Text", "Barcode", "QR", "Off")
ALIGNMENTS = ("Left", "Center", "Right")
LINE_MODES = ("Auto Wrap", "1 Line", "2 Lines")
OUTPUT_MODES = ("Printer", "BLE")


@dataclass
class LayerSettings:
    text: str = ""
    mode: str = "Text"
    align: str = "Left"


@dataclass
class AppSettings:
    font_name: str = "Arial Bold"
    line_mode: str = "Auto Wrap"
    layer1: LayerSettings = field(default_factory=LayerSettings)
    layer2: LayerSettings = field(
        default_factory=lambda: LayerSettings(mode="Off", align="Right")
    )
    profile_id: str = "tspl_small"
    label_width_mm: float = 40.0
    label_height_mm: float = 14.0
    gap_mm: float = 5.0
    render_dpi: int = 300
    print_dpi: int = 203
    baud_rate: int = 115200
    port: str = ""
    copies: int = 1
    density: int = 10
    print_offset_x_mm: float = 0.0
    print_offset_y_mm: float = 0.0
    print_scale: float = 1.0
    brightness: float = 1.0
    contrast: float = 2.0
    threshold: int = 180
    image_mode: str = "threshold"
    invert: bool = False
    auto_image: bool = True
    edge_enhance: bool = False
    output_mode: str = "Printer"
    ble_device_name: str = ""
    ble_device_address: str = ""
    ble_write_char_uuid: str = ""
    ble_pair: bool = False
    ble_write_with_response: bool = False
    ble_scan_timeout: float = 5.0
    ble_chunk_size: int = 180
    last_image_path: str = ""
    export_dir: str = ""
    command_output_dir: str = ""
    canvas_layout: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, raw: dict[str, Any] | None) -> "AppSettings":
        if not isinstance(raw, dict):
            return cls()

        settings = cls()
        for key, value in raw.items():
            if key in {"layer1", "layer2"} and isinstance(value, dict):
                current_layer = getattr(settings, key)
                layer = LayerSettings(
                    text=value.get("text", current_layer.text),
                    mode=value.get("mode", current_layer.mode),
                    align=value.get("align", current_layer.align),
                )
                setattr(settings, key, layer)
            elif hasattr(settings, key):
                setattr(settings, key, value)
        return settings
