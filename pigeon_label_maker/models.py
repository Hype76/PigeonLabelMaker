from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


LAYER_MODES = ("Text", "Barcode", "QR", "Off")
ALIGNMENTS = ("Left", "Center", "Right")
LINE_MODES = ("Auto Wrap", "1 Line", "2 Lines")
OUTPUT_MODES = ("Printer", "Mock File")


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
    contrast: float = 2.0
    threshold: int = 180
    invert: bool = True
    output_mode: str = "Printer"
    last_image_path: str = ""
    export_dir: str = ""
    command_output_dir: str = ""

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
