from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageEnhance, ImageOps
import serial
import serial.tools.list_ports

from .models import AppSettings
from .rendering import RESAMPLE, generate_barcode, generate_qr, mm_to_pixels


@dataclass(frozen=True)
class PrinterProfile:
    identifier: str
    name: str
    width_mm: float
    height_mm: float
    gap_mm: float
    render_dpi: int = 300
    print_dpi: int = 203
    baud_rate: int = 115200


PROFILE_TEMPLATES = {
    "tspl_small": PrinterProfile(
        identifier="tspl_small",
        name="Generic TSPL 40x14 mm",
        width_mm=40.0,
        height_mm=14.0,
        gap_mm=5.0,
    ),
    "tspl_medium": PrinterProfile(
        identifier="tspl_medium",
        name="Generic TSPL 50x20 mm",
        width_mm=50.0,
        height_mm=20.0,
        gap_mm=5.0,
    ),
    "tspl_square": PrinterProfile(
        identifier="tspl_square",
        name="Generic TSPL 30x30 mm",
        width_mm=30.0,
        height_mm=30.0,
        gap_mm=3.0,
    ),
}


def profile_names() -> list[str]:
    return [profile.name for profile in PROFILE_TEMPLATES.values()]


def profile_by_name(name: str) -> PrinterProfile | None:
    for profile in PROFILE_TEMPLATES.values():
        if profile.name == name:
            return profile
    return None


def print_pixel_size(settings: AppSettings) -> tuple[int, int]:
    width_px = mm_to_pixels(settings.label_width_mm, settings.print_dpi)
    height_px = mm_to_pixels(settings.label_height_mm, settings.print_dpi)
    return height_px, width_px


def apply_print_processing(image: Image.Image, settings: AppSettings) -> Image.Image:
    grayscale = ImageOps.grayscale(image)
    contrasted = ImageEnhance.Contrast(grayscale).enhance(settings.contrast)
    if settings.invert:
        contrasted = ImageOps.invert(contrasted)

    rotated = contrasted.rotate(90, expand=True)
    rotated = rotated.resize(print_pixel_size(settings), RESAMPLE)
    thresholded = rotated.point(
        lambda pixel: 255 if pixel >= settings.threshold else 0,
        mode="1",
    )

    aligned_width = int(math.ceil(thresholded.width / 8) * 8)
    if aligned_width == thresholded.width:
        return thresholded

    padded = Image.new("1", (aligned_width, thresholded.height), 1)
    padded.paste(thresholded, (0, 0))
    return padded


def validate_settings(settings: AppSettings) -> list[str]:
    errors: list[str] = []
    if settings.label_width_mm <= 0 or settings.label_height_mm <= 0:
        errors.append("Label width and height must be greater than zero.")
    if settings.gap_mm < 0:
        errors.append("Label gap cannot be negative.")
    if settings.render_dpi <= 0 or settings.print_dpi <= 0:
        errors.append("Render DPI and print DPI must be greater than zero.")
    if settings.copies < 1 or settings.copies > 50:
        errors.append("Copies must be between 1 and 50.")
    if settings.density < 1 or settings.density > 15:
        errors.append("Density must be between 1 and 15.")
    if settings.threshold < 0 or settings.threshold > 255:
        errors.append("Threshold must be between 0 and 255.")
    if settings.output_mode == "Printer" and not settings.port.strip():
        errors.append("Select a printer port before printing.")

    for index, layer in enumerate((settings.layer1, settings.layer2), start=1):
        if layer.mode == "Off":
            continue
        if not layer.text.strip():
            errors.append(f"Layer {index} needs content for {layer.mode}.")
            continue
        if layer.mode == "Barcode":
            try:
                generate_barcode(layer.text)
            except Exception as exc:
                errors.append(f"Layer {index} barcode is invalid: {exc}")
        if layer.mode == "QR":
            try:
                generate_qr(layer.text)
            except Exception as exc:
                errors.append(f"Layer {index} QR content is invalid: {exc}")
    return errors


def build_print_command(binary_image: Image.Image, settings: AppSettings) -> bytes:
    width_bytes = binary_image.width // 8
    height_dots = binary_image.height
    payload = binary_image.tobytes()

    command = b"\x1b!o\r\n"
    command += f"SIZE {settings.label_width_mm:.1f} mm,{settings.label_height_mm:.1f} mm\r\n".encode()
    command += f"GAP {settings.gap_mm:.1f} mm,0 mm\r\n".encode()
    command += b"DIRECTION 1,1\r\n"
    command += f"DENSITY {settings.density}\r\n".encode()
    command += b"CLS\r\n"
    command += f"BITMAP 0,0,{width_bytes},{height_dots},1,".encode() + payload
    command += f"\r\nPRINT {settings.copies}\r\n".encode()
    return command


def list_serial_ports() -> list[str]:
    return [port.device for port in serial.tools.list_ports.comports()]


def send_to_printer(command: bytes, settings: AppSettings) -> None:
    with serial.Serial(settings.port, settings.baud_rate, timeout=2) as serial_port:
        serial_port.write(command)


def save_command_file(path: Path, command: bytes) -> None:
    path.write_bytes(command)
