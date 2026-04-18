from __future__ import annotations

import math
from pathlib import Path

import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageChops, ImageDraw, ImageFont, ImageOps
import qrcode

from .models import AppSettings, LayerSettings


RESAMPLE = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
FONT_CANDIDATES = {
    "Arial Bold": "arialbd.ttf",
    "Arial": "arial.ttf",
    "Calibri": "calibri.ttf",
    "Consolas": "consola.ttf",
}


def mm_to_pixels(mm: float, dpi: int) -> int:
    return max(1, int(round((mm / 25.4) * dpi)))


def label_pixel_size(settings: AppSettings, dpi: int | None = None) -> tuple[int, int]:
    target_dpi = dpi or settings.render_dpi
    return (
        mm_to_pixels(settings.label_width_mm, target_dpi),
        mm_to_pixels(settings.label_height_mm, target_dpi),
    )


def list_font_names() -> list[str]:
    return list(FONT_CANDIDATES.keys())


def resolve_font_path(font_name: str) -> str:
    candidate = FONT_CANDIDATES.get(font_name, font_name)
    if Path(candidate).exists():
        return candidate

    windows_font = Path("C:/Windows/Fonts") / candidate
    if windows_font.exists():
        return str(windows_font)
    return candidate


def get_font(font_name: str, size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    try:
        return ImageFont.truetype(resolve_font_path(font_name), size)
    except OSError:
        return ImageFont.load_default()


def flatten_to_rgb(image: Image.Image) -> Image.Image:
    normalized = ImageOps.exif_transpose(image).convert("RGBA")
    background = Image.new("RGBA", normalized.size, "white")
    composited = Image.alpha_composite(background, normalized)
    return composited.convert("RGB")


def crop_white_space(image: Image.Image) -> Image.Image:
    rgb = image.convert("RGB")
    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white"))
    bbox = diff.getbbox()
    return rgb.crop(bbox) if bbox else rgb


def generate_barcode(text: str) -> Image.Image:
    code128 = barcode.get_barcode_class("code128")
    image = code128(text, writer=ImageWriter()).render(
        writer_options={
            "module_width": 0.25,
            "module_height": 12.0,
            "quiet_zone": 1.0,
            "font_size": 0,
            "text_distance": 1,
            "dpi": 300,
        }
    )
    return crop_white_space(image)


def generate_qr(text: str) -> Image.Image:
    image = qrcode.make(text).convert("RGB")
    return crop_white_space(image)


def measure_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.ImageFont,
) -> tuple[int, int]:
    left, top, right, bottom = draw.textbbox((0, 0), text or " ", font=font)
    return right - left, bottom - top


def wrap_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.ImageFont,
    max_width: int,
    mode: str,
) -> list[str]:
    stripped = text.strip()
    if not stripped:
        return []

    if mode == "1 Line":
        return [stripped]

    words = stripped.split()
    if len(words) == 1:
        return [stripped]

    if mode == "2 Lines":
        best_lines: list[str] | None = None
        best_score: tuple[int, int] | None = None
        for index in range(1, len(words)):
            candidate = [" ".join(words[:index]), " ".join(words[index:])]
            widths = [measure_text(draw, line, font)[0] for line in candidate]
            score = (max(widths), abs(widths[0] - widths[1]))
            if best_score is None or score < best_score:
                best_score = score
                best_lines = candidate
        return best_lines or [stripped]

    lines: list[str] = []
    current_words: list[str] = []
    for word in words:
        attempt = " ".join(current_words + [word])
        width, _ = measure_text(draw, attempt, font)
        if current_words and width > max_width:
            lines.append(" ".join(current_words))
            current_words = [word]
        else:
            current_words.append(word)

    if current_words:
        lines.append(" ".join(current_words))
    return lines


def fit_text_to_box(
    draw: ImageDraw.ImageDraw,
    text: str,
    font_name: str,
    max_width: int,
    max_height: int,
    line_mode: str,
) -> tuple[ImageFont.ImageFont, list[str], int]:
    min_font_size = 8
    max_font_size = max(18, min(140, max_height))

    for size in range(max_font_size, min_font_size - 1, -1):
        font = get_font(font_name, size)
        lines = wrap_text(draw, text, font, max_width, line_mode)
        if not lines:
            continue

        line_height = max(1, int(math.ceil(measure_text(draw, "Ag", font)[1] * 1.1)))
        total_height = line_height * len(lines)
        widest = max(measure_text(draw, line, font)[0] for line in lines)
        if widest <= max_width and total_height <= max_height:
            return font, lines, line_height

    font = get_font(font_name, min_font_size)
    lines = wrap_text(draw, text, font, max_width, line_mode)
    line_height = max(1, int(math.ceil(measure_text(draw, "Ag", font)[1] * 1.1)))
    return font, lines or [text.strip()], line_height


def active_layers(settings: AppSettings) -> list[tuple[int, LayerSettings]]:
    result: list[tuple[int, LayerSettings]] = []
    for index, layer in enumerate((settings.layer1, settings.layer2), start=1):
        if layer.mode != "Off" and layer.text.strip():
            result.append((index, layer))
    return result


def layer_boxes(
    settings: AppSettings,
    width: int,
    height: int,
) -> dict[int, tuple[int, int, int, int]]:
    layers = active_layers(settings)
    margin = max(8, int(width * 0.02))
    gap = max(8, int(width * 0.02))

    if len(layers) == 1:
        return {layers[0][0]: (margin, margin, width - margin, height - margin)}

    if len(layers) >= 2:
        inner_width = width - (margin * 2) - gap
        left_width = inner_width // 2
        return {
            layers[0][0]: (margin, margin, margin + left_width, height - margin),
            layers[1][0]: (margin + left_width + gap, margin, width - margin, height - margin),
        }
    return {}


def paste_background_image(canvas: Image.Image, current_image: Image.Image) -> None:
    background = flatten_to_rgb(current_image)
    background.thumbnail(canvas.size, RESAMPLE)
    x_pos = (canvas.width - background.width) // 2
    y_pos = (canvas.height - background.height) // 2
    canvas.paste(background, (x_pos, y_pos))


def render_layer(
    base: Image.Image,
    settings: AppSettings,
    layer: LayerSettings,
    box: tuple[int, int, int, int],
) -> None:
    draw = ImageDraw.Draw(base)
    left, top, right, bottom = box
    box_width = max(1, right - left)
    box_height = max(1, bottom - top)

    if layer.mode == "Text":
        font, lines, line_height = fit_text_to_box(
            draw=draw,
            text=layer.text,
            font_name=settings.font_name,
            max_width=box_width,
            max_height=box_height,
            line_mode=settings.line_mode,
        )
        total_height = line_height * len(lines)
        y_pos = top + max(0, (box_height - total_height) // 2)
        for line in lines:
            width, _ = measure_text(draw, line, font)
            if layer.align == "Left":
                x_pos = left
            elif layer.align == "Right":
                x_pos = right - width
            else:
                x_pos = left + max(0, (box_width - width) // 2)
            draw.text((x_pos, y_pos), line, fill="black", font=font)
            y_pos += line_height
        return

    if layer.mode == "Barcode":
        image = generate_barcode(layer.text)
    elif layer.mode == "QR":
        image = generate_qr(layer.text)
    else:
        return

    image.thumbnail((box_width, box_height), RESAMPLE)
    if layer.align == "Left":
        x_pos = left
    elif layer.align == "Right":
        x_pos = right - image.width
    else:
        x_pos = left + max(0, (box_width - image.width) // 2)
    y_pos = top + max(0, (box_height - image.height) // 2)
    base.paste(image, (x_pos, y_pos))


def render_label(settings: AppSettings, current_image: Image.Image | None) -> Image.Image:
    width, height = label_pixel_size(settings)
    canvas = Image.new("RGB", (width, height), "white")
    if current_image is not None:
        paste_background_image(canvas, current_image)

    boxes = layer_boxes(settings, width, height)
    for index, layer in ((1, settings.layer1), (2, settings.layer2)):
        box = boxes.get(index)
        if box:
            render_layer(canvas, settings, layer, box)
    return canvas
