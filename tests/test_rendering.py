import unittest

from PIL import Image, ImageDraw

from pigeon_label_maker.models import AppSettings
from pigeon_label_maker.rendering import flatten_to_rgb, get_font, render_label, wrap_text


class RenderingTests(unittest.TestCase):
    def test_flatten_to_rgb_handles_la_images(self) -> None:
        image = Image.new("LA", (20, 20), color=(128, 255))
        flattened = flatten_to_rgb(image)
        self.assertEqual(flattened.mode, "RGB")
        self.assertEqual(flattened.size, (20, 20))

    def test_render_uses_full_width_when_one_layer_active(self) -> None:
        settings = AppSettings()
        settings.layer1.text = "FULL WIDTH"
        settings.layer1.mode = "Text"
        settings.layer1.align = "Center"
        settings.layer2.mode = "Off"
        rendered = render_label(settings, None)
        bbox = rendered.convert("L").point(lambda pixel: 0 if pixel > 250 else 255).getbbox()
        self.assertIsNotNone(bbox)
        self.assertGreater(bbox[2] - bbox[0], rendered.width // 2)

    def test_wrap_text_does_not_create_blank_first_line(self) -> None:
        draw = ImageDraw.Draw(Image.new("RGB", (400, 100), "white"))
        font = get_font("Arial", 20)
        lines = wrap_text(draw, "SUPERCALIFRAGILISTIC", font, 80, "Auto Wrap")
        self.assertEqual(lines[0], "SUPERCALIFRAGILISTIC")


if __name__ == "__main__":
    unittest.main()
