import unittest

from PIL import Image

from pigeon_label_maker.models import AppSettings
from pigeon_label_maker.printing import apply_print_processing, build_print_command, chunk_bytes, validate_settings


class PrintingTests(unittest.TestCase):
    def test_apply_print_processing_returns_1bit_image(self) -> None:
        settings = AppSettings()
        source = Image.new("RGB", (400, 120), "white")
        processed = apply_print_processing(source, settings)
        self.assertEqual(processed.mode, "1")
        self.assertEqual(processed.width % 8, 0)

    def test_build_print_command_contains_bitmap_payload(self) -> None:
        settings = AppSettings()
        source = Image.new("1", (112, 320), 1)
        command = build_print_command(source, settings)
        self.assertIn(b"BITMAP 0,0,14,320,1,", command)
        self.assertTrue(command.endswith(b"\r\nPRINT 1\r\n"))

    def test_validation_requires_port_when_printing(self) -> None:
        settings = AppSettings()
        settings.layer1.text = "ABC123"
        errors = validate_settings(settings)
        self.assertTrue(any("printer port" in error.lower() for error in errors))

    def test_validation_requires_ble_fields_when_ble_selected(self) -> None:
        settings = AppSettings()
        settings.layer1.text = "ABC123"
        settings.output_mode = "BLE"
        errors = validate_settings(settings)
        self.assertTrue(any("ble printer" in error.lower() for error in errors))
        self.assertTrue(any("characteristic" in error.lower() for error in errors))

    def test_chunk_bytes_splits_payload(self) -> None:
        chunks = chunk_bytes(b"abcdefghij", 4)
        self.assertEqual(chunks, [b"abcd", b"efgh", b"ij"])


if __name__ == "__main__":
    unittest.main()
