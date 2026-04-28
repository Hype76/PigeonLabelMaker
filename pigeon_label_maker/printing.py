from __future__ import annotations

import asyncio
import math
import threading
import uuid
from dataclasses import dataclass
from typing import Any, Coroutine, TypeVar

from PIL import Image, ImageEnhance, ImageFilter, ImageOps, ImageStat
import serial
import serial.tools.list_ports

from .models import AppSettings
from .rendering import RESAMPLE, generate_barcode, generate_qr, mm_to_pixels

try:
    from bleak import BleakClient, BleakScanner
    from bleak.exc import BleakError
except ImportError:  # pragma: no cover
    BleakClient = None
    BleakScanner = None

    class BleakError(Exception):
        pass

try:
    from winrt.windows.devices.bluetooth import BluetoothCacheMode, BluetoothLEDevice
    from winrt.windows.devices.bluetooth.genericattributeprofile import (
        GattClientCharacteristicConfigurationDescriptorValue,
        GattCommunicationStatus,
        GattWriteOption,
    )
    from winrt.windows.storage.streams import DataReader, DataWriter
except ImportError:  # pragma: no cover
    BluetoothCacheMode = None
    BluetoothLEDevice = None
    GattClientCharacteristicConfigurationDescriptorValue = None
    GattCommunicationStatus = None
    GattWriteOption = None
    DataReader = None
    DataWriter = None


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


@dataclass(frozen=True)
class BLEDiscoveredDevice:
    name: str
    address: str
    rssi: int | None = None

    @property
    def label(self) -> str:
        if self.rssi is None:
            return f"{self.name} [{self.address}]"
        return f"{self.name} [{self.address}] RSSI {self.rssi}"


@dataclass(frozen=True)
class BLEWritableCharacteristic:
    uuid: str
    properties: tuple[str, ...]
    description: str = ""
    preferred: bool = False
    handle: int = 0

    @property
    def label(self) -> str:
        flags = ", ".join(self.properties)
        prefix = "[Preferred] " if self.preferred else ""
        handle_text = f" handle 0x{self.handle:04x}" if self.handle else ""
        if self.description:
            return f"{prefix}{self.uuid} ({flags}){handle_text} {self.description}"
        return f"{prefix}{self.uuid} ({flags}){handle_text}"


@dataclass(frozen=True)
class BLEConnectionState:
    connected: bool
    address: str = ""
    name: str = ""


@dataclass(frozen=True)
class BLEPrintResult:
    protocol: str
    address: str
    characteristic_uuid: str
    notify_uuid: str
    bytes_sent: int
    chunk_size: int
    chunk_count: int
    response: bool
    packet_count: int
    notification_count: int
    notifications: tuple[str, ...]
    first_packet_hex: str
    second_packet_hex: str
    first_data_packet_hex: str
    last_packet_hex: str


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


_ble_loop: asyncio.AbstractEventLoop | None = None
_ble_thread: threading.Thread | None = None
_ble_ready = threading.Event()
_ble_devices_by_address: dict[str, Any] = {}
_ble_connected_client: Any = None
_ble_connected_address: str = ""
_ble_connected_name: str = ""
_ble_session_address: str = ""
_ble_session_name: str = ""
_ble_active_notify_uuid: str = ""
_ble_notification_buffer: list[str] = []
_serial_connection: serial.Serial | None = None
_serial_connection_port: str = ""
_serial_connection_baud_rate: int = 0
_winrt_ble_device: Any = None
_winrt_ble_address: str = ""
_winrt_ble_write_uuid: str = ""
_winrt_ble_write_char: Any = None
_winrt_ble_notify_char: Any = None
_winrt_ble_notify_uuid: str = ""

KNOWN_PRINTER_WRITE_UUIDS = (
    "0000ff02-0000-1000-8000-00805f9b34fb",
    "0000ff04-0000-1000-8000-00805f9b34fb",
    "0000ae3b-0000-1000-8000-00805f9b34fb",
    "0000ae01-0000-1000-8000-00805f9b34fb",
)
KNOWN_PRINTER_NOTIFY_UUIDS = {
    "0000ff02-0000-1000-8000-00805f9b34fb": "0000ff01-0000-1000-8000-00805f9b34fb",
    "0000ff04-0000-1000-8000-00805f9b34fb": "0000ff03-0000-1000-8000-00805f9b34fb",
    "0000ae3b-0000-1000-8000-00805f9b34fb": "0000ae3c-0000-1000-8000-00805f9b34fb",
    "0000ae01-0000-1000-8000-00805f9b34fb": "0000ae02-0000-1000-8000-00805f9b34fb",
}
BLE_TSPL_CHUNK_SIZE = 244
P21_BLE_WIDTH_PX = 96
P21_BLE_HEIGHT_PX = 284
P21_BLE_ROW_BYTES = 12
P21_BLE_BITMAP_LENGTH = P21_BLE_ROW_BYTES * P21_BLE_HEIGHT_PX
P21_BLE_DENSITY = 15
BLE_TSPL_PREFLIGHT_COMMANDS = (b"\x1b!?\r\n", b"\x1b!o\r\n")
BLE_TSPL_PREFLIGHT_EXTENDED_COMMANDS = (
    b"CONFIG?\r\n",
    b"BATTERY?\r\n",
    b"CONFIG?\r\n",
    b"\x1b!?\r\n",
    b"\x1b!o\r\n",
)
BLE_TSPL_PREFLIGHT_DELAYS = (0.22, 0.075)
BLE_TSPL_DATA_DELAY = 0.001


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


def apply_print_calibration(image: Image.Image, settings: AppSettings, size: tuple[int, int]) -> Image.Image:
    target_width, target_height = size
    scale = max(0.5, min(1.5, float(getattr(settings, "print_scale", 1.0) or 1.0)))
    offset_x = mm_to_pixels(float(getattr(settings, "print_offset_x_mm", 0.0) or 0.0), settings.print_dpi)
    offset_y = mm_to_pixels(float(getattr(settings, "print_offset_y_mm", 0.0) or 0.0), settings.print_dpi)

    scaled_width = max(1, round(target_width * scale))
    scaled_height = max(1, round(target_height * scale))
    scaled = image.resize((scaled_width, scaled_height), RESAMPLE)

    fill = image.getpixel((0, 0)) if image.width and image.height else 255
    calibrated = Image.new("L", (target_width, target_height), fill)
    paste_x = round((target_width - scaled_width) / 2 + offset_x)
    paste_y = round((target_height - scaled_height) / 2 + offset_y)
    calibrated.paste(scaled, (paste_x, paste_y))
    return calibrated


def apply_print_processing(image: Image.Image, settings: AppSettings) -> Image.Image:
    image = image.convert("RGB")
    grayscale = ImageOps.grayscale(image)
    grayscale = ImageEnhance.Contrast(grayscale).enhance(settings.contrast)

    if not settings.invert:
        grayscale = ImageOps.invert(grayscale)

    width_px, height_px = print_pixel_size(settings)
    target_size = (height_px, width_px)
    grayscale = grayscale.resize(target_size, RESAMPLE)
    grayscale = apply_print_calibration(grayscale, settings, target_size)

    if settings.image_mode == "dither":
        thresholded = grayscale.convert("1")
    else:
        thresholded = grayscale.point(
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
    if settings.print_scale < 0.5 or settings.print_scale > 1.5:
        errors.append("Print scale must be between 0.5 and 1.5.")
    if settings.print_offset_x_mm < -30 or settings.print_offset_x_mm > 30:
        errors.append("Print X offset must be between -30 mm and 30 mm.")
    if settings.print_offset_y_mm < -30 or settings.print_offset_y_mm > 30:
        errors.append("Print Y offset must be between -30 mm and 30 mm.")
    if settings.threshold < 0 or settings.threshold > 255:
        errors.append("Threshold must be between 0 and 255.")
    if settings.output_mode == "Printer" and not settings.port.strip():
        errors.append("Select a printer port before printing.")
    if settings.output_mode == "BLE":
        if BleakScanner is None:
            errors.append("BLE support is not installed. Install the bleak package.")
        if not settings.ble_device_address.strip():
            errors.append("Select a BLE printer before printing.")
        if not settings.ble_write_char_uuid.strip():
            errors.append("Select a BLE writable characteristic before printing.")
        if settings.ble_chunk_size < 20:
            errors.append("BLE chunk size must be at least 20 bytes.")

    if settings.canvas_layout:
        has_content = False
        for index, item in enumerate(settings.canvas_layout, start=1):
            if not isinstance(item, dict):
                continue
            item_type = str(item.get("type", "text")).lower()
            if item_type == "image":
                if item.get("imageData"):
                    has_content = True
                continue

            text = str(item.get("text", "") or "").strip()
            if not text:
                continue

            has_content = True
            if item_type == "barcode":
                try:
                    generate_barcode(text)
                except Exception as exc:
                    errors.append(f"Canvas item {index} barcode is invalid: {exc}")
            if item_type == "qr":
                try:
                    generate_qr(text)
                except Exception as exc:
                    errors.append(f"Canvas item {index} QR content is invalid: {exc}")

        if not has_content:
            errors.append("Add at least one design element before printing.")
        return errors

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
    global _serial_connection, _serial_connection_baud_rate, _serial_connection_port

    target_port = settings.port.strip()
    target_baud_rate = int(settings.baud_rate)
    if not target_port:
        raise ValueError("Printer port is required.")

    needs_new_connection = (
        _serial_connection is None
        or not getattr(_serial_connection, "is_open", False)
        or _serial_connection_port != target_port
        or _serial_connection_baud_rate != target_baud_rate
    )

    if needs_new_connection:
        disconnect_serial()
        _serial_connection = serial.Serial(target_port, target_baud_rate, timeout=2)
        _serial_connection_port = target_port
        _serial_connection_baud_rate = target_baud_rate

    try:
        _serial_connection.write(command)
        _serial_connection.flush()
    except Exception:
        disconnect_serial()
        raise


def connect_serial(settings: AppSettings) -> None:
    global _serial_connection, _serial_connection_baud_rate, _serial_connection_port

    target_port = settings.port.strip()
    target_baud_rate = int(settings.baud_rate)
    if not target_port:
        raise ValueError("Printer port is required.")

    needs_new_connection = (
        _serial_connection is None
        or not getattr(_serial_connection, "is_open", False)
        or _serial_connection_port != target_port
        or _serial_connection_baud_rate != target_baud_rate
    )

    if needs_new_connection:
        disconnect_serial()
        _serial_connection = serial.Serial(target_port, target_baud_rate, timeout=2)
        _serial_connection_port = target_port
        _serial_connection_baud_rate = target_baud_rate


def disconnect_serial() -> None:
    global _serial_connection, _serial_connection_baud_rate, _serial_connection_port

    if _serial_connection is not None:
        try:
            if getattr(_serial_connection, "is_open", False):
                _serial_connection.close()
        finally:
            _serial_connection = None
            _serial_connection_port = ""
            _serial_connection_baud_rate = 0


def close_winrt_ble_device() -> None:
    global _winrt_ble_address, _winrt_ble_device, _winrt_ble_notify_char
    global _winrt_ble_notify_uuid, _winrt_ble_write_char, _winrt_ble_write_uuid

    if _winrt_ble_device is not None:
        try:
            close_device = getattr(_winrt_ble_device, "close", None)
            if callable(close_device):
                close_device()
        finally:
            _winrt_ble_device = None
            _winrt_ble_address = ""
            _winrt_ble_write_uuid = ""
            _winrt_ble_write_char = None
            _winrt_ble_notify_char = None
            _winrt_ble_notify_uuid = ""


async def discover_ble_devices_async(timeout: float = 5.0) -> list[BLEDiscoveredDevice]:
    if BleakScanner is None:
        raise RuntimeError("BLE support is not installed. Install bleak to use BLE.")

    merged_devices: dict[str, BLEDiscoveredDevice] = {}

    for _attempt in range(2):
        discovered = await BleakScanner.discover(timeout=timeout)
        for device in discovered:
            address = getattr(device, "address", "") or ""
            if not address:
                continue
            name = getattr(device, "name", None) or "Unnamed BLE Device"
            rssi = getattr(device, "rssi", None)
            _ble_devices_by_address[address] = device
            merged_devices[address] = BLEDiscoveredDevice(name=name, address=address, rssi=rssi)

    if not merged_devices:
        discovered_with_adv = await BleakScanner.discover(timeout=timeout, return_adv=True)
        for _address, item in discovered_with_adv.items():
            device, advertisement = item
            address = getattr(device, "address", "") or ""
            if not address:
                continue
            name = device.name or advertisement.local_name or "Unnamed BLE Device"
            _ble_devices_by_address[address] = device
            merged_devices[address] = BLEDiscoveredDevice(
                name=name,
                address=address,
                rssi=getattr(advertisement, "rssi", None),
            )

    devices = list(merged_devices.values())
    devices.sort(key=lambda entry: (entry.name.lower(), entry.address))
    return devices


def discover_ble_devices(timeout: float = 5.0) -> list[BLEDiscoveredDevice]:
    return run_ble_coro_in_worker(discover_ble_devices_async(timeout=timeout))


async def list_ble_writable_characteristics_async(
    address: str,
    pair: bool = False,
) -> list[BLEWritableCharacteristic]:
    client = await ensure_ble_connection_async(address=address, pair=pair)
    return extract_writable_characteristics(client)


def list_ble_writable_characteristics(
    address: str,
    pair: bool = False,
) -> list[BLEWritableCharacteristic]:
    return run_ble_coro_in_worker(
        list_ble_writable_characteristics_async(address=address, pair=pair)
    )


async def list_ble_notify_characteristics_async(
    address: str,
    pair: bool = False,
) -> list[BLEWritableCharacteristic]:
    client = await ensure_ble_connection_async(address=address, pair=pair)
    return extract_notify_characteristics(client)


def list_ble_notify_characteristics(
    address: str,
    pair: bool = False,
) -> list[BLEWritableCharacteristic]:
    return run_ble_coro_in_worker(
        list_ble_notify_characteristics_async(address=address, pair=pair)
    )


def chunk_bytes(payload: bytes, chunk_size: int) -> list[bytes]:
    return [payload[index : index + chunk_size] for index in range(0, len(payload), chunk_size)]


def mac_address_to_int(address: str) -> int:
    normalized = address.replace(":", "").replace("-", "").strip()
    return int(normalized, 16)


def make_winrt_buffer(data: bytes) -> Any:
    if DataWriter is None:
        raise RuntimeError("WinRT storage streams are unavailable.")
    writer = DataWriter()
    writer.write_bytes(bytes(data))
    return writer.detach_buffer()


def read_winrt_buffer_bytes(buffer: Any) -> bytes:
    if DataReader is None:
        return b""
    reader = DataReader.from_buffer(buffer)
    data = bytearray(buffer.length)
    reader.read_bytes(data)
    return bytes(data)


def normalize_uuid_text(value: Any) -> str:
    text = str(value).strip().lower().strip("{}")
    try:
        return str(uuid.UUID(text))
    except Exception:
        return text


def build_delayed_payload_sequence(
    preflight_commands: tuple[bytes, ...],
    command_chunks: list[bytes],
    preflight_delays: tuple[float, ...] | None = None,
    data_delay: float = 0.0,
) -> tuple[tuple[float, bytes], ...]:
    sequence: list[tuple[float, bytes]] = []
    pending_delay = 0.0
    for index, preflight in enumerate(preflight_commands):
        sequence.append((pending_delay, preflight))
        if preflight_delays and index < len(preflight_delays):
            pending_delay = preflight_delays[index]
        else:
            pending_delay = 0.0
    for chunk in command_chunks:
        sequence.append((pending_delay, chunk))
        pending_delay = data_delay
    return tuple(sequence)


async def send_raw_to_ble_printer_async(command: bytes, settings: AppSettings) -> BLEPrintResult:
    address = settings.ble_device_address.strip()
    characteristic_uuid = settings.ble_write_char_uuid.strip()
    response = settings.ble_write_with_response

    client = await ensure_ble_connection_async(address=address, pair=settings.ble_pair)
    target_characteristic = find_characteristic(client, characteristic_uuid)
    if target_characteristic is None:
        raise BleakError(f"Characteristic {characteristic_uuid} was not found on device {address}.")

    notify_uuid = await configure_ble_notifications_async(client, characteristic_uuid)

    max_without_response = getattr(target_characteristic, "max_write_without_response_size", None) or 20
    if response:
        chunk_size = min(settings.ble_chunk_size, 512)
    else:
        chunk_size = min(settings.ble_chunk_size, max_without_response)
    if chunk_size < 20:
        chunk_size = 20

    chunks = chunk_bytes(command, chunk_size)
    _ble_notification_buffer.clear()
    for chunk in chunks:
        await client.write_gatt_char(target_characteristic, chunk, response=response)
    await asyncio.sleep(0.35)

    return BLEPrintResult(
        protocol="raw",
        address=address,
        characteristic_uuid=characteristic_uuid,
        notify_uuid=notify_uuid,
        bytes_sent=len(command),
        chunk_size=chunk_size,
        chunk_count=len(chunks),
        response=response,
        packet_count=len(chunks),
        notification_count=len(_ble_notification_buffer),
        notifications=tuple(_ble_notification_buffer),
        first_packet_hex=chunks[0].hex() if chunks else "",
        second_packet_hex=chunks[1].hex() if len(chunks) > 1 else "",
        first_data_packet_hex=chunks[0].hex() if chunks else "",
        last_packet_hex=chunks[-1].hex() if chunks else "",
    )


def normalize_ble_tspl_command(command: bytes) -> bytes:
    normalized = command
    if normalized.startswith(b"\x1b!o\r\n"):
        normalized = normalized[len(b"\x1b!o\r\n") :]
    normalized = normalized.replace(b"DIRECTION 1,1\r\n", b"DIRECTION 0,0\r\n")
    return normalized


def build_p21_ble_command_stream(image: Image.Image, settings: AppSettings) -> bytes:
    grayscale = ImageOps.grayscale(image)
    grayscale = ImageOps.autocontrast(grayscale)
    grayscale = ImageEnhance.Contrast(grayscale).enhance(settings.contrast)
    if grayscale.width > grayscale.height:
        grayscale = grayscale.rotate(90, expand=True)

    nearest = Image.Resampling.NEAREST if hasattr(Image, "Resampling") else Image.NEAREST
    grayscale.thumbnail((P21_BLE_WIDTH_PX, P21_BLE_HEIGHT_PX), nearest)
    canvas = Image.new("L", (P21_BLE_WIDTH_PX, P21_BLE_HEIGHT_PX), 255)
    x_pos = (P21_BLE_WIDTH_PX - grayscale.width) // 2
    y_pos = (P21_BLE_HEIGHT_PX - grayscale.height) // 2
    canvas.paste(grayscale, (x_pos, y_pos))

    binary = canvas.convert("1", dither=Image.Dither.FLOYDSTEINBERG)
    payload = binary.tobytes()
    if len(payload) < P21_BLE_BITMAP_LENGTH:
        payload = payload.ljust(P21_BLE_BITMAP_LENGTH, b"\xff")
    elif len(payload) > P21_BLE_BITMAP_LENGTH:
        payload = payload[:P21_BLE_BITMAP_LENGTH]

    command = b"SIZE 14.0 mm,40.0 mm\r\n"
    command += b"GAP 5.0 mm,0 mm\r\n"
    command += b"DIRECTION 0,0\r\n"
    command += f"DENSITY {P21_BLE_DENSITY}\r\n".encode()
    command += b"CLS\r\n"
    command += b"BITMAP 0,18,12,284,1," + payload
    command += f"\r\nPRINT {settings.copies}\r\n".encode()
    return command


async def send_wrapped_tspl_stream_async(
    command_stream: bytes,
    characteristic_uuid: str,
    settings: AppSettings,
    preflight_commands: tuple[bytes, ...] = BLE_TSPL_PREFLIGHT_COMMANDS,
    response_override: bool | None = None,
    preflight_delays: tuple[float, ...] | None = None,
    data_delay: float = BLE_TSPL_DATA_DELAY,
) -> BLEPrintResult:
    address = settings.ble_device_address.strip()
    client = await ensure_ble_connection_async(address=address, pair=settings.ble_pair)
    target_characteristic = find_characteristic(client, characteristic_uuid)
    if target_characteristic is None:
        raise BleakError(f"Characteristic {characteristic_uuid} was not found on device {address}.")

    notify_uuid = await configure_ble_notifications_async(client, characteristic_uuid)
    if response_override is None:
        response = "write-without-response" not in getattr(target_characteristic, "properties", ())
    else:
        response = response_override
    command_chunks = chunk_bytes(command_stream, BLE_TSPL_CHUNK_SIZE)
    packet_count = 0
    bytes_sent = 0
    preflight_hexes: list[str] = []
    data_hexes: list[str] = []
    _ble_notification_buffer.clear()

    for packet_index, preflight in enumerate(preflight_commands, start=1):
        await client.write_gatt_char(target_characteristic, preflight, response=response)
        packet_count += 1
        bytes_sent += len(preflight)
        preflight_hexes.append(preflight.hex())
        if preflight_delays and packet_index - 1 < len(preflight_delays):
            await asyncio.sleep(preflight_delays[packet_index - 1])

    for chunk_index, chunk in enumerate(command_chunks, start=1):
        await client.write_gatt_char(target_characteristic, chunk, response=response)
        packet_count += 1
        bytes_sent += len(chunk)
        data_hexes.append(chunk.hex())
        if data_delay > 0:
            await asyncio.sleep(data_delay)

    await asyncio.sleep(0.5)
    return BLEPrintResult(
        protocol="tspl-direct",
        address=address,
        characteristic_uuid=characteristic_uuid,
        notify_uuid=notify_uuid,
        bytes_sent=bytes_sent,
        chunk_size=BLE_TSPL_CHUNK_SIZE,
        chunk_count=len(command_chunks),
        response=response,
        packet_count=packet_count,
        notification_count=len(_ble_notification_buffer),
        notifications=tuple(_ble_notification_buffer),
        first_packet_hex=preflight_hexes[0] if preflight_hexes else "",
        second_packet_hex=preflight_hexes[1] if len(preflight_hexes) > 1 else "",
        first_data_packet_hex=data_hexes[0] if data_hexes else "",
        last_packet_hex=data_hexes[-1] if data_hexes else "",
    )


async def send_wrapped_tspl_ble_printer_async(
    command: bytes,
    rendered_image: Image.Image | None,
    settings: AppSettings,
) -> BLEPrintResult:
    characteristic_uuid = settings.ble_write_char_uuid.strip()
    if rendered_image is not None:
        command_stream = build_p21_ble_command_stream(rendered_image, settings)
    else:
        command_stream = normalize_ble_tspl_command(command)
    return await send_wrapped_tspl_stream_async(
        command_stream=command_stream,
        characteristic_uuid=characteristic_uuid,
        settings=settings,
        preflight_commands=BLE_TSPL_PREFLIGHT_COMMANDS,
        preflight_delays=BLE_TSPL_PREFLIGHT_DELAYS,
        data_delay=BLE_TSPL_DATA_DELAY,
    )


async def send_wrapped_tspl_ble_printer_winrt_async(
    command: bytes,
    rendered_image: Image.Image | None,
    settings: AppSettings,
) -> BLEPrintResult:
    characteristic_uuid = settings.ble_write_char_uuid.strip()
    if rendered_image is not None:
        command_stream = build_p21_ble_command_stream(rendered_image, settings)
        preflight_commands = BLE_TSPL_PREFLIGHT_EXTENDED_COMMANDS
        preflight_delays: tuple[float, ...] = ()
    else:
        command_stream = normalize_ble_tspl_command(command)
        preflight_commands = BLE_TSPL_PREFLIGHT_COMMANDS
        preflight_delays = BLE_TSPL_PREFLIGHT_DELAYS
    command_chunks = chunk_bytes(command_stream, BLE_TSPL_CHUNK_SIZE)
    sequence = build_delayed_payload_sequence(
        preflight_commands=preflight_commands,
        command_chunks=command_chunks,
        preflight_delays=preflight_delays,
        data_delay=BLE_TSPL_DATA_DELAY,
    )
    return await send_winrt_payload_sequence_async(
        settings=settings,
        characteristic_uuid=characteristic_uuid,
        sequence=sequence,
        protocol="tspl-direct-winrt",
        first_data_index=len(preflight_commands),
    )


def send_to_ble_printer(
    command: bytes,
    settings: AppSettings,
    rendered_image: Image.Image | None = None,
) -> BLEPrintResult:
    characteristic_uuid = settings.ble_write_char_uuid.strip().lower()
    if characteristic_uuid in KNOWN_PRINTER_WRITE_UUIDS:
        if BluetoothLEDevice is not None:
            return run_ble_coro_in_worker(
                send_wrapped_tspl_ble_printer_winrt_async(
                    command=command,
                    rendered_image=rendered_image,
                    settings=settings,
                )
            )
        return run_ble_coro_in_worker(
            send_wrapped_tspl_ble_printer_async(
                command=command,
                rendered_image=rendered_image,
                settings=settings,
            )
        )
    return run_ble_coro_in_worker(send_raw_to_ble_printer_async(command=command, settings=settings))

async def find_winrt_ble_characteristics_async(
    address: str,
    characteristic_uuid: str,
    notify_uuid: str,
) -> tuple[Any, Any, Any, str]:
    global _winrt_ble_address, _winrt_ble_device, _winrt_ble_notify_char
    global _winrt_ble_notify_uuid, _winrt_ble_write_char, _winrt_ble_write_uuid

    if (
        BluetoothLEDevice is None
        or BluetoothCacheMode is None
        or GattCommunicationStatus is None
    ):
        raise RuntimeError("Native WinRT BLE support is unavailable.")

    normalized_address = address.strip().lower()
    normalized_characteristic_uuid = normalize_uuid_text(characteristic_uuid)
    normalized_notify_uuid = normalize_uuid_text(notify_uuid) if notify_uuid else ""

    if (
        _winrt_ble_device is not None
        and _winrt_ble_address == normalized_address
        and _winrt_ble_write_uuid == normalized_characteristic_uuid
        and _winrt_ble_write_char is not None
    ):
        return (
            _winrt_ble_device,
            _winrt_ble_write_char,
            _winrt_ble_notify_char,
            _winrt_ble_notify_uuid,
        )

    close_winrt_ble_device()

    device = await BluetoothLEDevice.from_bluetooth_address_async(mac_address_to_int(address))
    if device is None:
        raise RuntimeError(f"WinRT could not open BLE device {address}.")

    write_char = None
    notify_char = None
    discovered_characteristics: list[str] = []
    for cache_mode in (BluetoothCacheMode.UNCACHED, BluetoothCacheMode.CACHED):
        services_result = await device.get_gatt_services_with_cache_mode_async(cache_mode)
        if services_result.status != GattCommunicationStatus.SUCCESS:
            continue
        for service in services_result.services:
            chars_result = await service.get_characteristics_with_cache_mode_async(cache_mode)
            if chars_result.status != GattCommunicationStatus.SUCCESS:
                continue
            for characteristic in chars_result.characteristics:
                current_uuid = normalize_uuid_text(characteristic.uuid)
                handle = int(getattr(characteristic, "attribute_handle", 0) or 0)
                discovered_characteristics.append(f"{current_uuid}@0x{handle:04x}")
                if current_uuid == normalized_characteristic_uuid:
                    write_char = characteristic
                if normalized_notify_uuid and current_uuid == normalized_notify_uuid:
                    notify_char = characteristic
        if write_char is not None:
            break

    if write_char is None:
        discovered = ", ".join(dict.fromkeys(discovered_characteristics)) or "none"
        close_device = getattr(device, "close", None)
        if callable(close_device):
            close_device()
        raise RuntimeError(
            f"WinRT could not find characteristic {characteristic_uuid}. "
            f"Discovered: {discovered}"
        )

    _winrt_ble_device = device
    _winrt_ble_address = normalized_address
    _winrt_ble_write_uuid = normalized_characteristic_uuid
    _winrt_ble_write_char = write_char
    _winrt_ble_notify_char = notify_char
    _winrt_ble_notify_uuid = normalized_notify_uuid

    return device, write_char, notify_char, normalized_notify_uuid


async def send_winrt_payload_sequence_async(
    settings: AppSettings,
    characteristic_uuid: str,
    sequence: tuple[tuple[float, bytes], ...],
    protocol: str,
    first_data_index: int,
) -> BLEPrintResult:
    if (
        BluetoothLEDevice is None
        or GattClientCharacteristicConfigurationDescriptorValue is None
        or GattCommunicationStatus is None
        or GattWriteOption is None
    ):
        raise RuntimeError("Native WinRT BLE support is unavailable.")

    address = settings.ble_device_address.strip()
    normalized_characteristic_uuid = normalize_uuid_text(characteristic_uuid)
    notify_uuid = KNOWN_PRINTER_NOTIFY_UUIDS.get(normalized_characteristic_uuid, "")

    if (
        _winrt_ble_device is None
        or _winrt_ble_address != address.lower()
        or _winrt_ble_write_uuid != normalized_characteristic_uuid
    ):
        await disconnect_ble_async(clear_session=False)

    device, write_char, notify_char, notify_uuid = await find_winrt_ble_characteristics_async(
        address=address,
        characteristic_uuid=normalized_characteristic_uuid,
        notify_uuid=notify_uuid,
    )

    notifications: list[str] = []

    def on_value_changed(_sender: Any, args: Any) -> None:
        try:
            notifications.append(read_winrt_buffer_bytes(args.characteristic_value).hex())
        except Exception:
            pass

    notify_token = None
    operation_failed = False
    try:
        if notify_char is not None:
            notify_token = notify_char.add_value_changed(on_value_changed)
            cccd_result = await notify_char.write_client_characteristic_configuration_descriptor_with_result_async(
                GattClientCharacteristicConfigurationDescriptorValue.NOTIFY
            )
            if cccd_result.status != GattCommunicationStatus.SUCCESS:
                raise RuntimeError(f"WinRT notify setup failed with status {cccd_result.status}.")

        packet_hexes: list[str] = []
        total_bytes = 0
        for delay_seconds, payload in sequence:
            if delay_seconds > 0:
                await asyncio.sleep(delay_seconds)
            write_result = await write_char.write_value_with_result_and_option_async(
                make_winrt_buffer(payload),
                GattWriteOption.WRITE_WITHOUT_RESPONSE,
            )
            if write_result.status != GattCommunicationStatus.SUCCESS:
                raise RuntimeError(f"WinRT write failed with status {write_result.status}.")
            packet_hexes.append(payload.hex())
            total_bytes += len(payload)

        await asyncio.sleep(0.5)
        return BLEPrintResult(
            protocol=protocol,
            address=address,
            characteristic_uuid=normalized_characteristic_uuid,
            notify_uuid=notify_uuid,
            bytes_sent=total_bytes,
            chunk_size=BLE_TSPL_CHUNK_SIZE,
            chunk_count=len(sequence),
            response=False,
            packet_count=len(sequence),
            notification_count=len(notifications),
            notifications=tuple(notifications),
            first_packet_hex=packet_hexes[0] if packet_hexes else "",
            second_packet_hex=packet_hexes[1] if len(packet_hexes) > 1 else "",
            first_data_packet_hex=packet_hexes[first_data_index] if len(packet_hexes) > first_data_index else "",
            last_packet_hex=packet_hexes[-1] if packet_hexes else "",
        )
    except Exception:
        operation_failed = True
        raise
    finally:
        if notify_char is not None:
            try:
                await notify_char.write_client_characteristic_configuration_descriptor_with_result_async(
                    GattClientCharacteristicConfigurationDescriptorValue.NONE
                )
            except Exception:
                pass
        if notify_char is not None and notify_token is not None:
            try:
                notify_char.remove_value_changed(notify_token)
            except Exception:
                pass
        if operation_failed:
            close_winrt_ble_device()




async def ensure_ble_connection_async(address: str, pair: bool = False) -> Any:
    global _ble_connected_address, _ble_connected_client, _ble_connected_name
    global _ble_session_address, _ble_session_name

    if BleakClient is None:
        raise RuntimeError("BLE support is not installed. Install bleak to use BLE.")
    if not address:
        raise RuntimeError("BLE device address is required.")

    if (
        _ble_connected_client is not None
        and _ble_connected_address.lower() == address.lower()
        and getattr(_ble_connected_client, "is_connected", False)
    ):
        return _ble_connected_client

    await disconnect_ble_async()

    device = _ble_devices_by_address.get(address)
    if device is None:
        device = await BleakScanner.find_device_by_address(address, timeout=8.0)
        if device is None:
            raise BleakError(f"BLE device {address} could not be found.")
        _ble_devices_by_address[address] = device

    client = BleakClient(device, pair=pair)
    await client.connect()
    if not client.is_connected:
        raise BleakError(f"BLE device {address} failed to connect.")

    _ble_connected_client = client
    _ble_connected_address = address
    _ble_connected_name = getattr(device, "name", None) or "Unnamed BLE Device"
    _ble_session_address = _ble_connected_address
    _ble_session_name = _ble_connected_name
    return client


def ensure_ble_connection(address: str, pair: bool = False) -> BLEConnectionState:
    return run_ble_coro_in_worker(connect_ble_async(address=address, pair=pair))


async def connect_ble_async(address: str, pair: bool = False) -> BLEConnectionState:
    client = await ensure_ble_connection_async(address=address, pair=pair)
    return BLEConnectionState(
        connected=bool(getattr(client, "is_connected", False)),
        address=_ble_connected_address,
        name=_ble_connected_name,
    )


async def disconnect_ble_async(clear_session: bool = True) -> None:
    global _ble_active_notify_uuid, _ble_connected_address, _ble_connected_client, _ble_connected_name
    global _ble_session_address, _ble_session_name

    close_winrt_ble_device()

    if _ble_connected_client is not None:
        try:
            if _ble_active_notify_uuid and getattr(_ble_connected_client, "is_connected", False):
                try:
                    await _ble_connected_client.stop_notify(_ble_active_notify_uuid)
                except Exception:
                    pass
            if getattr(_ble_connected_client, "is_connected", False):
                await _ble_connected_client.disconnect()
        finally:
            _ble_active_notify_uuid = ""
            _ble_notification_buffer.clear()
            _ble_connected_client = None
            _ble_connected_address = ""
            _ble_connected_name = ""
    if clear_session:
        _ble_session_address = ""
        _ble_session_name = ""


def disconnect_ble() -> None:
    run_ble_coro_in_worker(disconnect_ble_async())


def ble_connection_state() -> BLEConnectionState:
    return run_ble_coro_in_worker(ble_connection_state_async())


async def ble_connection_state_async() -> BLEConnectionState:
    connected = bool(
        _ble_connected_client is not None and getattr(_ble_connected_client, "is_connected", False)
    )
    if not connected and _ble_session_address:
        connected = True
    return BLEConnectionState(
        connected=connected,
        address=_ble_connected_address or _ble_session_address,
        name=_ble_connected_name or _ble_session_name,
    )


def find_characteristic(client: Any, characteristic_uuid: str) -> Any:
    target_uuid = normalize_uuid_text(characteristic_uuid)
    for service in client.services:
        for characteristic in service.characteristics:
            if normalize_uuid_text(characteristic.uuid) == target_uuid:
                return characteristic
    return None


def ble_notification_handler(_sender: Any, data: bytearray) -> None:
    _ble_notification_buffer.append(bytes(data).hex())


async def configure_ble_notifications_async(client: Any, write_uuid: str) -> str:
    global _ble_active_notify_uuid

    notify_uuid = KNOWN_PRINTER_NOTIFY_UUIDS.get(write_uuid.lower(), "")
    if not notify_uuid:
        if _ble_active_notify_uuid:
            try:
                await client.stop_notify(_ble_active_notify_uuid)
            except Exception:
                pass
            _ble_active_notify_uuid = ""
        return ""

    if _ble_active_notify_uuid == notify_uuid:
        return notify_uuid

    if _ble_active_notify_uuid:
        try:
            await client.stop_notify(_ble_active_notify_uuid)
        except Exception:
            pass
        finally:
            _ble_active_notify_uuid = ""

    notify_characteristic = find_characteristic(client, notify_uuid)
    if notify_characteristic is None:
        return ""

    await client.start_notify(notify_characteristic, ble_notification_handler)
    _ble_active_notify_uuid = notify_uuid
    return notify_uuid


async def switch_ble_notify_uuid_async(client: Any, notify_uuid: str) -> str:
    global _ble_active_notify_uuid

    notify_uuid = notify_uuid.strip().lower()
    if not notify_uuid:
        if _ble_active_notify_uuid:
            try:
                await client.stop_notify(_ble_active_notify_uuid)
            except Exception:
                pass
            _ble_active_notify_uuid = ""
        return ""

    if _ble_active_notify_uuid == notify_uuid:
        return notify_uuid

    if _ble_active_notify_uuid:
        try:
            await client.stop_notify(_ble_active_notify_uuid)
        except Exception:
            pass
        finally:
            _ble_active_notify_uuid = ""

    notify_characteristic = find_characteristic(client, notify_uuid)
    if notify_characteristic is None:
        return ""

    await client.start_notify(notify_characteristic, ble_notification_handler)
    _ble_active_notify_uuid = notify_uuid
    return notify_uuid


def extract_writable_characteristics(client: Any) -> list[BLEWritableCharacteristic]:
    characteristics: list[BLEWritableCharacteristic] = []
    for service in client.services:
        for char in service.characteristics:
            properties = tuple(sorted(prop for prop in char.properties if "write" in prop))
            if properties:
                uuid = str(char.uuid).lower()
                characteristics.append(
                    BLEWritableCharacteristic(
                        uuid=normalize_uuid_text(uuid),
                        properties=properties,
                        description=f"service {service.uuid}",
                        preferred=normalize_uuid_text(uuid) in KNOWN_PRINTER_WRITE_UUIDS,
                        handle=int(getattr(char, "handle", 0) or 0),
                    )
                )
    characteristics.sort(key=characteristic_sort_key)
    return characteristics


def extract_notify_characteristics(client: Any) -> list[BLEWritableCharacteristic]:
    characteristics: list[BLEWritableCharacteristic] = []
    for service in client.services:
        for char in service.characteristics:
            properties = tuple(sorted(prop for prop in char.properties if "notify" in prop))
            if properties:
                uuid = str(char.uuid).lower()
                characteristics.append(
                    BLEWritableCharacteristic(
                        uuid=normalize_uuid_text(uuid),
                        properties=properties,
                        description=f"service {service.uuid}",
                        preferred=normalize_uuid_text(uuid) in KNOWN_PRINTER_NOTIFY_UUIDS.values(),
                        handle=int(getattr(char, "handle", 0) or 0),
                    )
                )
    characteristics.sort(key=characteristic_sort_key)
    return characteristics


def characteristic_sort_key(entry: BLEWritableCharacteristic) -> tuple[int, int, int, str]:
    uuid = entry.uuid.lower()
    is_standard_gatt = int(uuid.startswith("00002a") or uuid.startswith("000018"))
    has_write_without_response = int("write-without-response" in entry.properties)
    preferred_rank = (
        KNOWN_PRINTER_WRITE_UUIDS.index(uuid)
        if uuid in KNOWN_PRINTER_WRITE_UUIDS
        else len(KNOWN_PRINTER_WRITE_UUIDS)
    )
    return (preferred_rank, 0 if has_write_without_response else 1, is_standard_gatt, uuid)


T = TypeVar("T")


def run_ble_coro_in_worker(coro: Coroutine[Any, Any, T]) -> T:
    ensure_ble_worker_loop()
    assert _ble_loop is not None
    future = asyncio.run_coroutine_threadsafe(coro, _ble_loop)
    return future.result()


def ensure_ble_worker_loop() -> None:
    global _ble_loop, _ble_thread

    if _ble_thread is not None and _ble_thread.is_alive():
        _ble_ready.wait(timeout=2)
        return

    _ble_ready.clear()

    def worker() -> None:
        global _ble_loop
        try:
            try:
                from bleak.backends.winrt.util import uninitialize_sta
            except ImportError:
                uninitialize_sta = None

            if uninitialize_sta is not None:
                uninitialize_sta()

            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            _ble_loop = loop
            _ble_ready.set()
            loop.run_forever()
        finally:
            if _ble_loop is not None and not _ble_loop.is_closed():
                _ble_loop.close()

    _ble_thread = threading.Thread(target=worker, name="BLEWorker", daemon=True)
    _ble_thread.start()
    if not _ble_ready.wait(timeout=5):
        raise RuntimeError("BLE worker loop failed to start.")
