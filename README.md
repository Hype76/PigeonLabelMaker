# Pigeon Label Maker

Pigeon Label Maker is a Windows desktop tool for building thermal labels with text, barcodes, QR codes, and optional artwork. It uses Tkinter for the interface and generates TSPL printer commands for compatible serial thermal printers.

## What Changed

This project has been refactored from a single-file prototype into a small application package with:

- Modular rendering, printing, settings, and preset logic
- Safer input validation and user-facing error messages
- Persistent settings and user presets
- Real print preview with orientation, contrast, invert, and threshold processing
- PNG export and mock print command output
- Logging for runtime errors and print activity
- Dynamic serial port refresh
- Drag and drop image loading on Windows
- Unit tests for rendering and printer command generation

## Project Structure

```text
PigeonLabelMaker/
|-- main.py
|-- README.md
|-- requirements.txt
|-- requirements-dev.txt
|-- .gitignore
|-- build.ps1
|-- pigeon_label_maker/
|   |-- __init__.py
|   |-- app.py
|   |-- config.py
|   |-- models.py
|   |-- presets.py
|   |-- printing.py
|   |-- rendering.py
|-- tests/
|   |-- test_printing.py
|   |-- test_rendering.py
```

## Install

```powershell
python -m pip install -r requirements.txt
```

## Run

```powershell
python main.py
```

## Usage

1. Pick a preset or start from scratch.
2. Set the content and type for Layer 1 and Layer 2.
3. Load or paste an image if you want artwork behind the text or codes.
4. Choose the printer profile, label size, DPI, density, and threshold.
5. Review both previews:
   - Design Preview shows the full label layout.
   - Print Preview shows the processed black and white output that will be sent to the printer.
6. Choose `Printer` to print directly or `Mock File` to save the raw printer command to disk.
7. Use `Export PNG` if you want a normal image copy of the label.

## Keyboard Shortcuts

- `Ctrl+P`: print or save command
- `Ctrl+S`: export PNG
- `Ctrl+O`: load image
- `Ctrl+Shift+V`: paste image from clipboard
- `F5`: refresh preview
- `Ctrl+1`: uppercase Layer 1
- `Ctrl+2`: uppercase Layer 2
- `Ctrl+Shift+1`: title case Layer 1
- `Ctrl+Shift+2`: title case Layer 2
- `Delete`: clear image

## Settings and Presets

The app stores settings and user presets in your local app data folder:

- Settings: `%LOCALAPPDATA%\\PigeonLabelMaker\\settings.json`
- Presets: `%LOCALAPPDATA%\\PigeonLabelMaker\\user_presets.json`
- Logs: `%LOCALAPPDATA%\\PigeonLabelMaker\\logs\\app.log`

## Testing

```powershell
python -m unittest discover -s tests -v
```

## Build a Windows Executable

```powershell
python -m pip install -r requirements.txt -r requirements-dev.txt
.\build.ps1
```

The packaged executable will be placed in the `dist` folder.

## Printer Notes

- The app generates TSPL `BITMAP` commands.
- Serial settings default to `115200`.
- Use `Refresh Ports` if your printer was plugged in after the app started.
- If your printer output looks too dark or too light, adjust `Density`, `Contrast`, and `Threshold`.

## Troubleshooting

- If preview generation fails, check the log file in `%LOCALAPPDATA%\\PigeonLabelMaker\\logs\\app.log`.
- If printing fails, confirm the selected COM port and that the printer accepts TSPL over serial.
- If drag and drop does not work, reinstall the requirements so the `windnd` package is present.
- If fonts look different on another machine, install the referenced Windows fonts or choose another available font.
