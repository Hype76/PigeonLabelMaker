# Pigeon Label Maker

Pigeon Label Maker is a Windows label design and thermal printing app for small TSPL-compatible label printers. The desktop experience is an Electron shell backed by a Python rendering and print engine.

## Current App Overview

- Canvas-based label designer with draggable, resizable, and rotatable elements
- Text, QR, barcode, and image elements
- Theme toggle and simple or advanced UI mode
- COM and BLE printer support
- Live processed print preview generated from the design canvas
- Custom label sizes plus quick size presets
- Local recent label history
- Save and load design layout in the Electron UI
- Undo, redo, duplicate, lock, layer order, and keyboard nudge tools
- Expiry label helper and quick symbol picker
- Saved design thumbnails and PNG export
- Printer profiles for repeated printer and label setups
- Print calibration with X offset, Y offset, scale, and calibration test print
- First-launch release notes after updates
- Windows installer build via Electron Builder and PyInstaller

## Main Features

### Designer

- Add text, QR, barcode, and image elements
- Drag, resize, rotate, align, edit, and delete selected items
- Canvas safe area, center guides, and sticky edge bounds
- Snap to canvas center and other elements
- Font picker, manual font size, and optional system font discovery
- Bold and uppercase toggles for text elements
- Template shortcuts for `Text`, `Dual`, `QR`, and `Barcode`
- Quick symbol buttons for common label characters such as `°`, `%`, `£`, `€`, `©`, and `✓`
- Undo, redo, duplicate, lock, bring forward, send backward, fit text to box, and export PNG tools

### Print Pipeline

- The Electron app captures the design canvas and sends that image to the Python backend
- Preview and print use the same captured canvas image path
- Print calibration can shift and scale the final processed output
- Thermal image processing supports:
  - threshold mode for clean logo-style output
  - dither mode for photo-like output
  - brightness, contrast, threshold, invert, auto optimize, and edge enhance controls

### Connections

- Serial COM printing with explicit `Connect COM`
- BLE scan, connect, disconnect, inspect, battery query, and print
- Persistent connection handling so normal print jobs do not intentionally tear down the printer session
- COM speed is fixed to `115200` in the Electron UI
- Printer profiles store repeated COM, BLE, label size, image, and calibration setups
- Connection feedback includes connection status, timer, test badge, and BLE battery panel
- BLE battery now prefers showing `Unknown` over a false dead reading when the printer does not expose a reliable value

### UX

- One-time onboarding overlay
- Help overlay
- Inline error messaging in the Connection panel instead of modal popups
- Human-readable connection error text
- Print progress for multi-copy jobs
- Queue progress beside the copies control
- Stop button for active print queues
- Print success pulse
- Connection type, status, and timer in the action bar
- Three-column layout with the Connection panel on the right
- App version shown in the window title
- One-time update notes after installing a new version

## Project Structure

```text
PigeonLabelMaker/
|-- backend_entry.py
|-- package.json
|-- package-lock.json
|-- README.md
|-- requirements.txt
|-- requirements-dev.txt
|-- electron/
|   |-- icon.ico
|   |-- icon.png
|   |-- index.html
|   |-- main.js
|   |-- preload.js
|   |-- renderer.js
|   |-- styles.css
|-- pigeon_label_maker/
|   |-- __init__.py
|   |-- backend_service.py
|   |-- config.py
|   |-- models.py
|   |-- presets.py
|   |-- printing.py
|   |-- rendering.py
|-- tests/
|   |-- test_printing.py
|   |-- test_rendering.py
```

## Requirements

- Windows
- Python 3.11 or newer recommended
- Node.js and npm
- A TSPL-compatible printer over COM or BLE

## Install

```powershell
python -m pip install -r requirements.txt
npm install
```

## Run

### Electron

```powershell
npm start
```

Electron starts the Python backend automatically. In development it runs:

```powershell
python -m pigeon_label_maker.backend_service
```

## How to Use

1. Launch the Electron app with `npm start`.
2. Choose a label size preset or apply a custom size.
3. Add text, QR, barcode, or image elements from the Designer panel.
4. Move and resize elements on the canvas.
5. Adjust image processing settings if you are printing artwork or photos.
6. Select a COM port and click `Connect COM`, or scan and connect over BLE.
7. Set copies if needed and use `Stop` to stop the remaining queue after the current label.
8. Click `Print Label`.

## Keyboard Shortcuts

- `Ctrl+Enter`: Print label
- `Ctrl+L`: Focus selected content input
- `Ctrl+Backspace`: Clear selected text item text
- `Ctrl+T`: Toggle theme
- `Ctrl+M`: Toggle simple or advanced mode
- `Delete`: Delete selected canvas item
- `Ctrl+Z`: Undo
- `Ctrl+Y`: Redo
- `Ctrl+D`: Duplicate selected item
- `Arrow Keys`: Nudge selected item
- `Shift+Arrow Keys`: Larger nudge step

## Connection Notes

### COM

- COM ports are not auto-selected on startup
- Pick the port you want from the Connection panel
- Click `Connect COM` before printing
- The Electron UI keeps the COM speed fixed to `115200`

### BLE

- Scan for nearby devices
- Select a device
- Click `Connect`
- The app automatically chooses a writable printer characteristic
- Use `Disconnect` to intentionally end the BLE session

## Settings and Local Data

The Python backend stores app files under:

- Settings: `%LOCALAPPDATA%\\PigeonLabelMaker\\settings.json`
- Logs: `%LOCALAPPDATA%\\PigeonLabelMaker\\logs\\app.log`

The Electron UI also stores some local browser-style state such as:

- theme
- simple or advanced mode
- onboarding completion
- recent labels
- saved canvas design
- saved design thumbnail gallery

## Testing

```powershell
python -m unittest discover -s tests -v
```

## Build a Windows Installer

Install the Python and Node dependencies first, then run:

```powershell
npm run build
```

That build does two things:

1. freezes the Python backend with PyInstaller using `backend_entry.py`
2. packages the Electron app with Electron Builder

Expected outputs:

- frozen backend in `python-dist/`
- Windows installer in `dist/`

## Publish Updates with GitHub Releases

Auto-update is powered by Electron Builder and GitHub Releases.

For a normal local installer build:

```powershell
npm run build
```

For a release that installed apps can detect:

```powershell
$env:GH_TOKEN="your_github_token"
npm run release
```

Release notes:

- bump the `version` in `package.json` before publishing
- add release notes for the new version in `electron/renderer.js`
- the GitHub token needs permission to create releases in `Hype76/PigeonLabelMaker`
- installed builds can use `Check Updates` in the App panel
- development mode shows that updates only work in the installed app
- uploaded release assets should include the installer and generated update metadata from `dist/`
- updater errors are shortened in the App panel so raw GitHub responses are not shown to users

## Troubleshooting

### Preview is black or wrong

- make sure the design canvas is visible and not covered by another window layer
- check `%LOCALAPPDATA%\\PigeonLabelMaker\\logs\\app.log`
- if you changed image processing recently, try switching between `Clean (Logo)` and `Photo (Dither)`

### Printer does not print

- confirm the correct COM port or BLE device is selected
- for COM, click `Connect COM` first
- verify the printer accepts TSPL commands

### Fonts in print output look wrong

- the Electron app captures the canvas as an image before preview and print
- fonts that are not fully available to Chromium on the machine can still render differently
- prefer common installed Windows fonts if exact matching matters

### BLE issues

- rescan if the device list is stale
- reconnect manually if the printer was powered off and back on
- some printers expose several writable characteristics, but the app will prefer the known printer one automatically
- battery level may show `Unknown` on devices that do not expose a reliable BLE battery value

## Development Notes

- The Electron UI is intentionally thin and talks to the Python backend over stdio
- Rendering, printer command generation, serial I/O, and BLE logic stay in Python
- The current print path is image-first, meaning the frontend canvas is the source of truth for preview and print
