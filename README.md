# Spotlight Saver

A lightweight Windows desktop app that extracts and saves the wallpaper images from **Windows Spotlight** — the rotating lock screen backgrounds that Windows downloads automatically.

<img width="960" height="563" alt="{84F527A9-ADE3-4242-BECF-1067A2A27F6A}" src="https://github.com/user-attachments/assets/f00f455d-114e-4318-954c-b1e9b0cd6875" />


![Windows 10/11](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?logo=windows)
![Python 3.8+](https://img.shields.io/badge/Python-3.8+-3776AB?logo=python&logoColor=white)

## Features

- **Auto-detect** Spotlight image folders on Windows 10 and Windows 11
- **Thumbnail grid** with grouped display by source
- **Resolution filter** to skip low-quality images
- **Duplicate detection** via MD5 hash history — never saves the same image twice
- **Background monitoring** with system tray integration — auto-copies new wallpapers as they appear
- **Windows toast notifications** when new images are found
- **Start with Windows** option via registry autostart
- **Localized UI** — English and Spanish, auto-detected from system language
- **Dark theme** interface
- **Custom folder scanning** for manually downloaded Spotlight images

## Quick Start

### Option 1: Download the executable

Download `SpotlightSaver.exe` from [Releases](../../releases), run it — no installation needed.

### Option 2: Run from source

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/spotlight-saver.git
cd spotlight-saver

# Install dependencies
pip install -r requirements.txt

# Run
python spotlight_saver.py
```

## Building

To compile a standalone `.exe`:

```bash
build.bat
```

This installs dependencies and builds a single-file executable at `dist/SpotlightSaver.exe`.

## How It Works

Windows Spotlight downloads high-quality wallpapers to hidden system folders:

| Source | Path |
|--------|------|
| W10/W11 | `%LOCALAPPDATA%\Packages\Microsoft.Windows.ContentDeliveryManager_...\LocalState\Assets` |
| W11 | `%LOCALAPPDATA%\Packages\MicrosoftWindows.Client.CBS_...\LocalCache\Microsoft\IrisService` |

Spotlight Saver reads these folders, identifies valid landscape images (filtering out icons and portrait images), and lets you save them with a single click. It tracks what you've already saved so you never get duplicates.

## Dependencies

| Package | Purpose |
|---------|---------|
| [Pillow](https://pypi.org/project/Pillow/) | Image loading and thumbnails |
| [pystray](https://pypi.org/project/pystray/) | System tray integration *(optional)* |
| [watchdog](https://pypi.org/project/watchdog/) | Real-time folder monitoring *(optional)* |
| [winotify](https://pypi.org/project/winotify/) | Windows toast notifications *(optional)* |

The app works without the optional dependencies — monitoring and tray features are simply disabled.

## Usage

1. **Launch** the app — it automatically scans for Spotlight images
2. **Review** the thumbnail grid; images meeting the resolution filter are preselected
3. **Click "Save selected"** to copy them to your Documents folder
4. **Enable monitoring** to run in the background and auto-save new wallpapers

### Command-line flags

| Flag | Description |
|------|-------------|
| `--minimized` | Start minimized to tray with monitoring active (used by autostart) |

## License

[MIT](LICENSE)
