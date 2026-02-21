"""
Spotlight Saver - Extract and save Windows Spotlight wallpaper images
Supports Windows 10 and Windows 11
With background monitoring and auto-copy
"""

import os
import sys
import shutil
import json
import hashlib
import locale
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from datetime import datetime
from PIL import Image, ImageTk, ImageDraw
from collections import OrderedDict
import winreg

# Optional dependencies for background monitoring
try:
    import pystray
    from pystray import MenuItem as item
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    WATCHDOG_AVAILABLE = True
except ImportError:
    WATCHDOG_AVAILABLE = False

try:
    from winotify import Notification, audio
    TOAST_AVAILABLE = True
except ImportError:
    TOAST_AVAILABLE = False


# === Localization ===

def _detect_language():
    """Detect system language, default to English"""
    lang = ''
    for getter in (locale.getlocale, locale.getdefaultlocale):
        try:
            result = getter()
            if result and result[0]:
                lang = result[0]
                break
        except Exception:
            continue
    return 'es' if lang.startswith('es') else 'en'


LANG = _detect_language()

_STRINGS = {
    'en': {
        'folder_name': 'Spotlight Saved Wallpapers',
        'filter_label': 'Min resolution:',
        'apply_filter': 'Apply filter',
        'select_all': '✓ All',
        'deselect_all': '✗ None',
        'open_folder': '📁 Open folder...',
        'searching_spotlight': 'Searching for Spotlight images...',
        'save_selected': '💾 Save selected',
        'save_selected_n': '💾 Save selected ({})',
        'enable_monitoring': '👁 Enable monitoring',
        'stop_monitoring_btn': '⏹ Stop monitoring',
        'start_with_windows': 'Start with Windows',
        'optional_deps_title': 'Optional Dependencies',
        'optional_deps_msg': 'For background monitoring, install:\npip install {}',
        'scanning_existing': 'Scanning existing files...',
        'new_wallpapers_found': 'New wallpapers found',
        'n_new_images': '{} new images copied',
        'new_spotlight_wallpaper': 'New Spotlight wallpaper!',
        'saved_toast': 'Saved: {}\nResolution: {}x{}',
        'tray_open': 'Open',
        'tray_stop': 'Stop monitoring',
        'tray_exit': 'Exit',
        'tray_monitoring': ' - Monitoring',
        'monitoring_stopped': 'Monitoring stopped',
        'no_folders_found': 'No Spotlight folders found',
        'searching_images': 'Searching for images...',
        'custom_folder': 'Custom folder',
        'no_folders_msg': "No Spotlight folders found.\nUse 'Open folder...' to select manually.",
        'group_header': '📁 {} ({} images)',
        'status_format': '{} group(s) | {} images | {} preselected (≥{}x{})',
        'values_must_be_numbers': 'Values must be numbers',
        'no_images_selected': 'No images selected',
        'saved_n_images': '{} images saved to:\n{}',
        'skipped_n': '\n\nSkipped: {} (portrait or already copied)',
        'errors_n': '\n\nErrors ({}):\n{}',
        'complete': 'Complete',
        'autostart_error': "Could not modify autostart:\n{}",
        'monitor_error': 'Could not start monitoring',
        'select_folder_title': 'Select image folder',
        'warning': 'Warning',
        'error': 'Error',
        'info': 'Info',
    },
    'es': {
        'folder_name': 'Fondos Guardados de Spotlight',
        'filter_label': 'Filtro mínimo:',
        'apply_filter': 'Aplicar filtro',
        'select_all': '✓ Todo',
        'deselect_all': '✗ Nada',
        'open_folder': '📁 Abrir carpeta...',
        'searching_spotlight': 'Buscando imágenes de Spotlight...',
        'save_selected': '💾 Guardar seleccionados',
        'save_selected_n': '💾 Guardar seleccionados ({})',
        'enable_monitoring': '👁 Activar monitoreo',
        'stop_monitoring_btn': '⏹ Detener monitoreo',
        'start_with_windows': 'Iniciar con Windows',
        'optional_deps_title': 'Dependencias opcionales',
        'optional_deps_msg': 'Para monitoreo en background, instala:\npip install {}',
        'scanning_existing': 'Escaneando archivos existentes...',
        'new_wallpapers_found': 'Nuevos fondos encontrados',
        'n_new_images': 'Se copiaron {} imágenes nuevas',
        'new_spotlight_wallpaper': '¡Nuevo fondo de Spotlight!',
        'saved_toast': 'Guardado: {}\nResolución: {}x{}',
        'tray_open': 'Abrir',
        'tray_stop': 'Detener monitoreo',
        'tray_exit': 'Salir',
        'tray_monitoring': ' - Monitoreando',
        'monitoring_stopped': 'Monitoreo detenido',
        'no_folders_found': 'No se encontraron carpetas de Spotlight',
        'searching_images': 'Buscando imágenes...',
        'custom_folder': 'Carpeta personalizada',
        'no_folders_msg': "No se encontraron carpetas de Spotlight.\nUsa 'Abrir carpeta...' para seleccionar manualmente.",
        'group_header': '📁 {} ({} imágenes)',
        'status_format': '{} grupo(s) | {} imágenes | {} preseleccionadas (≥{}x{})',
        'values_must_be_numbers': 'Los valores deben ser números',
        'no_images_selected': 'No hay imágenes seleccionadas',
        'saved_n_images': 'Se guardaron {} imágenes en:\n{}',
        'skipped_n': '\n\nOmitidas: {} (verticales o ya copiadas)',
        'errors_n': '\n\nErrores ({}):\n{}',
        'complete': 'Completado',
        'autostart_error': "No se pudo modificar autostart:\n{}",
        'monitor_error': 'No se pudo iniciar el monitoreo',
        'select_folder_title': 'Seleccionar carpeta con imágenes',
        'warning': 'Aviso',
        'error': 'Error',
        'info': 'Info',
    },
}


def t(key, *args):
    """Get localized string by key, with optional format arguments"""
    s = _STRINGS.get(LANG, _STRINGS['en']).get(key, _STRINGS['en'].get(key, key))
    if args:
        return s.format(*args)
    return s


# === Utilities ===

def get_file_hash(filepath):
    """Generate MD5 hash of a file for unique identification"""
    hasher = hashlib.md5()
    with open(filepath, 'rb') as f:
        buf = f.read(65536)
        while len(buf) > 0:
            hasher.update(buf)
            buf = f.read(65536)
    return hasher.hexdigest()


def is_horizontal(width, height):
    """Return True if the image is horizontal (landscape)"""
    return width > height


# === History ===

class HistoryManager:
    """Tracks which files have already been copied via hash-based deduplication"""

    def __init__(self, history_path):
        self.history_path = Path(history_path)
        self.history = self._load()
        self._hash_set = set(self.history['copied_hashes'])

    def _load(self):
        if self.history_path.exists():
            try:
                with open(self.history_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError, KeyError):
                return {'copied_hashes': [], 'copied_files': []}
        return {'copied_hashes': [], 'copied_files': []}

    def save(self):
        self.history_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.history_path, 'w', encoding='utf-8') as f:
            json.dump(self.history, f, indent=2, ensure_ascii=False)

    def is_copied(self, file_hash):
        return file_hash in self._hash_set

    def add(self, file_hash, filename, dimensions):
        if file_hash not in self._hash_set:
            self._hash_set.add(file_hash)
            self.history['copied_hashes'].append(file_hash)
            self.history['copied_files'].append({
                'hash': file_hash,
                'original_name': filename,
                'dimensions': f"{dimensions[0]}x{dimensions[1]}",
                'copied_at': datetime.now().isoformat()
            })
            self.save()


# === Monitoring ===

class SpotlightMonitor:
    """Monitors Spotlight folders and auto-copies new images"""

    def __init__(self, paths_to_watch, output_folder, history_manager,
                 min_width=800, min_height=600, on_new_image=None):
        self.paths_to_watch = paths_to_watch
        self.output_folder = Path(output_folder)
        self.history = history_manager
        self.min_width = min_width
        self.min_height = min_height
        self.on_new_image = on_new_image
        self.observer = None
        self.running = False

    def _process_file(self, filepath):
        """Process a single file: validate, hash-check, and copy if new"""
        filepath = Path(filepath)

        if not filepath.is_file():
            return None

        try:
            with Image.open(filepath) as img:
                width, height = img.size
                fmt = img.format or 'JPEG'

            if not is_horizontal(width, height):
                return None

            if width < self.min_width or height < self.min_height:
                return None

            file_hash = get_file_hash(filepath)
            if self.history.is_copied(file_hash):
                return None

            self.output_folder.mkdir(parents=True, exist_ok=True)

            ext = 'jpg' if fmt.upper() == 'JPEG' else fmt.lower()
            date_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            new_name = f"spotlight_{date_str}.{ext}"
            dest_path = self.output_folder / new_name

            counter = 1
            while dest_path.exists():
                new_name = f"spotlight_{date_str}_{counter}.{ext}"
                dest_path = self.output_folder / new_name
                counter += 1

            shutil.copy2(filepath, dest_path)

            self.history.add(file_hash, filepath.name, (width, height))

            return {
                'path': dest_path,
                'dimensions': (width, height),
                'original': filepath.name
            }

        except Exception as e:
            print(f"Error processing {filepath}: {e}")
            return None

    def scan_existing(self):
        """Scan existing files in monitored folders"""
        new_files = []

        for base_path in self.paths_to_watch:
            if not base_path.exists():
                continue

            # IrisService stores images in subdirectories
            if 'IrisService' in str(base_path):
                for subdir in base_path.iterdir():
                    if subdir.is_dir():
                        for file in subdir.iterdir():
                            result = self._process_file(file)
                            if result:
                                new_files.append(result)
            else:
                for file in base_path.iterdir():
                    result = self._process_file(file)
                    if result:
                        new_files.append(result)

        return new_files

    def start(self):
        """Start real-time filesystem monitoring"""
        if not WATCHDOG_AVAILABLE:
            return False

        self.running = True
        self.observer = Observer()

        handler = SpotlightEventHandler(self)

        for path in self.paths_to_watch:
            if path.exists():
                self.observer.schedule(handler, str(path), recursive=True)

        self.observer.start()
        return True

    def stop(self):
        """Stop monitoring"""
        self.running = False
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None


# Placeholder if watchdog is not installed
class SpotlightEventHandler:
    pass

if WATCHDOG_AVAILABLE:
    class SpotlightEventHandler(FileSystemEventHandler):
        """Handles filesystem events with debouncing to avoid duplicate processing"""

        def __init__(self, monitor):
            self.monitor = monitor
            self._recent = {}
            self._lock = threading.Lock()

        def on_created(self, event):
            if event.is_directory:
                return

            filepath = event.src_path
            now = time.time()

            # Debounce: skip if same file was seen within the last 5 seconds
            with self._lock:
                if filepath in self._recent and now - self._recent[filepath] < 5:
                    return
                self._recent[filepath] = now

            def process():
                # Wait for the file to finish writing
                time.sleep(1)
                result = self.monitor._process_file(filepath)
                if result and self.monitor.on_new_image:
                    self.monitor.on_new_image(result)

            thread = threading.Thread(target=process, daemon=True)
            thread.start()

        def on_modified(self, event):
            self.on_created(event)


# === Main Application ===

class SpotlightSaver:
    # Known Spotlight image locations
    SPOTLIGHT_PATHS = {
        'W10/W11 (ContentDeliveryManager)': Path(os.environ.get('LOCALAPPDATA', '')) /
            'Packages/Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy/LocalState/Assets',
        'W11 (IrisService)': Path(os.environ.get('LOCALAPPDATA', '')) /
            'Packages/MicrosoftWindows.Client.CBS_cw5n1h2txyewy/LocalCache/Microsoft/IrisService',
    }

    THUMBNAIL_SIZE = (150, 100)
    GRID_COLUMNS = 5
    MIN_WIDTH = 800
    MIN_HEIGHT = 600

    APP_NAME = "Spotlight Saver"
    AUTOSTART_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

    def __init__(self, root):
        self.root = root
        self.root.title(self.APP_NAME)
        self.root.geometry("950x750")
        self.root.minsize(800, 600)

        # Localized output folder
        self.output_folder = Path.home() / 'Documents' / t('folder_name')
        self.history_path = self.output_folder / 'history.json'

        # {group: [img_data, ...]} - ordered
        self.grouped_images = OrderedDict()
        self.thumbnail_refs = []
        self.group_frames = {}

        # Monitoring state
        self.history_manager = HistoryManager(self.history_path)
        self.monitor = None
        self.tray_icon = None
        self.monitoring_active = False

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.setup_ui()
        self.load_images_async()

    def setup_ui(self):
        # Top frame - controls
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(fill=tk.X)

        # Resolution filter
        ttk.Label(control_frame, text=t('filter_label')).pack(side=tk.LEFT)

        self.min_w_var = tk.StringVar(value=str(self.MIN_WIDTH))
        self.min_h_var = tk.StringVar(value=str(self.MIN_HEIGHT))

        w_entry = ttk.Entry(control_frame, textvariable=self.min_w_var, width=6)
        w_entry.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(control_frame, text="x").pack(side=tk.LEFT, padx=2)
        h_entry = ttk.Entry(control_frame, textvariable=self.min_h_var, width=6)
        h_entry.pack(side=tk.LEFT)

        ttk.Button(control_frame, text=t('apply_filter'),
                   command=self.apply_filter).pack(side=tk.LEFT, padx=10)

        ttk.Separator(control_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # Selection buttons
        ttk.Button(control_frame, text=t('select_all'),
                   command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text=t('deselect_all'),
                   command=self.deselect_all).pack(side=tk.LEFT, padx=2)

        ttk.Separator(control_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # Manual folder button
        ttk.Button(control_frame, text=t('open_folder'),
                   command=self.open_custom_folder).pack(side=tk.LEFT, padx=5)

        # Center frame - scrollable image grid
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.canvas = tk.Canvas(container, bg='#1e1e1e')
        scrollbar_y = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.canvas.yview)

        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar_y.set)

        # Match inner frame width to canvas
        self.canvas.bind('<Configure>', self._on_canvas_configure)

        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # Bottom frame - status and action buttons
        bottom_frame = ttk.Frame(self.root, padding=10)
        bottom_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(bottom_frame, text=t('searching_spotlight'))
        self.status_label.pack(side=tk.LEFT)

        # Save button
        self.save_btn = ttk.Button(bottom_frame, text=t('save_selected'),
                                    command=self.save_selected, state=tk.DISABLED)
        self.save_btn.pack(side=tk.RIGHT, padx=5)

        # Monitoring button
        monitor_available = TRAY_AVAILABLE and WATCHDOG_AVAILABLE
        self.monitor_btn = ttk.Button(
            bottom_frame,
            text=t('enable_monitoring'),
            command=self.toggle_monitoring,
            state=tk.NORMAL if monitor_available else tk.DISABLED
        )
        self.monitor_btn.pack(side=tk.RIGHT, padx=5)

        # Autostart checkbox
        self.autostart_var = tk.BooleanVar(value=self.is_autostart_enabled())
        self.autostart_cb = ttk.Checkbutton(
            bottom_frame,
            text=t('start_with_windows'),
            variable=self.autostart_var,
            command=self.toggle_autostart
        )
        self.autostart_cb.pack(side=tk.RIGHT, padx=10)

        self.progress = ttk.Progressbar(bottom_frame, mode='indeterminate', length=150)
        self.progress.pack(side=tk.RIGHT, padx=10)

        # Warn about missing optional dependencies
        if not monitor_available:
            missing = []
            if not TRAY_AVAILABLE:
                missing.append("pystray")
            if not WATCHDOG_AVAILABLE:
                missing.append("watchdog")
            if not TOAST_AVAILABLE:
                missing.append("winotify")
            self.root.after(1000, lambda: messagebox.showinfo(
                t('optional_deps_title'),
                t('optional_deps_msg', ' '.join(missing))
            ))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # === Autostart ===

    def get_exe_path(self):
        """Return the path to the executable or script"""
        if getattr(sys, 'frozen', False):
            return sys.executable
        return f'pythonw "{os.path.abspath(__file__)}"'

    def is_autostart_enabled(self):
        """Check if autostart is enabled in the registry"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.AUTOSTART_KEY, 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, self.APP_NAME)
            winreg.CloseKey(key)
            return True
        except OSError:
            return False

    def toggle_autostart(self):
        """Enable or disable Windows autostart"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.AUTOSTART_KEY, 0, winreg.KEY_SET_VALUE)

            if self.autostart_var.get():
                exe_path = self.get_exe_path()
                # Add --minimized flag to start in tray
                if getattr(sys, 'frozen', False):
                    exe_path = f'"{exe_path}" --minimized'
                winreg.SetValueEx(key, self.APP_NAME, 0, winreg.REG_SZ, exe_path)
            else:
                try:
                    winreg.DeleteValue(key, self.APP_NAME)
                except OSError:
                    pass

            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showerror(t('error'), t('autostart_error', e))
            self.autostart_var.set(not self.autostart_var.get())

    # === Monitoring & System Tray ===

    def toggle_monitoring(self):
        if self.monitoring_active:
            self.stop_monitoring()
        else:
            self.start_monitoring()

    def start_monitoring(self):
        """Start monitoring and minimize to tray"""
        paths = [p for p in self.SPOTLIGHT_PATHS.values() if p.exists()]

        if not paths:
            messagebox.showwarning(t('warning'), t('no_folders_found'))
            return

        self.monitor = SpotlightMonitor(
            paths_to_watch=paths,
            output_folder=self.output_folder,
            history_manager=self.history_manager,
            min_width=int(self.min_w_var.get() or 800),
            min_height=int(self.min_h_var.get() or 600),
            on_new_image=self.on_new_image_found
        )

        # Scan existing files first
        self.status_label.config(text=t('scanning_existing'))
        self.root.update()

        new_files = self.monitor.scan_existing()
        if new_files:
            self.show_toast(
                t('new_wallpapers_found'),
                t('n_new_images', len(new_files))
            )

        # Start real-time monitoring
        if self.monitor.start():
            self.monitoring_active = True
            self.monitor_btn.config(text=t('stop_monitoring_btn'))
            self.minimize_to_tray()
        else:
            messagebox.showerror(t('error'), t('monitor_error'))

    def stop_monitoring(self):
        if self.monitor:
            self.monitor.stop()
            self.monitor = None

        self.monitoring_active = False
        self.monitor_btn.config(text=t('enable_monitoring'))
        self.status_label.config(text=t('monitoring_stopped'))

    def on_new_image_found(self, result):
        """Callback when a new image is detected and copied"""
        w, h = result['dimensions']
        self.show_toast(
            t('new_spotlight_wallpaper'),
            t('saved_toast', result['path'].name, w, h)
        )

    def show_toast(self, title, message):
        """Show a Windows toast notification"""
        if TOAST_AVAILABLE:
            try:
                toast = Notification(
                    app_id=self.APP_NAME,
                    title=title,
                    msg=message,
                    duration="short"
                )
                toast.set_audio(audio.Default, loop=False)
                toast.show()
            except Exception as e:
                print(f"Toast error: {e}")

    def minimize_to_tray(self):
        """Minimize the window to the system tray"""
        if not TRAY_AVAILABLE:
            return

        self.root.withdraw()

        icon_image = self._create_tray_icon()

        menu = pystray.Menu(
            item(t('tray_open'), self.restore_from_tray, default=True),
            item(t('tray_stop'), self.stop_monitoring),
            item(t('tray_exit'), self.quit_app)
        )

        self.tray_icon = pystray.Icon(
            self.APP_NAME,
            icon_image,
            self.APP_NAME + t('tray_monitoring'),
            menu
        )

        tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        tray_thread.start()

    def _create_tray_icon(self):
        """Create a simple icon for the system tray"""
        size = 64
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Circular gradient background
        for i in range(size // 2, 0, -1):
            color = (30 + i * 2, 100 + i * 2, 200, 255)
            draw.ellipse([size // 2 - i, size // 2 - i, size // 2 + i, size // 2 + i], fill=color)

        # "S" letter in center
        draw.text((size // 2 - 8, size // 2 - 12), "S", fill=(255, 255, 255, 255))

        return img

    def restore_from_tray(self, icon=None, item=None):
        """Restore the window from the system tray"""
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None

        self.root.after(0, self.root.deiconify)

    def quit_app(self, icon=None, item=None):
        """Quit the application completely"""
        self.stop_monitoring()

        if self.tray_icon:
            self.tray_icon.stop()

        self.root.after(0, self.root.destroy)

    def on_close(self):
        """Handle window close: minimize to tray if monitoring, else quit"""
        if self.monitoring_active:
            self.minimize_to_tray()
        else:
            self.root.destroy()

    # === Image Loading ===

    def open_custom_folder(self):
        """Open a file dialog to select a custom image folder"""
        folder = filedialog.askdirectory(title=t('select_folder_title'))
        if folder:
            self.load_images_async(custom_path=Path(folder))

    def load_images_async(self, custom_path=None):
        self.progress.start()
        self.status_label.config(text=t('searching_images'))
        thread = threading.Thread(target=self._load_images_thread, args=(custom_path,), daemon=True)
        thread.start()

    def _load_images_thread(self, custom_path=None):
        """Load images in a background thread"""
        grouped = OrderedDict()

        paths_to_scan = {}

        if custom_path:
            paths_to_scan = {t('custom_folder'): custom_path}
        else:
            for name, path in self.SPOTLIGHT_PATHS.items():
                if path.exists():
                    paths_to_scan[name] = path

        if not paths_to_scan:
            self.root.after(0, lambda: messagebox.showwarning(
                t('warning'), t('no_folders_msg')))
            self.root.after(0, self.progress.stop)
            return

        for source_name, base_path in paths_to_scan.items():
            # IrisService uses subdirectories; ContentDeliveryManager is flat
            if 'IrisService' in str(base_path):
                for subdir in base_path.iterdir():
                    if subdir.is_dir():
                        group_name = f"IrisService/{subdir.name}"
                        images = self._scan_folder(subdir)
                        if images:
                            grouped[group_name] = images
            else:
                images = self._scan_folder(base_path)
                if images:
                    grouped[source_name] = images

        self.root.after(0, lambda: self._display_grouped_images(grouped))

    def _scan_folder(self, folder_path):
        """Scan a folder and return a list of valid image data dicts"""
        images = []

        try:
            files = list(folder_path.iterdir())
        except PermissionError:
            return images

        for file_path in files:
            if not file_path.is_file():
                continue

            try:
                with Image.open(file_path) as img:
                    width, height = img.size

                    img_copy = img.copy()
                    img_copy.thumbnail(self.THUMBNAIL_SIZE, Image.Resampling.LANCZOS)

                    if img_copy.mode != 'RGB':
                        img_copy = img_copy.convert('RGB')

                    images.append({
                        'path': file_path,
                        'size': file_path.stat().st_size,
                        'dimensions': (width, height),
                        'thumbnail': img_copy.copy(),
                        'var': None
                    })
            except Exception:
                continue

        # Sort by file size descending
        images.sort(key=lambda x: x['size'], reverse=True)
        return images

    # === Image Display ===

    def _display_grouped_images(self, grouped):
        """Display images in a grouped grid layout"""
        self.progress.stop()
        self.grouped_images = grouped
        self.thumbnail_refs = []
        self.group_frames = {}

        # Clear existing content
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        min_w = int(self.min_w_var.get() or 0)
        min_h = int(self.min_h_var.get() or 0)

        total_images = 0
        selected_count = 0
        current_row = 0

        for group_name, images in grouped.items():
            if not images:
                continue

            total_images += len(images)

            # Group header
            header_frame = ttk.Frame(self.scrollable_frame)
            header_frame.grid(row=current_row, column=0, columnspan=self.GRID_COLUMNS,
                            sticky='ew', pady=(15, 5), padx=5)

            header_text = t('group_header', group_name, len(images))
            header_label = ttk.Label(header_frame, text=header_text,
                                     font=('Segoe UI', 10, 'bold'))
            header_label.pack(side=tk.LEFT)

            # Per-group selection buttons
            ttk.Button(header_frame, text="✓", width=3,
                      command=lambda g=group_name: self.select_group(g)).pack(side=tk.RIGHT, padx=2)
            ttk.Button(header_frame, text="✗", width=3,
                      command=lambda g=group_name: self.deselect_group(g)).pack(side=tk.RIGHT, padx=2)

            sep = ttk.Separator(self.scrollable_frame, orient=tk.HORIZONTAL)
            sep.grid(row=current_row + 1, column=0, columnspan=self.GRID_COLUMNS,
                    sticky='ew', padx=5)

            current_row += 2

            # Image grid for this group
            images_frame = ttk.Frame(self.scrollable_frame)
            images_frame.grid(row=current_row, column=0, columnspan=self.GRID_COLUMNS,
                            sticky='ew', padx=5)
            self.group_frames[group_name] = images_frame

            for idx, img_data in enumerate(images):
                row = idx // self.GRID_COLUMNS
                col = idx % self.GRID_COLUMNS

                frame = ttk.Frame(images_frame, padding=3)
                frame.grid(row=row, column=col, padx=3, pady=3)

                # Thumbnail
                photo = ImageTk.PhotoImage(img_data['thumbnail'])
                self.thumbnail_refs.append(photo)

                img_label = ttk.Label(frame, image=photo)
                img_label.pack()

                # Dimensions and size info
                w, h = img_data['dimensions']
                size_kb = img_data['size'] // 1024
                info_label = ttk.Label(frame, text=f"{w}x{h} ({size_kb}KB)",
                                       font=('Consolas', 7))
                info_label.pack()

                # Selection checkbox (preselect if meets resolution filter)
                var = tk.BooleanVar()
                if w >= min_w and h >= min_h:
                    var.set(True)
                    selected_count += 1

                img_data['var'] = var
                cb = ttk.Checkbutton(frame, variable=var, command=self.update_status)
                cb.pack()

            current_row += 1

        self.save_btn.config(state=tk.NORMAL)
        self.update_status()

        self.status_label.config(
            text=t('status_format', len(grouped), total_images, selected_count, min_w, min_h))

    # === Selection ===

    def select_group(self, group_name):
        if group_name in self.grouped_images:
            for img in self.grouped_images[group_name]:
                img['var'].set(True)
        self.update_status()

    def deselect_group(self, group_name):
        if group_name in self.grouped_images:
            for img in self.grouped_images[group_name]:
                img['var'].set(False)
        self.update_status()

    def apply_filter(self):
        """Apply resolution filter and update selection"""
        try:
            min_w = int(self.min_w_var.get() or 0)
            min_h = int(self.min_h_var.get() or 0)
        except ValueError:
            messagebox.showwarning(t('warning'), t('values_must_be_numbers'))
            return

        selected_count = 0
        total = 0

        for images in self.grouped_images.values():
            for img in images:
                total += 1
                w, h = img['dimensions']
                meets = w >= min_w and h >= min_h
                img['var'].set(meets)
                if meets:
                    selected_count += 1

        self.update_status()
        self.status_label.config(
            text=t('status_format', len(self.grouped_images), total, selected_count, min_w, min_h))

    def select_all(self):
        for images in self.grouped_images.values():
            for img in images:
                img['var'].set(True)
        self.update_status()

    def deselect_all(self):
        for images in self.grouped_images.values():
            for img in images:
                img['var'].set(False)
        self.update_status()

    def update_status(self):
        selected = sum(1 for imgs in self.grouped_images.values()
                      for img in imgs if img['var'].get())
        self.save_btn.config(text=t('save_selected_n', selected))

    def _get_all_images(self):
        """Return a flat list of all images across groups"""
        return [img for imgs in self.grouped_images.values() for img in imgs]

    # === Saving ===

    def save_selected(self):
        """Copy selected images to the output folder"""
        selected = [img for img in self._get_all_images() if img['var'].get()]

        if not selected:
            messagebox.showinfo(t('info'), t('no_images_selected'))
            return

        self.output_folder.mkdir(parents=True, exist_ok=True)

        date_str = datetime.now().strftime('%Y%m%d')
        saved_count = 0
        skipped_count = 0
        errors = []

        for idx, img_data in enumerate(selected, 1):
            try:
                filepath = img_data['path']
                w, h = img_data['dimensions']

                if not is_horizontal(w, h):
                    skipped_count += 1
                    continue

                file_hash = get_file_hash(filepath)
                if self.history_manager.is_copied(file_hash):
                    skipped_count += 1
                    continue

                with Image.open(filepath) as img:
                    fmt = img.format or 'JPEG'
                    ext = 'jpg' if fmt.upper() == 'JPEG' else fmt.lower()

                new_name = f"spotlight_{idx:03d}_{date_str}.{ext}"
                dest_path = self.output_folder / new_name

                counter = 1
                while dest_path.exists():
                    new_name = f"spotlight_{idx:03d}_{date_str}_{counter}.{ext}"
                    dest_path = self.output_folder / new_name
                    counter += 1

                shutil.copy2(filepath, dest_path)

                self.history_manager.add(file_hash, filepath.name, (w, h))
                saved_count += 1

            except Exception as e:
                errors.append(f"{img_data['path'].name}: {e}")

        msg = t('saved_n_images', saved_count, self.output_folder)
        if skipped_count:
            msg += t('skipped_n', skipped_count)
        if errors:
            msg += t('errors_n', len(errors), "\n".join(errors[:5]))

        messagebox.showinfo(t('complete'), msg)

        if saved_count > 0:
            os.startfile(self.output_folder)


def main():
    start_minimized = '--minimized' in sys.argv

    root = tk.Tk()

    style = ttk.Style()
    style.theme_use('clam')

    # Dark theme
    style.configure('TFrame', background='#2b2b2b')
    style.configure('TLabel', background='#2b2b2b', foreground='#ffffff')
    style.configure('TCheckbutton', background='#2b2b2b', foreground='#ffffff')
    style.configure('TButton', padding=5)

    root.configure(bg='#2b2b2b')

    app = SpotlightSaver(root)

    # If started minimized (via autostart), begin monitoring automatically
    if start_minimized and TRAY_AVAILABLE and WATCHDOG_AVAILABLE:
        root.after(500, app.start_monitoring)

    root.mainloop()


if __name__ == '__main__':
    main()
