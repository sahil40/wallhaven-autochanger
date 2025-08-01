import sys
import os
import json
import random
import ctypes
import requests
from PIL import Image
from io import BytesIO
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QIcon, QCloseEvent
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QHBoxLayout, QCheckBox, QSpinBox, QFileDialog, QMessageBox, QComboBox,
    QSystemTrayIcon, QMenu
)

CONFIG_FILE = "config.json"
ICON_PATH = os.path.join(os.path.dirname(__file__), "icon.png")

DEFAULT_CONFIG = {
    "api_key": "",
    "query": "",
    "categories": "111",
    "purity": "100",
    "resolutions": "1920x1080",
    "ratios": "16x9",
    "sorting": "random",
    "order": "desc",
    "change_interval": 60,
    "download_dir": os.path.join(os.path.expanduser("~"), "Pictures", "Wallpapers"),
    "topRange": "1M",
    "start_minimized": False,
    "launch_on_boot": False,
}

class WallpaperApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WallHaven AutoChanger")
        
        # Load icon with fallback
        if os.path.exists(ICON_PATH):
            self.app_icon = QIcon(ICON_PATH)
        else:
            # Fallback to a default icon if icon.png doesn't exist
            self.app_icon = QIcon.fromTheme("applications-graphics")
            if self.app_icon.isNull():
                # Create a simple default icon if theme icon is also null
                self.app_icon = QIcon()
        
        self.setWindowIcon(self.app_icon)
        self.config = self.load_config()
        self.setup_ui()
        self.setup_tray()
        self.start_timer()
        
        # Check tray status after setup
        self.check_tray_status()

    def load_config(self):
        config = DEFAULT_CONFIG.copy()
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    saved = json.load(f)
                    config.update(saved)
            except Exception as e:
                print("Error reading config file:", e)
        config["categories"] = config.get("categories", "111").ljust(3, "0")[:3]
        config["purity"] = config.get("purity", "100").ljust(3, "0")[:3]
        return config

    def save_config(self):
        with open(CONFIG_FILE, "w") as f:
            json.dump(self.config, f, indent=4)
        self.status_label.setText("Settings saved successfully.")

    def setup_ui(self):
        layout = QVBoxLayout()

        self.api_key_input = QLineEdit(self.config["api_key"])
        layout.addWidget(QLabel("Wallhaven API Key:"))
        layout.addWidget(self.api_key_input)

        self.query_input = QLineEdit(self.config["query"])
        layout.addWidget(QLabel("Search Query:"))
        layout.addWidget(self.query_input)

        cat = self.config["categories"]
        pur = self.config["purity"]

        self.general = QCheckBox("General")
        self.anime = QCheckBox("Anime")
        self.people = QCheckBox("People")
        self.general.setChecked(cat[0] == "1")
        self.anime.setChecked(cat[1] == "1")
        self.people.setChecked(cat[2] == "1")
        layout.addWidget(QLabel("Categories:"))
        layout.addLayout(self.wrap_hbox([self.general, self.anime, self.people]))

        self.sfw = QCheckBox("SFW")
        self.sketchy = QCheckBox("Sketchy")
        self.nsfw = QCheckBox("NSFW")
        self.sfw.setChecked(pur[0] == "1")
        self.sketchy.setChecked(pur[1] == "1")
        self.nsfw.setChecked(pur[2] == "1")
        layout.addWidget(QLabel("Purity:"))
        layout.addLayout(self.wrap_hbox([self.sfw, self.sketchy, self.nsfw]))

        self.res_input = QLineEdit(self.config["resolutions"])
        layout.addWidget(QLabel("Min Resolution (e.g., 1920x1080):"))
        layout.addWidget(self.res_input)

        self.ratio_input = QLineEdit(self.config["ratios"])
        layout.addWidget(QLabel("Aspect Ratio (e.g., 16x9):"))
        layout.addWidget(self.ratio_input)

        self.sort_input = QComboBox()
        self.sort_input.addItems(["random", "toplist", "date_added"])
        self.sort_input.setCurrentText(self.config["sorting"])
        layout.addWidget(QLabel("Sorting:"))
        layout.addWidget(self.sort_input)

        self.toprange_input = QComboBox()
        self.toprange_input.addItems(["1d", "3d", "1w", "1M", "3M", "6M", "1y"])
        self.toprange_input.setCurrentText(self.config.get("topRange", "1M"))
        layout.addWidget(QLabel("Toplist Range (only if sorting = toplist):"))
        layout.addWidget(self.toprange_input)

        self.order_input = QComboBox()
        self.order_input.addItems(["desc", "asc"])
        self.order_input.setCurrentText(self.config["order"])
        layout.addWidget(QLabel("Order:"))
        layout.addWidget(self.order_input)

        self.interval_input = QSpinBox()
        self.interval_input.setValue(self.config["change_interval"])
        self.interval_input.setSuffix(" min")
        layout.addWidget(QLabel("Auto-change interval (min):"))
        layout.addWidget(self.interval_input)

        self.start_minimized = QCheckBox("Start Minimized")
        self.start_minimized.setChecked(self.config.get("start_minimized", False))
        layout.addWidget(self.start_minimized)

        self.launch_on_boot = QCheckBox("Launch on System Startup")
        self.launch_on_boot.setChecked(self.config.get("launch_on_boot", False))
        layout.addWidget(self.launch_on_boot)


        self.dir_btn = QPushButton("Select Download Directory")
        self.dir_btn.clicked.connect(self.select_dir)
        layout.addWidget(self.dir_btn)

        save_btn = QPushButton("Save Settings")
        save_btn.clicked.connect(self.save_and_update)
        layout.addWidget(save_btn)

        change_btn = QPushButton("Change Wallpaper Now")
        change_btn.clicked.connect(self.change_wallpaper_now)
        layout.addWidget(change_btn)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def setup_tray(self):
        # Check if system tray is available
        if not QSystemTrayIcon.isSystemTrayAvailable():
            print("System tray is not available")
            return
            
        # Create tray icon with proper error handling
        try:
            self.tray = QSystemTrayIcon(self.app_icon, self)
            
            self.tray_menu = QMenu()
            self.tray_menu.addAction("Change Now", self.change_wallpaper_now)
            self.tray_menu.addAction("Show", self.showNormal)
            self.tray_menu.addAction("Quit", QApplication.instance().quit)
            self.tray.setContextMenu(self.tray_menu)
            self.tray.setToolTip("WallHaven AutoChanger")
            # Connect tray icon click to show the app
            self.tray.activated.connect(self.tray_icon_activated)
            
            # Show the tray icon with retry
            if not self.tray.show():
                # Try again after a short delay
                QTimer.singleShot(100, self.tray.show)
                print("Tray icon created, retrying to show...")
            else:
                print("Tray icon created and shown successfully")
                
        except Exception as e:
            print(f"Error setting up tray icon: {e}")

    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.showNormal()
            self.raise_()
            self.activateWindow()

    def start_timer(self):
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.change_wallpaper_now)
        self.timer.start(self.config["change_interval"] * 60000)

    def wrap_hbox(self, widgets):
        hbox = QHBoxLayout()
        for w in widgets:
            hbox.addWidget(w)
        return hbox

    def select_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder", self.config["download_dir"])
        if folder:
            self.config["download_dir"] = folder

    def save_and_update(self):
        self.config["api_key"] = self.api_key_input.text()
        self.config["query"] = self.query_input.text()
        self.config["categories"] = f"{int(self.general.isChecked())}{int(self.anime.isChecked())}{int(self.people.isChecked())}"
        self.config["purity"] = f"{int(self.sfw.isChecked())}{int(self.sketchy.isChecked())}{int(self.nsfw.isChecked())}"
        self.config["resolutions"] = self.res_input.text()
        self.config["ratios"] = self.ratio_input.text()
        self.config["sorting"] = self.sort_input.currentText()
        self.config["topRange"] = self.toprange_input.currentText()
        self.config["order"] = self.order_input.currentText()
        self.config["change_interval"] = self.interval_input.value()
        self.config["start_minimized"] = self.start_minimized.isChecked()
        self.config["launch_on_boot"] = self.launch_on_boot.isChecked()
        self.save_config()
        self.timer.start(self.config["change_interval"] * 60000)
        self.update_autostart(self.config["launch_on_boot"])

    def change_wallpaper_now(self):
        self.save_and_update()
        cfg = self.config
        headers = {"X-API-Key": cfg["api_key"]}
        params = {
            "q": cfg["query"],
            "categories": cfg["categories"],
            "purity": cfg["purity"],
            "resolutions": cfg["resolutions"],
            "ratios": cfg["ratios"],
            "sorting": cfg["sorting"],
            "order": cfg["order"],
            "at_least": cfg["resolutions"],
            "page": random.randint(1, 10),
            "per_page": 24
        }
        if cfg["sorting"] == "toplist":
            params["topRange"] = cfg.get("topRange", "1M")
        try:
            res = requests.get("https://wallhaven.cc/api/v1/search", headers=headers, params=params)
            res.raise_for_status()
            data = res.json()
            wallpapers = data.get("data", [])
            if not wallpapers:
                self.status_label.setText("No wallpapers found.")
                return
            chosen = random.choice(wallpapers)
            img_url = chosen["path"]
            file_ext = img_url.split(".")[-1]
            filename = os.path.join(cfg["download_dir"], f"wallhaven_{chosen['id']}.{file_ext}")
            img_data = requests.get(img_url).content
            with open(filename, "wb") as f:
                f.write(img_data)
            ctypes.windll.user32.SystemParametersInfoW(20, 0, filename, 3)
            self.status_label.setText(f"Wallpaper changed: {filename}")
        except Exception as e:
            self.status_label.setText(f"Error: {e}")

    def closeEvent(self, event: QCloseEvent):
        event.ignore()
        self.hide()
        
        # Ensure tray icon is visible
        if hasattr(self, 'tray') and self.tray:
            if self.tray.isVisible():
                self.tray.showMessage("WallHaven AutoChanger", "App minimized to tray.", QSystemTrayIcon.Information)
            else:
                print("Tray icon is not visible, trying to show...")
                # Try to show the tray icon
                QTimer.singleShot(100, self.ensure_tray_visible)
                self.tray.showMessage("WallHaven AutoChanger", "App minimized to tray.", QSystemTrayIcon.Information)
        else:
            print("Tray icon not available")

    def update_autostart(self, enable):
        try:
            startup_dir = os.path.join(os.getenv('APPDATA'), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
            script_path = sys.argv[0]
            shortcut_path = os.path.join(startup_dir, "WallHaven AutoChanger.lnk")

            if enable:
                try:
                    import pythoncom
                    from win32com.client import Dispatch
                    shell = Dispatch('WScript.Shell')
                    shortcut = shell.CreateShortCut(shortcut_path)
                    shortcut.Targetpath = sys.executable
                    shortcut.Arguments = f'"{script_path}"'
                    shortcut.WorkingDirectory = os.path.dirname(script_path)
                    shortcut.IconLocation = script_path
                    shortcut.save()
                    print("Autostart shortcut created successfully")
                except ImportError:
                    print("pywin32 not installed. Autostart functionality disabled.")
                    print("Install with: pip install pywin32")
                except Exception as e:
                    print(f"Error creating autostart shortcut: {e}")
            else:
                if os.path.exists(shortcut_path):
                    try:
                        os.remove(shortcut_path)
                        print("Autostart shortcut removed successfully")
                    except Exception as e:
                        print(f"Error removing autostart shortcut: {e}")
        except Exception as e:
            print(f"Autostart error: {e}")

    def check_tray_status(self):
        """Check if tray icon is working properly"""
        if hasattr(self, 'tray') and self.tray:
            if self.tray.isVisible():
                print("✅ Tray icon is visible and working")
            else:
                print("❌ Tray icon is not visible")
                # Try to show it again
                QTimer.singleShot(500, self.ensure_tray_visible)
        else:
            print("❌ Tray icon not created")

    def ensure_tray_visible(self):
        """Ensure tray icon is visible"""
        if hasattr(self, 'tray') and self.tray:
            if not self.tray.isVisible():
                self.tray.show()
                print("Retrying to show tray icon...")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WallpaperApp()
    if window.config.get("start_minimized"):
        window.hide()
    else:
        window.show()
    sys.exit(app.exec_())
