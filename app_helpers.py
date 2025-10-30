"""
Helper functions dan ConfigManager untuk Aplikasi Arsip Digital
"""
import os
import json
from typing import Dict


def get_appdata_path():
    """Get AppData Local path untuk menyimpan database dan export files"""
    appdata = os.getenv('LOCALAPPDATA')  # C:\Users\Username\AppData\Local
    if not appdata:
        appdata = os.path.expanduser('~\\AppData\\Local')
    
    # Buat folder khusus aplikasi
    app_folder = os.path.join(appdata, 'ArsipDigitalOwnCloud')
    if not os.path.exists(app_folder):
        os.makedirs(app_folder)
    
    return app_folder


def get_database_path():
    """Get full path untuk database.xlsx di AppData"""
    return os.path.join(get_appdata_path(), 'database.xlsx')


def get_export_path():
    """Get full path untuk file_export.xlsx di AppData"""
    return os.path.join(get_appdata_path(), 'file_export.xlsx')


def get_responsive_dimensions(base_width, base_height, screen_width, screen_height):
    """Calculate responsive window dimensions based on screen size"""
    if screen_width >= 1920:  # Large screens (4K, etc)
        width = base_width
        height = base_height
        padding = 30
        fonts = {'title': 18, 'subtitle': 11, 'normal': 10, 'small': 9}
    elif screen_width >= 1366:  # Medium screens (standard laptop)
        width = int(base_width * 0.95)
        height = int(base_height * 0.95)
        padding = 25
        fonts = {'title': 16, 'subtitle': 10, 'normal': 9, 'small': 8}
    elif screen_width >= 1024:  # Small laptop
        width = int(base_width * 0.85)
        height = int(base_height * 0.85)
        padding = 20
        fonts = {'title': 14, 'subtitle': 9, 'normal': 8, 'small': 7}
    else:  # Very small screens
        width = int(base_width * 0.75)
        height = int(base_height * 0.75)
        padding = 15
        fonts = {'title': 12, 'subtitle': 8, 'normal': 7, 'small': 6}
    
    # Ensure window doesn't exceed 85% of screen size
    max_width = int(screen_width * 0.85)
    max_height = int(screen_height * 0.85)
    width = min(width, max_width)
    height = min(height, max_height)
    
    return width, height, padding, fonts


class ConfigManager:
    """Manager untuk menyimpan dan membaca konfigurasi aplikasi"""
    
    def __init__(self):
        self.config_file = "app_config.json"
        self.default_config = {
            "default_folder": "",
            "web_server_enabled": False,
            "web_server_port": 1212
        }
        self.config = self.load_config()
    
    def load_config(self):
        """Load konfigurasi dari file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return self.default_config.copy()
        except Exception as e:
            print(f"Error loading config: {e}")
            return self.default_config.copy()
    
    def save_config(self):
        """Simpan konfigurasi ke file"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving config: {e}")
            return False
    
    def get_default_folder(self):
        """Get default folder path"""
        return self.config.get("default_folder", "")
    
    def set_default_folder(self, folder_path):
        """Set default folder path"""
        self.config["default_folder"] = folder_path
        return self.save_config()
    
    def get_web_server_enabled(self):
        """Get web server enabled status"""
        return self.config.get("web_server_enabled", False)
    
    def set_web_server_enabled(self, enabled):
        """Set web server enabled status"""
        self.config["web_server_enabled"] = enabled
        return self.save_config()
    
    def get_web_server_port(self):
        """Get web server port"""
        return self.config.get("web_server_port", 1212)
    
    def set_web_server_port(self, port):
        """Set web server port"""
        self.config["web_server_port"] = port
        return self.save_config()


# Global config manager instance
config_manager = ConfigManager()
