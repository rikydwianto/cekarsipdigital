"""
Komida Tool - Main Entry Point
Version: v1.1.6
Author: Riky Dwianto

Main menu dan entry point untuk Komida Tool.
Semua form dipisahkan ke file terpisah untuk kemudahan maintenance.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os

# Import helper functions
from app_helpers import get_export_path, get_responsive_dimensions

# Import all app forms
from app_settings import SettingsApp
from app_kk_checker import CekNoKKApp
from app_dana_checker import CekPengajuanDanaApp
from app_pdf_tools import PDFToolApp
from app_arsip import ArsipDigitalApp, ScanFolderApp, UniversalScanApp
from app_scan_files import ScanLargeFilesApp


class MainMenu:
    """Main Menu - Menu utama Komida Tool"""
    
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.create_menu_widgets()
    
    def setup_window(self):
        """Setup window utama untuk menu"""
        self.root.title("Komida Tool - Menu Utama")
        
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Get responsive dimensions
        width, height, self.padding, self.fonts = get_responsive_dimensions(
            600, 800, screen_width, screen_height
        )
        
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(True, True)
        
        # Set minimum size to prevent too small windows
        self.root.minsize(350, 450)
        
        # Center window
        self.center_window()
        
        # Set window icon (optional)
        try:
            self.root.iconbitmap("icon.ico")  # Jika ada file icon
        except:
            pass
    
    def center_window(self):
        """Center window di layar"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_menu_widgets(self):
        """Membuat widget menu utama dengan tampilan modern dan grid horizontal"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding=str(self.padding))
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Logo dan Title
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        logo_label = ttk.Label(
            header_frame, text="📁", font=("Arial", max(32, self.fonts['title'] * 2))
        )
        logo_label.grid(row=0, column=0, padx=10)

        title_label = ttk.Label(
            header_frame, text="Komida Tool",
            font=("Arial", self.fonts['title'], "bold")
        )
        title_label.grid(row=0, column=1, sticky="w")

        subtitle_label = ttk.Label(
            header_frame,
            text="Sistem Manajemen Arsip Digital",
            font=("Arial", self.fonts['subtitle']),
            foreground="gray"
        )
        subtitle_label.grid(row=1, column=0, columnspan=2)

        # Style tombol modern
        style = ttk.Style()
        button_padding = (max(12, self.padding-8), max(10, self.padding-12))
        style.configure(
            "Menu.TButton",
            padding=button_padding,
            font=("Arial", self.fonts['normal']),
            relief="flat"
        )

        # Hover effect
        style.map(
            "Menu.TButton",
            background=[("active", "#e6e6e6")],
            relief=[("pressed", "sunken")]
        )

        buttons = [
            ("📋 Cek Arsip Digital", self.open_cek_arsip),
            ("📂 Scan Folder Arsip Digital", self.open_scan_folder),
            ("🌐 Universal Scan Database", self.open_universal_scan),
            ("📊 Scan File Besar", self.open_scan_large_files),
            ("💰 Cek Pengajuan Dana", self.open_cek_pengajuan_dana),
            ("👨‍👩‍👧‍👦 Cek NO KK", self.open_cek_no_kk),
            ("📃 PDF Tool", self.open_pdf_tool),
            ("⚙️ Pengaturan", self.open_settings)
        ]

        # Frame menu dengan layout 2 kolom
        menu_frame = ttk.Frame(main_frame)
        menu_frame.grid(row=1, column=0, columnspan=2, pady=20)

        col = 0
        row = 0
        for i, (text, command) in enumerate(buttons):
            btn = ttk.Button(menu_frame, text=text, command=command, style="Menu.TButton")
            btn.grid(row=row, column=col, sticky="nsew", padx=10, pady=10, ipadx=10, ipady=10)

            # Buat grid 2 kolom
            col += 1
            if col > 1:
                col = 0
                row += 1

        # Footer
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=2, column=0, columnspan=2, pady=20)

        exit_btn = ttk.Button(footer_frame, text="Keluar", command=self.exit_app)
        exit_btn.grid(row=0, column=0, padx=20)

        version_label = ttk.Label(
            footer_frame,
            text="v1.1.6 - Developed by Riky Dwianto",
            font=("Arial", self.fonts['small']),
            foreground="gray"
        )
        version_label.grid(row=1, column=0, pady=(10, 0))

    
    def open_cek_arsip(self):
        """Membuka form Cek Arsip Digital"""
        self.root.withdraw()
        arsip_window = tk.Toplevel(self.root)
        arsip_app = ArsipDigitalApp(arsip_window, self.root)
        
        def on_arsip_close():
            arsip_window.destroy()
            self.root.deiconify()
        
        arsip_window.protocol("WM_DELETE_WINDOW", on_arsip_close)
    
    def open_scan_folder(self):
        """Membuka form Scan Folder Arsip Digital"""
        self.root.withdraw()
        scan_window = tk.Toplevel(self.root)
        scan_app = ScanFolderApp(scan_window, self.root)
        
        def on_scan_close():
            scan_window.destroy()
            self.root.deiconify()
        
        scan_window.protocol("WM_DELETE_WINDOW", on_scan_close)
    
    def open_scan_large_files(self):
        """Membuka form Scan File Besar (>10MB)"""
        self.root.withdraw()
        scan_window = tk.Toplevel(self.root)
        scan_app = ScanLargeFilesApp(scan_window, self.root)
        
        def on_scan_close():
            scan_window.destroy()
            self.root.deiconify()
        
        scan_window.protocol("WM_DELETE_WINDOW", on_scan_close)
    
    def open_cek_pengajuan_dana(self):
        """Membuka form Cek Pengajuan Dana"""
        self.root.withdraw()
        pengajuan_window = tk.Toplevel(self.root)
        pengajuan_app = CekPengajuanDanaApp(pengajuan_window, self.root)
        
        def on_pengajuan_close():
            pengajuan_window.destroy()
            self.root.deiconify()
        
        pengajuan_window.protocol("WM_DELETE_WINDOW", on_pengajuan_close)

    def open_pdf_tool(self):
        """Membuka form PDF Tool (merge/split/convert)"""
        self.root.withdraw()
        pdf_window = tk.Toplevel(self.root)
        pdf_app = PDFToolApp(pdf_window, self.root)

        def on_pdf_close():
            pdf_window.destroy()
            self.root.deiconify()

        pdf_window.protocol("WM_DELETE_WINDOW", on_pdf_close)
    
    def open_settings(self):
        """Membuka form Pengaturan"""
        self.root.withdraw()
        settings_window = tk.Toplevel(self.root)
        settings_app = SettingsApp(settings_window, self.root)
        
        def on_settings_close():
            settings_window.destroy()
            self.root.deiconify()
        
        settings_window.protocol("WM_DELETE_WINDOW", on_settings_close)
    
    def open_universal_scan(self):
        """Membuka form Universal Scan Database"""
        self.root.withdraw()
        universal_window = tk.Toplevel(self.root)
        universal_app = UniversalScanApp(universal_window, self.root)
        
        def on_universal_close():
            universal_window.destroy()
            self.root.deiconify()
        
        universal_window.protocol("WM_DELETE_WINDOW", on_universal_close)
    
    def open_cek_no_kk(self):
        """Membuka form Cek NO KK"""
        self.root.withdraw()
        nokk_window = tk.Toplevel(self.root)
        nokk_app = CekNoKKApp(nokk_window, self.root)
        
        def on_nokk_close():
            nokk_window.destroy()
            self.root.deiconify()
        
        nokk_window.protocol("WM_DELETE_WINDOW", on_nokk_close)
    
    def exit_app(self):
        """Keluar dari aplikasi"""
        if messagebox.askokcancel("Keluar", "Apakah Anda yakin ingin keluar dari aplikasi?"):
            self.root.destroy()


# ========== ENTRY POINT ==========
if __name__ == "__main__":
    # Hapus file_export.xlsx jika ada (di AppData)
    export_path = get_export_path()
    if os.path.exists(export_path):
        try:
            os.remove(export_path)
            print(f"✅ File {export_path} dihapus")
        except Exception as e:
            print(f"⚠️ Gagal menghapus {export_path}: {e}")
    
    # Jalankan aplikasi
    root = tk.Tk()
    app = MainMenu(root)
    root.mainloop()
