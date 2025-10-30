"""
Settings App - Form Pengaturan Aplikasi Arsip Digital
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
from datetime import datetime

# QR Code import (optional)
try:
    import qrcode
    from PIL import Image, ImageTk
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False
    qrcode = None
    Image = None
    ImageTk = None

from app_helpers import (
    get_appdata_path,
    get_database_path,
    get_export_path,
    get_responsive_dimensions,
    config_manager
)
from web_server import get_web_server_manager

# Global web server manager instance
web_server_manager = get_web_server_manager()


class SettingsApp:
    """Form untuk Pengaturan Aplikasi"""
    
    def __init__(self, root, parent_window=None):
        self.root = root
        self.parent_window = parent_window
        
        self.setup_window()
        self.create_widgets()
    
    def setup_window(self):
        """Setup window pengaturan"""
        self.root.title("Pengaturan - Aplikasi Arsip Digital")
        
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Get responsive dimensions
        width, height, self.padding, self.fonts = get_responsive_dimensions(
            700, 800, screen_width, screen_height
        )
        
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(True, True)
        
        # Set minimum size to prevent too small windows
        self.root.minsize(480, 500)
        
        # Center window
        self.center_window()
    
    def center_window(self):
        """Center window di layar"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Membuat widget untuk pengaturan dengan scrollable canvas"""
        # Main container frame
        container = ttk.Frame(self.root)
        container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        
        # Scrollable frame inside canvas
        scrollable_frame = ttk.Frame(canvas)
        
        # Configure canvas scroll region
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Create window in canvas
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Grid layout
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Main frame dengan padding yang responsif
        screen_width = self.root.winfo_screenwidth()
        padding_size = 30 if screen_width >= 1366 else 20 if screen_width >= 1024 else 15
        
        main_frame = ttk.Frame(scrollable_frame, padding=str(padding_size))
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure main_frame
        scrollable_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Title dengan ukuran font responsif
        title_size = 16 if screen_width >= 1366 else 14 if screen_width >= 1024 else 12
        title_label = ttk.Label(
            main_frame, 
            text="⚙️ PENGATURAN", 
            font=("Arial", title_size, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # Subtitle
        subtitle_size = 10 if screen_width >= 1366 else 9 if screen_width >= 1024 else 8
        subtitle_label = ttk.Label(
            main_frame, 
            text="Konfigurasi default untuk aplikasi",
            font=("Arial", subtitle_size),
            foreground="gray"
        )
        subtitle_label.grid(row=1, column=0, pady=(0, 20))
        
        # Frame untuk Default Folder dengan padding responsif
        folder_padding = 15 if screen_width >= 1366 else 12 if screen_width >= 1024 else 10
        folder_frame = ttk.LabelFrame(main_frame, text="Default Folder Arsip Digital", padding=str(folder_padding))
        folder_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        folder_frame.columnconfigure(0, weight=1)
        
        # Info label dengan wraplength responsif
        wrap_length = 500 if screen_width >= 1366 else 400 if screen_width >= 1024 else 350
        info_label = ttk.Label(
            folder_frame,
            text="Folder ini akan digunakan sebagai default saat membuka form lain",
            font=("Arial", 9),
            foreground="gray",
            wraplength=wrap_length
        )
        info_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Current default folder display
        current_default = config_manager.get_default_folder()
        display_text = current_default if current_default else "Belum ada folder default yang dipilih"
        
        self.folder_var = tk.StringVar(value=display_text)
        folder_path_label = ttk.Label(
            folder_frame, 
            textvariable=self.folder_var,
            foreground="blue" if current_default else "gray",
            wraplength=wrap_length,
            font=("Arial", 9, "bold" if current_default else "normal")
        )
        folder_path_label.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Buttons frame
        btn_frame = ttk.Frame(folder_frame)
        btn_frame.grid(row=2, column=0)
        
        # Browse button
        browse_btn = ttk.Button(
            btn_frame, 
            text="📂 Pilih Folder Default", 
            command=self.browse_default_folder
        )
        browse_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Clear button
        clear_btn = ttk.Button(
            btn_frame, 
            text="🗑️ Hapus Default", 
            command=self.clear_default_folder
        )
        clear_btn.grid(row=0, column=1, padx=(10, 0))
        
        # ===== WEB SERVER SECTION =====
        # Frame untuk Web Server dengan padding responsif
        webserver_frame = ttk.LabelFrame(main_frame, text="🌐 Web Server", padding=str(folder_padding))
        webserver_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        webserver_frame.columnconfigure(0, weight=1)
        
        # Info label
        webserver_info_label = ttk.Label(
            webserver_frame,
            text="Aktifkan web server untuk akses file arsip melalui browser",
            font=("Arial", 9),
            foreground="gray",
            wraplength=wrap_length
        )
        webserver_info_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Server status frame
        status_frame = ttk.Frame(webserver_frame)
        status_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(status_frame, text="Status:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.server_status_var = tk.StringVar(value="⚫ Tidak Aktif")
        ttk.Label(
            status_frame, 
            textvariable=self.server_status_var,
            font=("Arial", 9)
        ).grid(row=0, column=1, sticky=tk.W)
        
        # Local IP frame
        ip_frame = ttk.Frame(webserver_frame)
        ip_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(ip_frame, text="IP Lokal:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        local_ip = web_server_manager.get_local_ip()
        self.local_ip_var = tk.StringVar(value=local_ip)
        ttk.Label(
            ip_frame, 
            textvariable=self.local_ip_var,
            font=("Arial", 9),
            foreground="blue"
        ).grid(row=0, column=1, sticky=tk.W)
        
        # Port frame
        port_frame = ttk.Frame(webserver_frame)
        port_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(port_frame, text="Port:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.port_var = tk.StringVar(value=str(config_manager.get_web_server_port()))
        port_entry = ttk.Entry(port_frame, textvariable=self.port_var, width=10)
        port_entry.grid(row=0, column=1, sticky=tk.W)
        
        # URL frame
        url_frame = ttk.Frame(webserver_frame)
        url_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        url_frame.columnconfigure(0, weight=1)
        
        ttk.Label(url_frame, text="URL Akses:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.url_local_var = tk.StringVar(value="http://localhost:1212")
        url_local_label = ttk.Label(
            url_frame, 
            textvariable=self.url_local_var,
            font=("Arial", 8),
            foreground="blue",
            cursor="hand2"
        )
        url_local_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 0))
        
        self.url_network_var = tk.StringVar(value=f"http://{local_ip}:1212")
        url_network_label = ttk.Label(
            url_frame, 
            textvariable=self.url_network_var,
            font=("Arial", 8),
            foreground="blue",
            cursor="hand2"
        )
        url_network_label.grid(row=2, column=0, sticky=tk.W, padx=(0, 0))
        
        # QR Code frame
        qr_frame = ttk.LabelFrame(webserver_frame, text="📱 QR Code untuk HP", padding="10")
        qr_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(10, 10))
        qr_frame.columnconfigure(0, weight=1)
        
        # QR Code info
        qr_info_label = ttk.Label(
            qr_frame,
            text="Scan QR code ini dari HP (pastikan 1 jaringan WiFi)",
            font=("Arial", 8),
            foreground="gray"
        )
        qr_info_label.grid(row=0, column=0, pady=(0, 10))
        
        # QR Code container
        self.qr_label = ttk.Label(qr_frame, text="QR Code akan muncul saat server aktif", foreground="gray")
        self.qr_label.grid(row=1, column=0)
        
        # Refresh QR button
        refresh_qr_btn = ttk.Button(
            qr_frame,
            text="🔄 Refresh QR Code",
            command=self.refresh_qr_code
        )
        refresh_qr_btn.grid(row=2, column=0, pady=(10, 0))
        
        # Server control buttons
        server_btn_frame = ttk.Frame(webserver_frame)
        server_btn_frame.grid(row=6, column=0, pady=(10, 0))
        
        self.start_server_btn = ttk.Button(
            server_btn_frame, 
            text="▶️ Start Server", 
            command=self.start_web_server
        )
        self.start_server_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.stop_server_btn = ttk.Button(
            server_btn_frame, 
            text="⏹️ Stop Server", 
            command=self.stop_web_server,
            state=tk.DISABLED
        )
        self.stop_server_btn.grid(row=0, column=1, padx=(10, 0))
        
        # Update server status saat load
        self.update_server_status()
        
        # Status label
        self.status_var = tk.StringVar(value="")
        status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            font=("Arial", 9),
            foreground="green"
        )
        status_label.grid(row=4, column=0, pady=(0, 15))
        
        # Footer buttons
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=5, column=0, pady=(10, 0))
        
        # Back button
        if self.parent_window:
            back_btn = ttk.Button(
                footer_frame, 
                text="⬅️ Kembali ke Menu", 
                command=self.back_to_menu
            )
            back_btn.grid(row=0, column=0)
        
        # Update canvas width to match window
        def _configure_canvas(event):
            canvas.itemconfig(canvas.find_withtag("all")[0], width=event.width)
        
        canvas.bind("<Configure>", _configure_canvas)
    
    def update_server_status(self):
        """Update status web server di UI"""
        info = web_server_manager.get_server_info()
        
        if info["status"] == "Running":
            self.server_status_var.set("🟢 Aktif")
            self.start_server_btn.config(state=tk.DISABLED)
            self.stop_server_btn.config(state=tk.NORMAL)
            # Generate QR Code saat server aktif
            self.generate_qr_code(info["url_network"])
        else:
            self.server_status_var.set("⚫ Tidak Aktif")
            self.start_server_btn.config(state=tk.NORMAL)
            self.stop_server_btn.config(state=tk.DISABLED)
            # Clear QR Code saat server mati
            self.qr_label.config(image='', text="QR Code akan muncul saat server aktif", foreground="gray")
        
        self.url_local_var.set(info["url_local"])
        self.url_network_var.set(info["url_network"])
    
    def generate_qr_code(self, url):
        """Generate QR code untuk URL"""
        if not QR_AVAILABLE:
            self.qr_label.config(image='', text="QR Code library tidak tersedia\nInstall: pip install qrcode[pil]", foreground="orange")
            return
        
        try:
            # Create QR code
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(url)
            qr.make(fit=True)
            
            # Create image
            img = qr.make_image(fill_color="black", back_color="white")
            
            # Resize untuk fit di window (200x200)
            img = img.resize((150, 150), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(img)
            
            # Store reference to prevent garbage collection
            self.qr_photo = photo
            
            # Update label
            self.qr_label.config(image=photo, text="")
            
        except Exception as e:
            self.qr_label.config(image='', text=f"Error: {str(e)}", foreground="red")
    
    def refresh_qr_code(self):
        """Refresh QR code"""
        if not QR_AVAILABLE:
            messagebox.showwarning("Warning", "QR Code library tidak tersedia!\n\nInstall dengan: pip install qrcode[pil]")
            return
        
        info = web_server_manager.get_server_info()
        if info["status"] == "Running":
            self.generate_qr_code(info["url_network"])
            messagebox.showinfo("QR Code", "QR Code berhasil di-refresh!")
        else:
            messagebox.showwarning("Warning", "Server belum aktif!\nStart server terlebih dahulu.")

    
    def start_web_server(self):
        """Start web server"""
        # cek dulu apakah ada file database.xlsx di AppData
        database_path = get_database_path()
        if not os.path.exists(database_path):
            messagebox.showerror("Error", f"File database.xlsx tidak ditemukan di:\n{database_path}\n\nSilakan scan folder arsip digital terlebih dahulu lalu pilih simpan dan singkron.")
            return
        else:
            try:
                port = int(self.port_var.get())
                if port < 1024 or port > 65535:
                    messagebox.showerror("Error", "Port harus antara 1024 dan 65535")
                    return
            except ValueError:
                messagebox.showerror("Error", "Port harus berupa angka")
                return
            
            # Save port to config
            config_manager.set_web_server_port(port)
            
            success, message = web_server_manager.start_server(port)
            
            if success:
                self.status_var.set("✅ Web server berhasil dijalankan!")
                messagebox.showinfo("Server Started", message)
                self.update_server_status()
                
                # Clear status after 3 seconds
                self.root.after(3000, lambda: self.status_var.set(""))
            else:
                messagebox.showerror("Error", f"Gagal start server:\n{message}")
    
    def stop_web_server(self):
        """Stop web server"""
        success, message = web_server_manager.stop_server()
        
        if success:
            self.status_var.set("✅ Web server berhasil dihentikan!")
            messagebox.showinfo("Server Stopped", message)
            self.update_server_status()
            
            # Clear QR code
            self.qr_label.config(image='', text="QR Code akan muncul saat server aktif", foreground="gray")
            
            # Clear status after 3 seconds
            self.root.after(3000, lambda: self.status_var.set(""))
        else:
            messagebox.showerror("Error", f"Gagal stop server:\n{message}")
    
    def browse_default_folder(self):
        """Pilih folder default dan auto-generate database"""
        current_default = config_manager.get_default_folder()
        initial_dir = current_default if current_default and os.path.exists(current_default) else os.getcwd()
        
        folder_path = filedialog.askdirectory(
            title="Pilih Folder Default Arsip Digital",
            initialdir=initial_dir
        )
        
        if folder_path:
            if os.path.exists(folder_path) and os.path.isdir(folder_path):
                # Save to config
                if config_manager.set_default_folder(folder_path):
                    self.folder_var.set(folder_path)
                    self.status_var.set("✅ Folder default berhasil disimpan!")
                    
                    # Clear status after 3 seconds
                    self.root.after(3000, lambda: self.status_var.set(""))
                else:
                    messagebox.showerror("Error", "Gagal menyimpan konfigurasi!")
            else:
                messagebox.showerror("Error", "Folder yang dipilih tidak valid!")
    
    def clear_default_folder(self):
        """Hapus folder default"""
        if config_manager.get_default_folder():
            result = messagebox.askyesno(
                "Konfirmasi",
                "Apakah Anda yakin ingin menghapus folder default?\n\n"
                "Anda perlu memilih folder secara manual di setiap form."
            )
            
            if result:
                if config_manager.set_default_folder(""):
                    self.folder_var.set("Belum ada folder default yang dipilih")
                    self.status_var.set("✅ Folder default berhasil dihapus!")
                    
                    # Clear status after 3 seconds
                    self.root.after(3000, lambda: self.status_var.set(""))
                    
                    messagebox.showinfo("Berhasil", "Folder default berhasil dihapus!")
                else:
                    messagebox.showerror("Error", "Gagal menghapus konfigurasi!")
        else:
            messagebox.showinfo("Info", "Tidak ada folder default yang tersimpan.")
    
    def back_to_menu(self):
        """Kembali ke menu utama"""
        if self.parent_window:
            self.root.destroy()
            self.parent_window.deiconify()
