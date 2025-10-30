"""
Scan Large Files App - Form untuk scan file besar (>10MB)
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
from datetime import datetime

from app_helpers import (
    get_appdata_path,
    get_database_path,
    get_export_path,
    get_responsive_dimensions
)

class ScanLargeFilesApp:
    """Form untuk Scan File Besar (>10MB) dari Folder Arsip Digital Owncloud"""
    
    def __init__(self, root, parent_window=None):
        self.root = root
        self.parent_window = parent_window
        
        # File yang akan diabaikan (owncloud sync files)
        self.ignored_files = {
            '.owncloudsync.log',
            '.owncloudsync.log.1',
            '.sync_journal.db',
            '.sync_journal.db-wal'
        }
        
        # Format dokumen yang umum/diizinkan (untuk mode format)
        self.allowed_extensions = {
            # Office Documents
            '.doc', '.docx', '.xls', '.xlsx', '.xlsm', '.ppt', '.pptx',
            '.odt', '.ods', '.odp',  # OpenOffice/LibreOffice
            # PDF
            '.pdf',
            # Text
            '.txt', '.rtf', '.csv',
            # Images
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', 
        }
        
        # Default minimum size in MB
        self.min_size_mb = 10
        
        self.setup_window()
        self.create_widgets()
        
        # Variables
        self.selected_folder = ""
        self.scan_results = []
    
    def setup_window(self):
        """Setup window utama aplikasi"""
        self.root.title("Scan File Besar (>10MB) - Arsip Digital Owncloud")
        
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Get responsive dimensions
        width, height, self.padding, self.fonts = get_responsive_dimensions(
            900, 900, screen_width, screen_height
        )
        
        self.root.geometry(f"{width}x{height}")
        self.root.resizable(True, True)
        
        # Set minimum size
        self.root.minsize(650, 600)
        
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
        """Membuat semua widget GUI dengan scrollable canvas"""
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
        
        # Main frame dengan padding responsif
        main_frame = ttk.Frame(scrollable_frame, padding=str(self.padding))
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure main_frame
        scrollable_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)  # Result frame expandable
        
        # Title dengan font responsif
        title_label = ttk.Label(
            main_frame, 
            text="🔍 SCAN FILE BESAR & FORMAT NON-DOKUMEN", 
            font=("Arial", self.fonts['title'], "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # Subtitle
        subtitle_label = ttk.Label(
            main_frame, 
            text="Temukan file berukuran besar atau file dengan format tidak umum",
            font=("Arial", self.fonts['subtitle'])
        )
        subtitle_label.grid(row=1, column=0, pady=(0, 20))
        
        # Frame untuk mode scan dengan padding responsif
        frame_padding = max(10, self.padding - 5)
        mode_frame = ttk.LabelFrame(main_frame, text="Mode Scan", padding=str(frame_padding))
        mode_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        mode_frame.columnconfigure(0, weight=1)
        
        # Radio buttons untuk memilih mode
        self.scan_mode = tk.StringVar(value="size")
        
        mode_size_rb = ttk.Radiobutton(
            mode_frame,
            text="📏 File Besar (berdasarkan ukuran minimum)",
            variable=self.scan_mode,
            value="size",
            command=self.on_mode_change
        )
        mode_size_rb.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        mode_format_rb = ttk.Radiobutton(
            mode_frame,
            text="📄 Format Non-Dokumen (selain Office, Image, PDF, TXT)",
            variable=self.scan_mode,
            value="format",
            command=self.on_mode_change
        )
        mode_format_rb.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        
        # Frame untuk ukuran minimum (hanya aktif jika mode = size)
        self.size_frame = ttk.LabelFrame(main_frame, text="Pengaturan Ukuran", padding=str(frame_padding))
        self.size_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        self.size_frame.columnconfigure(1, weight=1)
        
        # Label dan input untuk ukuran minimum
        ttk.Label(self.size_frame, text="Ukuran Minimum (MB):").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.size_var = tk.StringVar(value="10")
        self.size_entry = ttk.Entry(self.size_frame, textvariable=self.size_var, width=10)
        self.size_entry.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(self.size_frame, text="(File yang lebih kecil akan diabaikan)", 
                 font=("Arial", self.fonts['small']), foreground="gray").grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        
        # Frame untuk folder selection
        folder_frame = ttk.LabelFrame(main_frame, text="Pilih Folder Arsip Digital Owncloud", padding=str(frame_padding))
        folder_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        folder_frame.columnconfigure(0, weight=1)
        
        # Folder path display dengan wraplength responsif
        wrap_length = max(500, int(self.root.winfo_screenwidth() * 0.7))
        self.folder_var = tk.StringVar(value="Belum ada folder yang dipilih...")
        folder_path_label = ttk.Label(
            folder_frame, 
            textvariable=self.folder_var,
            foreground="gray",
            wraplength=wrap_length
        )
        folder_path_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Folder browse button
        folder_btn = ttk.Button(
            folder_frame, 
            text="📂 Browse Folder", 
            command=self.browse_folder
        )
        folder_btn.grid(row=1, column=0)
        
        # Frame untuk hasil scan
        result_frame = ttk.LabelFrame(main_frame, text="Hasil Scan", padding="15")
        result_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # Treeview untuk menampilkan hasil
        columns = ("No", "Nama File", "Ekstensi", "Ukuran", "Path")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)
        
        # Define headings
        self.tree.heading("No", text="No")
        self.tree.heading("Nama File", text="Nama File")
        self.tree.heading("Ekstensi", text="Ekstensi")
        self.tree.heading("Ukuran", text="Ukuran")
        self.tree.heading("Path", text="Path Lengkap")
        
        # Define column widths
        self.tree.column("No", width=50, anchor=tk.CENTER)
        self.tree.column("Nama File", width=200)
        self.tree.column("Ekstensi", width=80, anchor=tk.CENTER)
        self.tree.column("Ukuran", width=100, anchor=tk.E)
        self.tree.column("Path", width=400)
        
        # Scrollbars
        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Info label
        self.info_var = tk.StringVar(value="Pilih folder untuk memulai scan")
        info_label = ttk.Label(main_frame, textvariable=self.info_var, font=("Arial", 9), foreground="blue")
        info_label.grid(row=6, column=0, pady=(0, 10))
        
        # Action buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, pady=(10, 0))
        
        # Scan button
        self.scan_btn = ttk.Button(
            button_frame, 
            text="🔍 Mulai Scan", 
            command=self.start_scan,
            state="disabled"
        )
        self.scan_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Export button
        self.export_btn = ttk.Button(
            button_frame, 
            text="📊 Export ke Excel", 
            command=self.export_to_excel,
            state="disabled"
        )
        self.export_btn.grid(row=0, column=1, padx=(10, 10))
        
        # Clear button
        self.clear_btn = ttk.Button(
            button_frame, 
            text="🗑️ Clear", 
            command=self.clear_results,
            state="disabled"
        )
        self.clear_btn.grid(row=0, column=2, padx=(10, 10))
        
        # Back button
        if self.parent_window:
            back_btn = ttk.Button(
                button_frame, 
                text="⬅️ Kembali", 
                command=self.back_to_menu
            )
            back_btn.grid(row=0, column=3, padx=(10, 0))
        
        # Update canvas width to match window
        def _configure_canvas(event):
            canvas.itemconfig(canvas.find_withtag("all")[0], width=event.width)
        
        canvas.bind("<Configure>", _configure_canvas)
    
    def on_mode_change(self):
        """Handle perubahan mode scan"""
        mode = self.scan_mode.get()
        
        if mode == "size":
            # Enable size input
            self.size_entry.config(state="normal")
            self.info_var.set("Mode: File Besar - Pilih folder untuk memulai scan")
        else:  # format
            # Disable size input
            self.size_entry.config(state="disabled")
            self.info_var.set("Mode: Format Non-Dokumen - Pilih folder untuk memulai scan")
    
    def browse_folder(self):
        """Fungsi untuk memilih folder"""
        # Gunakan default folder jika ada
        default_folder = config_manager.get_default_folder()
        initial_dir = default_folder if default_folder and os.path.exists(default_folder) else os.getcwd()
        
        folder_path = filedialog.askdirectory(
            title="Pilih Folder Arsip Digital Owncloud",
            initialdir=initial_dir
        )
        
        if folder_path:
            if os.path.exists(folder_path) and os.path.isdir(folder_path):
                self.selected_folder = folder_path
                self.folder_var.set(folder_path)
                self.scan_btn.config(state="normal")
                self.info_var.set(f"Folder dipilih: {os.path.basename(folder_path)}")
            else:
                messagebox.showerror("Error", "Folder yang dipilih tidak valid!")
                self.selected_folder = ""
    
    def start_scan(self):
        """Mulai scan file besar atau format non-dokumen"""
        if not self.selected_folder:
            messagebox.showwarning("Peringatan", "Silakan pilih folder terlebih dahulu!")
            return
        
        mode = self.scan_mode.get()
        
        # Validasi input ukuran minimum jika mode = size
        if mode == "size":
            try:
                self.min_size_mb = float(self.size_var.get())
                if self.min_size_mb <= 0:
                    messagebox.showerror("Error", "Ukuran minimum harus lebih dari 0 MB!")
                    return
            except ValueError:
                messagebox.showerror("Error", "Ukuran minimum harus berupa angka!")
                return
        
        # Clear previous results
        self.clear_results()
        
        # Progress dialog
        if mode == "size":
            progress_msg = f"Scanning file lebih dari {self.min_size_mb} MB..."
        else:
            progress_msg = "Scanning file dengan format non-dokumen..."
        
        progress_window = self.show_progress_dialog(progress_msg)
        
        try:
            # Scan folder
            self.scan_results = []
            self.scan_folder_recursive(self.selected_folder)
            
            progress_window.destroy()
            
            # Sort by size (descending)
            self.scan_results.sort(key=lambda x: x['size_bytes'], reverse=True)
            
            # Display results
            self.display_results()
            
            # Update info
            total_files = len(self.scan_results)
            total_size = sum(f['size_bytes'] for f in self.scan_results)
            total_size_mb = total_size / (1024 * 1024)
            
            if mode == "size":
                self.info_var.set(
                    f"✅ Scan selesai! Ditemukan {total_files} file >{self.min_size_mb}MB (Total: {total_size_mb:.2f} MB)"
                )
            else:
                self.info_var.set(
                    f"✅ Scan selesai! Ditemukan {total_files} file format non-dokumen (Total: {total_size_mb:.2f} MB)"
                )
            
            # Enable buttons
            if total_files > 0:
                self.export_btn.config(state="normal")
                self.clear_btn.config(state="normal")
            
        except Exception as e:
            progress_window.destroy()
            messagebox.showerror("Error", f"Terjadi kesalahan saat scan:\n{str(e)}")
    
    def scan_folder_recursive(self, folder_path):
        """Scan folder secara recursive berdasarkan mode yang dipilih"""
        try:
            mode = self.scan_mode.get()
            
            if mode == "size":
                min_size_bytes = self.min_size_mb * 1024 * 1024  # Convert MB to bytes
            
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    try:
                        # Skip ignored files (owncloud sync files)
                        if file in self.ignored_files:
                            continue
                        
                        file_path = os.path.join(root, file)
                        file_size = os.path.getsize(file_path)
                        
                        # Get file extension
                        _, ext = os.path.splitext(file)
                        ext = ext.lower()
                        
                        # Check berdasarkan mode
                        if mode == "size":
                            # Mode: File Besar - Check if file >= min_size_mb
                            if file_size >= min_size_bytes:
                                self.scan_results.append({
                                    'name': file,
                                    'size_bytes': file_size,
                                    'size_mb': file_size / (1024 * 1024),
                                    'path': file_path,
                                    'extension': ext if ext else '(no ext)'
                                })
                        else:  # mode == "format"
                            # Mode: Format Non-Dokumen - Check if extension NOT in allowed list
                            if ext not in self.allowed_extensions:
                                self.scan_results.append({
                                    'name': file,
                                    'size_bytes': file_size,
                                    'size_mb': file_size / (1024 * 1024),
                                    'path': file_path,
                                    'extension': ext if ext else '(no ext)'
                                })
                    except Exception as e:
                        # Skip files that can't be accessed
                        continue
        except Exception as e:
            print(f"Error scanning folder: {str(e)}")
    
    def display_results(self):
        """Menampilkan hasil scan di treeview"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add results
        for idx, file_info in enumerate(self.scan_results, 1):
            self.tree.insert("", tk.END, values=(
                idx,
                file_info['name'],
                file_info['extension'],
                f"{file_info['size_mb']:.2f} MB",
                file_info['path']
            ))
    
    def export_to_excel(self):
        """Export hasil scan ke Excel"""
        if not self.scan_results:
            messagebox.showwarning("Peringatan", "Tidak ada data untuk di-export!")
            return
        
        try:
            file_path = filedialog.asksaveasfilename(
                title="Export ke Excel",
                defaultextension=".xlsx",
                initialfile=f"file_besar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                filetypes=[
                    ("Excel Files", "*.xlsx"),
                    ("All Files", "*.*")
                ]
            )
            
            if file_path:
                # Create DataFrame
                data = []
                for idx, file_info in enumerate(self.scan_results, 1):
                    data.append({
                        'No': idx,
                        'Nama File': file_info['name'],
                        'Ekstensi': file_info['extension'],
                        'Ukuran (MB)': round(file_info['size_mb'], 2),
                        'Ukuran (Bytes)': file_info['size_bytes'],
                        'Path Lengkap': file_info['path']
                    })
                
                df = pd.DataFrame(data)
                
                # Export to Excel
                df.to_excel(file_path, index=False, sheet_name="File_Besar")
                
                messagebox.showinfo(
                    "Export Berhasil",
                    f"Data berhasil di-export!\n\n"
                    f"File: {file_path}\n"
                    f"Total: {len(self.scan_results)} file"
                )
        except Exception as e:
            messagebox.showerror("Error", f"Gagal export ke Excel:\n{str(e)}")
    
    def clear_results(self):
        """Clear hasil scan"""
        # Clear treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Clear results
        self.scan_results = []
        
        # Disable buttons
        self.export_btn.config(state="disabled")
        self.clear_btn.config(state="disabled")
        
        # Update info
        self.info_var.set("Hasil scan telah dibersihkan. Pilih folder untuk scan ulang.")
    
    def show_progress_dialog(self, message):
        """Menampilkan dialog progress"""
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Processing...")
        progress_window.geometry("350x100")
        progress_window.resizable(False, False)
        
        # Center window
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (175)
        y = (progress_window.winfo_screenheight() // 2) - (50)
        progress_window.geometry(f'350x100+{x}+{y}')
        
        # Make it modal
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Progress content
        frame = ttk.Frame(progress_window, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(frame, text=message, font=("Arial", 10)).grid(row=0, column=0, pady=(0, 10))
        
        progress_bar = ttk.Progressbar(frame, mode='indeterminate')
        progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        progress_bar.start()
        
        progress_window.update()
        return progress_window
    
    def back_to_menu(self):
        """Kembali ke menu utama"""
        if self.parent_window:
            self.root.destroy()
            self.parent_window.deiconify()