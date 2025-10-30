"""
Cek Pengajuan Dana App - Form untuk pengecekan pengajuan dana dari surat keluar
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

class CekPengajuanDanaApp:
    """Form untuk Cek Pengajuan Dana dari Surat Keluar"""
    
    def __init__(self, root, parent_window=None):
        self.root = root
        self.parent_window = parent_window
        
        self.setup_window()
        self.create_widgets()
    
    def setup_window(self):
        """Setup window cek pengajuan dana"""
        self.root.title("Cek Pengajuan Dana - Surat Keluar")
        
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Set responsive window size (90% of screen width, 90% of height)
        window_width = int(screen_width * 0.9)
        window_height = int(screen_height * 0.9)
        
        # Minimum size constraints
        min_width = 1200
        min_height = 600
        window_width = max(window_width, min_width)
        window_height = max(window_height, min_height)
        
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.minsize(min_width, min_height)
        self.root.resizable(True, True)
        
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
        """Membuat widget untuk cek pengajuan dana dengan scrollable canvas"""
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
        
        # Main frame
        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure main_frame
        scrollable_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)  # Row 5 untuk results_frame (treeview)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="💰 CEK PENGAJUAN DANA", 
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # Subtitle
        subtitle_label = ttk.Label(
            main_frame, 
            text="Scan file PENGAJUAN_DANA.xlsm dari folder Surat Keluar",
            font=("Arial", 10),
            foreground="gray"
        )
        subtitle_label.grid(row=1, column=0, pady=(0, 5))
        
        # Info label
        info_label = ttk.Label(
            main_frame,
            text="💡 Tip: Double-click pada baris untuk membuka file Excel",
            font=("Arial", 9, "italic"),
            foreground="blue"
        )
        info_label.grid(row=2, column=0, pady=(0, 15))
        
        # Button frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, pady=(0, 10))
        
        # Scan button
        scan_btn = ttk.Button(
            btn_frame, 
            text="🔍 Mulai Scan", 
            command=self.scan_pengajuan_dana,
            style="Accent.TButton"
        )
        scan_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Analisa button
        self.analisa_btn = ttk.Button(
            btn_frame, 
            text="🔬 Analisa Data", 
            command=self.analisa_data,
            state=tk.DISABLED
        )
        self.analisa_btn.grid(row=0, column=1, padx=(10, 10))
        
        # Export button
        self.export_btn = ttk.Button(
            btn_frame, 
            text="📊 Export ke Excel", 
            command=self.export_to_excel,
            state=tk.DISABLED
        )
        self.export_btn.grid(row=0, column=2, padx=(10, 10))
        
        # Back button
        if self.parent_window:
            back_btn = ttk.Button(
                btn_frame, 
                text="⬅️ Kembali", 
                command=self.back_to_menu
            )
            back_btn.grid(row=0, column=3, padx=(10, 0))
        
        # Status info bar (di atas treeview)
        status_info_frame = ttk.Frame(main_frame)
        status_info_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(5, 10))
        
        self.status_var = tk.StringVar(value="Siap untuk scan")
        status_label = ttk.Label(
            status_info_frame,
            textvariable=self.status_var,
            font=("Arial", 9),
            foreground="blue"
        )
        status_label.grid(row=0, column=0, sticky=tk.W)
        
        # Results frame dengan treeview
        results_frame = ttk.LabelFrame(main_frame, text="Hasil Scan", padding="10")
        results_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Create scrollbars (vertical dan horizontal)
        tree_scroll_y = ttk.Scrollbar(results_frame, orient=tk.VERTICAL)
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        tree_scroll_x = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        self.tree = ttk.Treeview(
            results_frame,
            columns=("No", "Tahun", "Bulan", "Nomor Surat", "Nama File", "Path"),
            show="headings",
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set
        )
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        # Define columns
        self.tree.heading("No", text="No")
        self.tree.heading("Tahun", text="Tahun")
        self.tree.heading("Bulan", text="Bulan")
        self.tree.heading("Nomor Surat", text="Nomor Surat")
        self.tree.heading("Nama File", text="Nama File")
        self.tree.heading("Path", text="Path Lengkap")
        
        # Set column widths
        self.tree.column("No", width=50, anchor=tk.CENTER)
        self.tree.column("Tahun", width=80, anchor=tk.CENTER)
        self.tree.column("Bulan", width=100, anchor=tk.CENTER)
        self.tree.column("Nomor Surat", width=100, anchor=tk.CENTER)
        self.tree.column("Nama File", width=250, anchor=tk.W)
        self.tree.column("Path", width=300, anchor=tk.W)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Bind double-click event untuk membuka file
        self.tree.bind("<Double-Button-1>", self.open_selected_file)
        
        # Data storage
        self.scan_results = []
        
        # Update canvas width to match window
        def _configure_canvas(event):
            canvas.itemconfig(canvas.find_withtag("all")[0], width=event.width)
        
        canvas.bind("<Configure>", _configure_canvas)
    
    def scan_pengajuan_dana(self):
        """Scan folder surat keluar untuk file PENGAJUAN_DANA.xlsm"""
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.scan_results = []
        
        # Gunakan default folder dari config
        default_folder = config_manager.get_default_folder()
        
        if not default_folder or not os.path.exists(default_folder):
            messagebox.showwarning(
                "Folder Tidak Ditemukan",
                "Default folder belum diset atau tidak ditemukan.\n\n"
                "Silakan set default folder di menu Pengaturan terlebih dahulu."
            )
            return
        
        # Path ke folder surat keluar
        base_path = os.path.join(default_folder, "01.SURAT_MENYURAT", "02.SURAT_KELUAR")
        
        if not os.path.exists(base_path):
            messagebox.showerror(
                "Folder Tidak Ada",
                f"Folder Surat Keluar tidak ditemukan:\n{base_path}\n\n"
                f"Pastikan struktur folder sudah benar."
            )
            return
        
        self.status_var.set("🔄 Scanning...")
        self.root.update()
        
        # Nama-nama bulan
        bulan_names = {
            "01": "JANUARI", "02": "FEBRUARI", "03": "MARET", "04": "APRIL",
            "05": "MEI", "06": "JUNI", "07": "JULI", "08": "AGUSTUS",
            "09": "SEPTEMBER", "10": "OKTOBER", "11": "NOVEMBER", "12": "DESEMBER"
        }
        
        # Scan dari tahun 2020 sampai tahun sekarang + 1
        current_year = datetime.now().year
        found_count = 0
        
        for year in range(2020, current_year + 2):
            year_folder = os.path.join(base_path, str(year))
            
            if not os.path.exists(year_folder):
                continue
            
            # Loop untuk setiap bulan (01-12)
            for bulan_code, bulan_name in bulan_names.items():
                # Format: 01.JANUARI, 02.FEBRUARI, dst
                bulan_folder_name = f"{bulan_code}.{bulan_name}"
                bulan_folder = os.path.join(year_folder, bulan_folder_name)
                
                if not os.path.exists(bulan_folder):
                    continue
                
                # Scan semua file di folder bulan
                try:
                    for file in os.listdir(bulan_folder):
                        # Skip file temporary (dimulai dengan ~)
                        if file.startswith('~'):
                            continue
                        
                        # Cek apakah file berakhiran PENGAJUAN_DANA.xlsm
                        if file.upper().endswith("PENGAJUAN_DANA.XLSM"):
                            # Extract nomor surat (3 digit di awal)
                            nomor_surat = file[:3] if len(file) >= 3 else "???"
                            
                            file_path = os.path.join(bulan_folder, file)
                            
                            # Simpan hasil (tanpa data analisa dulu)
                            self.scan_results.append({
                                "tahun": year,
                                "bulan": bulan_name,
                                "bulan_code": bulan_code,
                                "nomor_surat": nomor_surat,
                                "nama_file": file,
                                "path": file_path,
                                "nomor_surat_f8": "",  # Akan diisi saat analisa
                                "nominal_input": "",  # Akan diisi saat analisa
                                "nominal_kebutuhan": "",  # Akan diisi saat analisa
                                "status_balance": "",  # Akan diisi saat analisa
                                "tanggal_disburse_awal": "",  # Akan diisi saat analisa
                                "tanggal_disburse_akhir": "",  # Akan diisi saat analisa
                                "nama_bm": "",  # Akan diisi saat analisa
                                "data_analisa": {}  # Untuk data analisa lainnya
                            })
                            
                            found_count += 1
                            
                            # Insert ke treeview
                            self.tree.insert("", tk.END, values=(
                                found_count,
                                year,
                                bulan_name,
                                nomor_surat,
                                file,
                                file_path
                            ))
                except Exception as e:
                    print(f"Error scanning {bulan_folder}: {e}")
        
        # Update status
        if found_count > 0:
            self.status_var.set(f"✅ Ditemukan {found_count} file PENGAJUAN_DANA.xlsm")
            self.export_btn.config(state=tk.NORMAL)
            self.analisa_btn.config(state=tk.NORMAL)
            messagebox.showinfo(
                "Scan Selesai",
                f"Berhasil menemukan {found_count} file PENGAJUAN_DANA.xlsm\n\n"
                f"Klik 'Export ke Excel' untuk menyimpan hasil.\n"
                f"Klik 'Analisa Data' untuk ekstrak data dari dalam file."
            )
        else:
            self.status_var.set("⚠️ Tidak ada file PENGAJUAN_DANA.xlsm ditemukan")
            self.export_btn.config(state=tk.DISABLED)
            self.analisa_btn.config(state=tk.DISABLED)
            messagebox.showinfo(
                "Scan Selesai",
                "Tidak ditemukan file PENGAJUAN_DANA.xlsm\n\n"
                f"Path yang di-scan: {base_path}"
            )
    
    def analisa_data(self):
        """Analisa data dari dalam file PENGAJUAN_DANA.xlsm"""
        if not self.scan_results:
            messagebox.showwarning("Tidak Ada Data", "Tidak ada data untuk dianalisa!")
            return
        
        # Konfirmasi dengan user
        result = messagebox.askyesno(
            "Konfirmasi Analisa",
            f"Akan menganalisa {len(self.scan_results)} file PENGAJUAN_DANA.xlsm\n\n"
            f"Proses ini akan:\n"
            f"• Membaca data dari dalam setiap file\n"
            f"• Mengekstrak nomor surat dari cell F8\n"
            f"• Menambahkan kolom data hasil analisa\n\n"
            f"Lanjutkan?"
        )
        
        if not result:
            return
        
        # Progress dialog
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Analisa Data...")
        progress_window.geometry("400x150")
        progress_window.resizable(False, False)
        
        # Center window
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - 200
        y = (progress_window.winfo_screenheight() // 2) - 75
        progress_window.geometry(f'400x150+{x}+{y}')
        
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Progress content
        progress_frame = ttk.Frame(progress_window, padding="20")
        progress_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        progress_label = ttk.Label(
            progress_frame, 
            text="Memproses file...",
            font=("Arial", 10)
        )
        progress_label.grid(row=0, column=0, pady=(0, 10))
        
        progress_bar = ttk.Progressbar(
            progress_frame, 
            mode='determinate',
            maximum=len(self.scan_results)
        )
        progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        status_label = ttk.Label(
            progress_frame,
            text="",
            font=("Arial", 9),
            foreground="gray"
        )
        status_label.grid(row=2, column=0, pady=(10, 0))
        
        progress_window.update()
        
        # Analisa setiap file
        success_count = 0
        error_count = 0
        
        for idx, result in enumerate(self.scan_results):
            file_path = result['path']
            file_name = result['nama_file']
            
            # Update progress
            progress_bar['value'] = idx + 1
            status_label.config(text=f"File {idx+1}/{len(self.scan_results)}: {file_name}")
            progress_window.update()
            
            try:
                # === SHEET SURAT ===
                df_surat = pd.read_excel(file_path, sheet_name='Surat', header=None)
                
                # Ambil Nomor Surat dari cell F8 (row 7, col 5 - zero-indexed)
                nomor_surat_file = df_surat.iloc[7, 5] if len(df_surat) > 7 and len(df_surat.columns) > 5 else None
                
                # Ambil Nominal Input Kebutuhan dari cell I8 (row 7, col 8 - zero-indexed)
                nominal_input = df_surat.iloc[7, 8] if len(df_surat) > 7 and len(df_surat.columns) > 8 else None
                
                # === SHEET LAPORAN ===
                status_balance = None
                nominal_kebutuhan = None
                nama_bm = None
                
                try:
                    df_laporan = pd.read_excel(file_path, sheet_name='Laporan', header=None)
                    
                    # Ambil Status Balance dari cell A4 (row 3, col 0 - zero-indexed)
                    if len(df_laporan) > 3 and len(df_laporan.columns) > 0:
                        cell_a4 = str(df_laporan.iloc[3, 0])
                        if 'Ket.' in cell_a4 and ':' in cell_a4:
                            parts = cell_a4.split(':', 1)
                            if len(parts) > 1:
                                status_balance = parts[1].strip()
                    
                    # Ambil Nominal Kebutuhan dari cell F68 (row 67, col 5 - zero-indexed)
                    if len(df_laporan) > 67 and len(df_laporan.columns) > 5:
                        nominal_kebutuhan = df_laporan.iloc[67, 5]
                    
                    # Ambil Nama BM dari cell A83 (row 82, col 0 - zero-indexed)
                    if len(df_laporan) > 82 and len(df_laporan.columns) > 0:
                        nama_bm = df_laporan.iloc[82, 0]
                
                except Exception as e_laporan:
                    pass  # Jika gagal baca sheet Laporan, set None
                
                # === SHEET LAMPIRAN ===
                tanggal_disburse_awal = None
                tanggal_disburse_akhir = None
                
                try:
                    df_lampiran = pd.read_excel(file_path, sheet_name='Lampiran', header=None)
                    
                    # Ambil Tanggal Disburse Awal dari cell C3 (row 2, col 2 - zero-indexed)
                    if len(df_lampiran) > 2 and len(df_lampiran.columns) > 2:
                        tanggal_disburse_awal = df_lampiran.iloc[2, 2]
                    
                    # Ambil Tanggal Disburse Akhir dari cell E3 (row 2, col 4 - zero-indexed)
                    if len(df_lampiran) > 2 and len(df_lampiran.columns) > 4:
                        tanggal_disburse_akhir = df_lampiran.iloc[2, 4]
                
                except Exception as e_lampiran:
                    pass  # Jika gagal baca sheet Lampiran, set None
                
                # Simpan hasil analisa
                result['nomor_surat_file'] = nomor_surat_file
                result['nominal_input'] = nominal_input
                result['status_balance'] = status_balance
                result['nominal_kebutuhan'] = nominal_kebutuhan
                result['tanggal_disburse_awal'] = tanggal_disburse_awal
                result['tanggal_disburse_akhir'] = tanggal_disburse_akhir
                result['nama_bm'] = nama_bm
                result['status_analisa'] = 'SUCCESS'
                success_count += 1
                
            except Exception as e:
                result['nomor_surat_file'] = None
                result['nominal_input'] = None
                result['status_balance'] = None
                result['nominal_kebutuhan'] = None
                result['tanggal_disburse_awal'] = None
                result['tanggal_disburse_akhir'] = None
                result['nama_bm'] = None
                result['status_analisa'] = f'ERROR: {str(e)}'
                error_count += 1
        
        progress_window.destroy()
        
        # Update treeview dengan menambah kolom
        # Clear dan rebuild treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Update columns untuk include data analisa (14 kolom total)
        self.tree['columns'] = ("No", "Tahun", "Bulan", "Nomor Surat", "Nomor di File", 
                                "Nominal Input", "Nominal Kebutuhan", "Status Balance", 
                                "Tgl Disburse Awal", "Tgl Disburse Akhir", "Nama BM", 
                                "Status", "Nama File", "Path")
        
        # Redefine headings
        self.tree.heading("No", text="No")
        self.tree.heading("Tahun", text="Tahun")
        self.tree.heading("Bulan", text="Bulan")
        self.tree.heading("Nomor Surat", text="No. Surat (Nama)")
        self.tree.heading("Nomor di File", text="No. Surat (F8)")
        self.tree.heading("Nominal Input", text="Nominal Input")
        self.tree.heading("Nominal Kebutuhan", text="Nominal Kebutuhan")
        self.tree.heading("Status Balance", text="Status Balance")
        self.tree.heading("Tgl Disburse Awal", text="Tgl Disburse Awal")
        self.tree.heading("Tgl Disburse Akhir", text="Tgl Disburse Akhir")
        self.tree.heading("Nama BM", text="Nama BM")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Nama File", text="Nama File")
        self.tree.heading("Path", text="Path Lengkap")
        
        # Set column widths
        self.tree.column("No", width=50, anchor=tk.CENTER)
        self.tree.column("Tahun", width=70, anchor=tk.CENTER)
        self.tree.column("Bulan", width=90, anchor=tk.CENTER)
        self.tree.column("Nomor Surat", width=100, anchor=tk.CENTER)
        self.tree.column("Nomor di File", width=100, anchor=tk.CENTER)
        self.tree.column("Nominal Input", width=120, anchor=tk.E)
        self.tree.column("Nominal Kebutuhan", width=120, anchor=tk.E)
        self.tree.column("Status Balance", width=100, anchor=tk.CENTER)
        self.tree.column("Tgl Disburse Awal", width=120, anchor=tk.CENTER)
        self.tree.column("Tgl Disburse Akhir", width=120, anchor=tk.CENTER)
        self.tree.column("Nama BM", width=150, anchor=tk.W)
        self.tree.column("Status", width=80, anchor=tk.CENTER)
        self.tree.column("Nama File", width=200, anchor=tk.W)
        self.tree.column("Path", width=250, anchor=tk.W)
        
        # Repopulate treeview
        for idx, result in enumerate(self.scan_results, 1):
            nomor_file = result.get('nomor_surat_file', '')
            nominal_input = result.get('nominal_input', '')
            nominal_kebutuhan = result.get('nominal_kebutuhan', '')
            status_balance = result.get('status_balance', '')
            tgl_disburse_awal = result.get('tanggal_disburse_awal', '')
            tgl_disburse_akhir = result.get('tanggal_disburse_akhir', '')
            nama_bm = result.get('nama_bm', '')
            status = result.get('status_analisa', '')
            
            # Tentukan status display
            if status == 'SUCCESS':
                status_display = "✅"
            else:
                status_display = "❌"
            
            # Format tanggal jika ada
            tgl_awal_str = str(tgl_disburse_awal) if tgl_disburse_awal else '-'
            tgl_akhir_str = str(tgl_disburse_akhir) if tgl_disburse_akhir else '-'
            
            self.tree.insert("", tk.END, values=(
                idx,
                result['tahun'],
                result['bulan'],
                result['nomor_surat'],
                nomor_file if nomor_file else '-',
                nominal_input if nominal_input else '-',
                nominal_kebutuhan if nominal_kebutuhan else '-',
                status_balance if status_balance else '-',
                tgl_awal_str,
                tgl_akhir_str,
                nama_bm if nama_bm else '-',
                status_display,
                result['nama_file'],
                result['path']
            ))
        
        # Show result
        messagebox.showinfo(
            "Analisa Selesai",
            f"Analisa data selesai!\n\n"
            f"✅ Berhasil: {success_count} file\n"
            f"❌ Error: {error_count} file\n\n"
            f"Data yang diekstrak:\n"
            f"• Nomor Surat (F8) - Sheet Surat\n"
            f"• Nominal Input (I8) - Sheet Surat\n"
            f"• Nominal Kebutuhan (F68) - Sheet Laporan\n"
            f"• Status Balance (A4) - Sheet Laporan\n"
            f"• Tanggal Disburse Awal (C3) - Sheet Lampiran\n"
            f"• Tanggal Disburse Akhir (E3) - Sheet Lampiran\n"
            f"• Nama BM (A83) - Sheet Laporan"
        )
        
        self.status_var.set(f"✅ Analisa selesai: {success_count} sukses, {error_count} error")
    
    def export_to_excel(self):
        """Export hasil scan ke Excel"""
        if not self.scan_results:
            messagebox.showwarning("Tidak Ada Data", "Tidak ada data untuk di-export!")
            return
        
        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            title="Export Hasil Scan Pengajuan Dana",
            defaultextension=".xlsx",
            initialfile=f"pengajuan_dana_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            # Create DataFrame
            df = pd.DataFrame(self.scan_results)
            
            # Check apakah sudah ada data analisa
            has_analisa = 'nomor_surat_file' in df.columns
            
            if has_analisa:
                # Include semua kolom analisa
                df = df[["tahun", "bulan", "bulan_code", "nomor_surat", "nomor_surat_file",
                        "nominal_input", "nominal_kebutuhan", "status_balance", 
                        "tanggal_disburse_awal", "tanggal_disburse_akhir", "nama_bm",
                        "status_analisa", "nama_file", "path"]]
                
                # Rename columns untuk export
                df.columns = ["Tahun", "Bulan", "Kode Bulan", "Nomor Surat (Nama File)", 
                             "Nomor Surat (F8)", "Nominal Input Kebutuhan (I8)", 
                             "Nominal Kebutuhan (F68)", "Status Balance (A4)", 
                             "Tanggal Disburse Awal (C3)", "Tanggal Disburse Akhir (E3)", 
                             "Nama BM (A83)", "Status Analisa", "Nama File", "Path Lengkap"]
            else:
                # Tanpa kolom analisa (export biasa)
                df = df[["tahun", "bulan", "bulan_code", "nomor_surat", "nama_file", "path"]]
                
                # Rename columns untuk export
                df.columns = ["Tahun", "Bulan", "Kode Bulan", "Nomor Surat", "Nama File", "Path Lengkap"]
            
            # Export ke Excel
            df.to_excel(file_path, index=False, sheet_name="Pengajuan Dana")
            
            info_msg = f"Data berhasil di-export!\n\n"
            info_msg += f"File: {os.path.basename(file_path)}\n"
            info_msg += f"Total: {len(self.scan_results)} rows"
            
            if has_analisa:
                info_msg += f"\n\nIncludes data analisa lengkap:\n"
                info_msg += f"• Nomor Surat (F8) - Sheet Surat\n"
                info_msg += f"• Nominal Input (I8) - Sheet Surat\n"
                info_msg += f"• Nominal Kebutuhan (F68) - Sheet Laporan\n"
                info_msg += f"• Status Balance (A4) - Sheet Laporan\n"
                info_msg += f"• Tanggal Disburse Awal (C3) - Sheet Lampiran\n"
                info_msg += f"• Tanggal Disburse Akhir (E3) - Sheet Lampiran\n"
                info_msg += f"• Nama BM (A83) - Sheet Laporan\n"
                info_msg += f"• Status analisa"
            
            messagebox.showinfo("Export Berhasil", info_msg)
            
        except Exception as e:
            messagebox.showerror("Export Gagal", f"Gagal export ke Excel:\n{str(e)}")
    
    def open_selected_file(self, event):
        """Buka file Excel yang dipilih dengan double-click"""
        # Ambil item yang dipilih
        selected_item = self.tree.selection()
        
        if not selected_item:
            return
        
        # Ambil values dari item yang dipilih
        item_values = self.tree.item(selected_item[0], "values")
        
        if not item_values:
            return
        
        # Path ada di kolom terakhir
        # Sebelum analisa: 6 kolom, path di index 5
        # Setelah analisa: 9 kolom, path di index 8
        file_path = item_values[-1]  # Ambil kolom terakhir (Path)
        
        # Validasi file exists
        if not os.path.exists(file_path):
            messagebox.showerror(
                "File Tidak Ditemukan",
                f"File tidak ditemukan:\n{file_path}"
            )
            return
        
        # Buka file dengan aplikasi default
        try:
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror(
                "Gagal Membuka File",
                f"Gagal membuka file:\n{str(e)}\n\n"
                f"Path: {file_path}"
            )
    
    def back_to_menu(self):
        """Kembali ke menu utama"""
        if self.parent_window:
            self.root.destroy()
            self.parent_window.deiconify()
