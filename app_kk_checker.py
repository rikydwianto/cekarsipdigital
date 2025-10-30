"""
Cek NO KK App - Form untuk pengecekan Nomor Kartu Keluarga dengan OCR
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
import numpy as np
import re
from datetime import datetime
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

from app_helpers import (
    get_appdata_path,
    get_database_path,
    get_export_path,
    get_responsive_dimensions
)

# Import untuk PDF dan OCR
try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    pytesseract = None
    convert_from_path = None

class CekNoKKApp:
    """Form untuk Cek NO KK (Nomor Kartu Keluarga)"""
    
    def __init__(self, root, parent_window=None):
        self.root = root
        self.parent_window = parent_window
        
        # Initialize variables
        self.results = []
        self.status_var = tk.StringVar(value="✅ Ready - Klik 'PROSES CEK NO KK' untuk memulai")
        self.is_paused = False
        self.is_processing = False
        
        self.setup_window()
        self.create_widgets()
    
    def setup_window(self):
        """Setup window cek no kk"""
        self.root.title("Cek NO KK - Nomor Kartu Keluarga")
        
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Get responsive dimensions
        width, height, self.padding, self.fonts = get_responsive_dimensions(
            900, 700, screen_width, screen_height
        )
        
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(800, 600)
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
        """Membuat widget untuk cek no kk"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding=str(self.padding))
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)  # Results frame expandable
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="👨‍👩‍👧‍👦 CEK NOMOR KK", 
            font=("Arial", self.fonts['title'], "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 5))
        
        # Subtitle
        subtitle_label = ttk.Label(
            main_frame, 
            text="Validasi dan pengecekan Nomor Kartu Keluarga dari Database",
            font=("Arial", self.fonts['subtitle']),
            foreground="gray"
        )
        subtitle_label.grid(row=1, column=0, pady=(0, 20))
        
        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="ℹ️ Informasi", padding="15")
        info_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        database_path = get_database_path()
        info_text = (
            "Proses ini akan:\n\n"
            f"1. Membaca data dari:\n   {database_path}\n"
            "   (Sheet: 02.DATA_ANGGOTA)\n"
            "2. Filter file PDF yang diawali dengan '02'\n"
            "3. Ekstrak NO KK dari PDF menggunakan OCR\n"
            "4. Validasi format NO KK (16 digit angka)\n"
            "5. Menampilkan hasil pengecekan\n"
            "6. Export hasil ke Excel"
        )
        
        info_label = ttk.Label(
            info_frame,
            text=info_text,
            font=("Arial", self.fonts['normal']),
            foreground="blue",
            justify=tk.LEFT
        )
        info_label.grid(row=0, column=0, sticky=tk.W)
        
        # Button frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, pady=(0, 15))
        
        # Proses button (tombol utama)
        self.proses_btn = ttk.Button(
            btn_frame, 
            text="▶️ PROSES CEK NO KK", 
            command=self.proses_cek_nokk,
            style="Accent.TButton"
        )
        self.proses_btn.grid(row=0, column=0, padx=(0, 10), ipadx=20, ipady=10)
        
        # Pause/Resume button
        self.pause_btn = ttk.Button(
            btn_frame, 
            text="⏸️ Pause", 
            command=self.toggle_pause,
            state=tk.DISABLED
        )
        self.pause_btn.grid(row=0, column=1, padx=(10, 10))
        
        # Export button
        self.export_btn = ttk.Button(
            btn_frame, 
            text="💾 Export Hasil", 
            command=self.export_results,
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
        
        # Results frame dengan treeview
        results_frame = ttk.LabelFrame(main_frame, text="Hasil Pengecekan NO KK", padding="10")
        results_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Create scrollbars
        tree_scroll_y = ttk.Scrollbar(results_frame, orient=tk.VERTICAL)
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        tree_scroll_x = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Treeview untuk hasil
        self.tree = ttk.Treeview(
            results_frame,
            columns=("No", "NO KK", "Status", "Panjang", "Format", "Keterangan", "Nama", "Nomor Center", "Status File", "Path"),
            show="headings",
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set
        )
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        # Define columns
        self.tree.heading("No", text="No")
        self.tree.heading("NO KK", text="NO KK")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Panjang", text="Panjang")
        self.tree.heading("Format", text="Format")
        self.tree.heading("Keterangan", text="Keterangan")
        self.tree.heading("Nama", text="Nama Anggota")
        self.tree.heading("Nomor Center", text="Nomor Center")
        self.tree.heading("Status File", text="Status File")
        self.tree.heading("Path", text="Path File")
        
        # Set column widths
        self.tree.column("No", width=50, anchor=tk.CENTER)
        self.tree.column("NO KK", width=150, anchor=tk.CENTER)
        self.tree.column("Status", width=80, anchor=tk.CENTER)
        self.tree.column("Panjang", width=70, anchor=tk.CENTER)
        self.tree.column("Format", width=100, anchor=tk.CENTER)
        self.tree.column("Keterangan", width=250, anchor=tk.W)
        self.tree.column("Nama", width=200, anchor=tk.W)
        self.tree.column("Nomor Center", width=120, anchor=tk.CENTER)
        self.tree.column("Status File", width=100, anchor=tk.CENTER)
        self.tree.column("Path", width=400, anchor=tk.W)
        self.tree.column("Nama", width=200, anchor=tk.W)
        self.tree.column("Nomor Center", width=120, anchor=tk.CENTER)
        self.tree.column("Path", width=400, anchor=tk.W)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status label
        status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            font=("Arial", self.fonts['small']),
          
        )
        status_label.grid(row=5, column=0, sticky=tk.W, pady=(5, 0))
    
    def validate_nokk(self, nokk):
        """Validasi format NO KK"""
        result = {
            "nokk": nokk,
            "valid": False,
            "panjang": len(str(nokk)) if nokk else 0,
            "format": "Invalid",
            "keterangan": ""
        }
        
        # Handle NaN or empty
        if pd.isna(nokk) or nokk == "" or nokk is None:
            result["keterangan"] = "NO KK kosong"
            result["format"] = "Kosong"
            return result
        
        # Convert to string and remove spaces
        nokk_str = str(nokk).strip().replace(" ", "")
        result["nokk"] = nokk_str
        result["panjang"] = len(nokk_str)
        
        # Check panjang
        if len(nokk_str) != 16:
            result["keterangan"] = f"Panjang tidak sesuai (harus 16 digit, saat ini {len(nokk_str)})"
            return result
        
        # Check apakah semua digit
        if not nokk_str.isdigit():
            result["format"] = "Bukan Angka"
            result["keterangan"] = "NO KK harus berisi angka saja"
            return result
        
        # Valid
        result["valid"] = True
        result["format"] = "Valid"
        result["keterangan"] = "Format NO KK valid (16 digit angka)"
        
        return result
    
    def toggle_pause(self):
        """Toggle pause/resume state"""
        self.is_paused = not self.is_paused
        
        if self.is_paused:
            self.pause_btn.config(text="▶️ Resume")
            self.status_var.set("⏸️ PAUSED - Klik 'Resume' untuk melanjutkan")
        else:
            self.pause_btn.config(text="⏸️ Pause")
            self.status_var.set("▶️ RESUMED - Melanjutkan proses...")
    
    def wait_if_paused(self):
        """Wait loop saat paused"""
        while self.is_paused and self.is_processing:
            self.root.update()
            self.root.after(100)  # Check every 100ms
    
    def deskew_image(self, image):
        """Straighten skewed/tilted image menggunakan projection profile"""
        try:
            import numpy as np
            from PIL import Image
            
            # Convert PIL Image to numpy array
            img_array = np.array(image)
            
            # Binarize dengan threshold
            threshold = 128
            binary = img_array < threshold
            
            # Hitung projection profile untuk berbagai sudut (-20 sampai +20 derajat)
            best_angle = 0
            max_variance = 0
            
            for angle in range(-20, 21, 1):
                # Rotate image
                rotated = image.rotate(angle, expand=False, fillcolor=255)
                rotated_array = np.array(rotated.convert('L'))
                rotated_binary = rotated_array < threshold
                
                # Horizontal projection (sum across rows)
                h_projection = np.sum(rotated_binary, axis=1)
                
                # Variance of projection - higher = better alignment
                variance = np.var(h_projection)
                
                if variance > max_variance:
                    max_variance = variance
                    best_angle = angle
            
            # Apply best rotation
            if best_angle != 0:
                print(f"🔄 Deskewing image: {best_angle} degrees")
                deskewed = image.rotate(best_angle, expand=True, fillcolor=255)
                return deskewed
            else:
                print("✅ Image already straight")
                return image
                
        except Exception as e:
            print(f"⚠️ Deskew failed: {str(e)}, using original image")
            return image
    
    def extract_nokk_from_pdf(self, pdf_path):
        """Ekstrak NO KK dari PDF menggunakan OCR dengan fokus ke header"""
        if not os.path.exists(pdf_path):
            return None
        
        try:
            # Auto-detect Tesseract path
            tesseract_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                r"C:\Tesseract-OCR\tesseract.exe",
            ]
            
            tesseract_found = False
            for tess_path in tesseract_paths:
                if os.path.exists(tess_path):
                    pytesseract.pytesseract.tesseract_cmd = tess_path
                    tesseract_found = True
                    break
            
            if not tesseract_found:
                print("⚠️ Tesseract OCR tidak ditemukan. Install dari: https://github.com/UB-Mannheim/tesseract/wiki")
                return None
            
            # Detect Poppler path (same as PDF Tool)
            poppler_paths = [
                "poppler-25.07.0/Library/bin",  # Portable version
                r"C:\Program Files\poppler\Library\bin",
                r"C:\poppler\Library\bin"
            ]
            
            poppler_path = None
            for path in poppler_paths:
                if os.path.exists(path):
                    poppler_path = path
                    break
            
            # Convert PDF to images (first page only)
            try:
                if poppler_path:
                    images = convert_from_path(
                        pdf_path, 
                        first_page=1, 
                        last_page=1, 
                        poppler_path=poppler_path,
                        dpi=400  # Higher DPI untuk OCR lebih akurat
                    )
                else:
                    images = convert_from_path(
                        pdf_path, 
                        first_page=1, 
                        last_page=1,
                        dpi=400
                    )
            except Exception as e:
                print(f"Error converting PDF to image: {str(e)}")
                return None
            
            if not images:
                return None
            
            # Preprocessing image untuk OCR lebih baik
            from PIL import ImageEnhance, ImageFilter, ImageOps
            import numpy as np
            
            image = images[0]
            width, height = image.size
            
            # Crop hanya bagian header (20% atas) - area NO KK biasanya di sini
            header_height = int(height * 0.2)
            image_header = image.crop((0, 0, width, header_height))
            
            # Convert to grayscale
            image_header = image_header.convert('L')
            
            # DESKEW: Straighten image jika miring
            image_header = self.deskew_image(image_header)
            
            # Enhance contrast lebih kuat
            enhancer = ImageEnhance.Contrast(image_header)
            image_header = enhancer.enhance(3.0)
            
            # Enhance sharpness
            enhancer = ImageEnhance.Sharpness(image_header)
            image_header = enhancer.enhance(2.0)
            
            # Threshold untuk binarisasi (black & white)
            image_header = ImageOps.autocontrast(image_header)
            
            # OCR dengan config optimized untuk angka
            configs = [
                '--psm 6 -c tessedit_char_whitelist=0123456789',  # Hanya angka
                '--psm 11 -c tessedit_char_whitelist=0123456789',
                '--psm 6',  # Tanpa whitelist sebagai fallback
            ]
            
            all_text = ""
            for config in configs:
                try:
                    text = pytesseract.image_to_string(image_header, lang='eng', config=config)
                    all_text += text + "\n"
                except:
                    continue
            
            if not all_text:
                return None
            
            # Debug: print extracted text
            print(f"\n📄 Header text from {os.path.basename(pdf_path)}:")
            print(all_text[:300])
            
            # Clean up common OCR errors before pattern matching
            # Ganti huruf yang mirip angka
            cleaned_text = all_text
            replacements = {
                'b': '6',  # huruf b sering dibaca untuk angka 6
                'B': '8',
                'O': '0',
                'o': '0',
                'l': '1',
                'I': '1',
                'S': '5',
                'Z': '2',
            }
            
            for old, new in replacements.items():
                cleaned_text = cleaned_text.replace(old, new)
            
            # Search for 16-digit number pattern (NO KK)
            # Pattern 1: 16 consecutive digits (dari cleaned text)
            pattern = r'\b\d{16}\b'
            matches = re.findall(pattern, cleaned_text)
            
            if matches:
                # Ambil yang pertama (biasanya NO KK di header)
                print(f"✅ Found NO KK (after cleanup): {matches[0]}")
                return matches[0]
            
            # Pattern 2: From original text (tanpa cleanup)
            matches = re.findall(pattern, all_text)
            if matches:
                print(f"✅ Found NO KK: {matches[0]}")
                return matches[0]
            
            # Pattern 3: With spaces/dots (e.g., "3302 0403 0205 2186")
            pattern_with_space = r'(\d{4}[\s\.\-]?\d{4}[\s\.\-]?\d{4}[\s\.\-]?\d{4})'
            matches = re.findall(pattern_with_space, cleaned_text)
            
            if matches:
                # Remove spaces, dots, dashes
                nokk = re.sub(r'[\s\.\-]', '', matches[0])
                if len(nokk) == 16 and nokk.isdigit():
                    print(f"✅ Found NO KK (with separators): {nokk}")
                    return nokk
            
            # Pattern 4: Cari angka 15-18 digit (kadang OCR salah)
            pattern_flexible = r'\d{15,18}'
            matches = re.findall(pattern_flexible, cleaned_text)
            
            if matches:
                # Coba berbagai kemungkinan untuk mendapat 16 digit
                for match in matches:
                    if len(match) == 16:
                        print(f"✅ Found NO KK (flexible): {match}")
                        return match
                    elif len(match) == 17:
                        # Coba ambil 16 digit pertama atau terakhir
                        candidate1 = match[:16]
                        candidate2 = match[1:]
                        # Prioritas yang dimulai dengan 33 (kode Jawa Tengah)
                        if candidate1.startswith('33'):
                            print(f"✅ Found NO KK (17→16, first): {candidate1}")
                            return candidate1
                        elif candidate2.startswith('33'):
                            print(f"✅ Found NO KK (17→16, last): {candidate2}")
                            return candidate2
                        else:
                            print(f"✅ Found NO KK (17→16): {candidate1}")
                            return candidate1
                    elif len(match) == 15:
                        # Mungkin kurang 1 digit, tapi tetap return
                        print(f"⚠️ Found 15 digits (might be incomplete): {match}")
                        # Don't return, keep searching
            
            print("❌ NO KK tidak ditemukan dalam header")
            return None
            
        except Exception as e:
            print(f"❌ Error extracting NO KK from {pdf_path}: {str(e)}")
            return None
    
    def proses_cek_nokk(self):
        """Proses cek NO KK dari database.xlsx - ekstrak dari PDF file yang dimulai dengan 02"""
        # Set processing flag
        self.is_processing = True
        self.is_paused = False
        
        # Enable pause button, disable proses button
        self.pause_btn.config(state=tk.NORMAL, text="⏸️ Pause")
        self.proses_btn.config(state=tk.DISABLED)
        
        # Check OCR availability
        if not OCR_AVAILABLE:
            messagebox.showerror(
                "OCR Not Available",
                "pytesseract atau pdf2image tidak tersedia!\n\n"
                "Install dengan: pip install pytesseract pdf2image\n\n"
                "Dan pastikan Tesseract OCR sudah terinstall di sistem."
            )
            return
        
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results = []
        
        # Check if database.xlsx exists di AppData
        database_path = get_database_path()
        if not os.path.exists(database_path):
            messagebox.showerror(
                "File Tidak Ditemukan",
                f"File database.xlsx tidak ditemukan di:\n{database_path}\n\n"
                "Silakan scan folder arsip digital terlebih dahulu:\n"
                "Menu → Scan Folder Arsip Digital\n\n"
                "Setelah scan selesai, pilih 'Simpan dan Sinkronkan'\n"
                "untuk membuat file database.xlsx"
            )
            self.status_var.set("❌ Error: database.xlsx tidak ditemukan")
            return
        
        try:
            self.status_var.set("🔄 Membaca database.xlsx...")
            self.root.update()
            
            # Read Excel - sheet 02.DATA_ANGGOTA dari AppData
            df = pd.read_excel(database_path, sheet_name="02.DATA_ANGGOTA")
            
            # Check required columns
            required_cols = ["TYPE", "NAMA_FILE", "PATH"]
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                messagebox.showerror(
                    "Kolom Tidak Ditemukan",
                    f"Kolom berikut tidak ditemukan:\n{', '.join(missing_cols)}\n\n"
                    "Pastikan database.xlsx memiliki struktur yang benar."
                )
                self.status_var.set("❌ Error: Struktur database tidak sesuai")
                return
            
            # Filter: TYPE = "FILE" dan NAMA_FILE dimulai dengan "02"
            df_filtered = df[
                (df["TYPE"] == "FILE") & 
                (df["NAMA_FILE"].astype(str).str.startswith("02"))
            ].copy()
            
            total_rows = len(df_filtered)
            
            if total_rows == 0:
                messagebox.showwarning(
                    "Data Kosong",
                    "Tidak ada file yang dimulai dengan '02' di sheet 02.DATA_ANGGOTA!"
                )
                self.status_var.set("⚠️ Warning: Tidak ada data yang sesuai filter")
                return
            
            self.status_var.set(f"🔄 Ditemukan {total_rows} file PDF untuk diproses...")
            self.root.update()
            
            # Get additional columns if available
            id_nama_col = "ID_NAMA_ANGGOTA" if "ID_NAMA_ANGGOTA" in df_filtered.columns else None
            nomor_center_col = "NOMOR_CENTER" if "NOMOR_CENTER" in df_filtered.columns else None
            
            # Process each PDF file
            valid_count = 0
            invalid_count = 0
            not_found_count = 0
            
            for idx, row in df_filtered.iterrows():
                # Check if paused
                self.wait_if_paused()
                
                pdf_path = row["PATH"]
                nama_file = row["NAMA_FILE"]
                id_nama = row[id_nama_col] if id_nama_col else "-"
                nomor_center = row[nomor_center_col] if nomor_center_col else "-"
                
                # Update progress
                current = len(self.results) + 1
                if not self.is_paused:
                    self.status_var.set(f"🔄 Memproses {current}/{total_rows}: {nama_file}...")
                self.root.update()
                
                # Check if file exists
                file_exists = os.path.exists(pdf_path)
                file_status = "✅ Ada" if file_exists else "❌ Tidak Ada"
                
                # Extract NO KK from PDF only if file exists
                nokk = None
                if file_exists:
                    nokk = self.extract_nokk_from_pdf(pdf_path)
                
                if nokk:
                    # Validate
                    result = self.validate_nokk(nokk)
                    result["nama"] = id_nama
                    result["nomor_center"] = nomor_center
                    result["path"] = pdf_path
                    result["file_status"] = file_status
                    
                    # Count
                    if result["valid"]:
                        valid_count += 1
                    else:
                        invalid_count += 1
                elif not file_exists:
                    # File not found
                    not_found_count += 1
                    result = {
                        "nokk": "-",
                        "valid": False,
                        "panjang": 0,
                        "format": "-",
                        "keterangan": "File PDF tidak ditemukan",
                        "nama": id_nama,
                        "nomor_center": nomor_center,
                        "path": pdf_path,
                        "file_status": file_status
                    }
                else:
                    # File exists but NO KK not found in OCR
                    not_found_count += 1
                    result = {
                        "nokk": "-",
                        "valid": False,
                        "panjang": 0,
                        "format": "-",
                        "keterangan": "NO KK tidak ditemukan di PDF",
                        "nama": id_nama,
                        "nomor_center": nomor_center,
                        "path": pdf_path,
                        "file_status": file_status
                    }
                
                # Add to results
                self.results.append(result)
                
                # Add to treeview
                if result["nokk"] == "-":
                    status_icon = "⚠️"
                    tag = "not_found"
                elif result["valid"]:
                    status_icon = "✅"
                    tag = "valid"
                else:
                    status_icon = "❌"
                    tag = "invalid"
                
                self.tree.insert(
                    "", 
                    tk.END, 
                    values=(
                        current,
                        result["nokk"],
                        status_icon,
                        result["panjang"],
                        result["format"],
                        result["keterangan"],
                        id_nama,
                        nomor_center,
                        result.get("file_status", "-"),
                        pdf_path
                    ),
                    tags=(tag,)
                )
            
            # Configure tags untuk warna
            self.tree.tag_configure("valid", foreground="green")
            self.tree.tag_configure("invalid", foreground="red")
            self.tree.tag_configure("not_found", foreground="orange")
            
            # Enable export button
            self.export_btn.config(state=tk.NORMAL)
            
            # Update status
            self.status_var.set(
                f"✅ Selesai: {total_rows} file | ✅ Valid: {valid_count} | ❌ Invalid: {invalid_count} | ⚠️ Tidak Ditemukan: {not_found_count}"
            )
            
            # Show result
            messagebox.showinfo(
                "Proses Selesai",
                f"Pengecekan NO KK dari PDF selesai!\n\n"
                f"Total File PDF: {total_rows}\n"
                f"✅ Valid: {valid_count}\n"
                f"❌ Invalid: {invalid_count}\n"
                f"⚠️ NO KK Tidak Ditemukan: {not_found_count}\n\n"
                f"Klik 'Export Hasil' untuk menyimpan hasil pengecekan."
            )
            
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Terjadi kesalahan saat memproses data:\n\n{str(e)}"
            )
            self.status_var.set(f"❌ Error: {str(e)}")
        finally:
            # Reset processing state
            self.is_processing = False
            self.is_paused = False
            self.pause_btn.config(state=tk.DISABLED, text="⏸️ Pause")
            self.proses_btn.config(state=tk.NORMAL)
    
    def export_results(self):
        """Export hasil pengecekan ke Excel"""
        if not self.results:
            messagebox.showwarning("Tidak Ada Data", "Tidak ada hasil untuk di-export!")
            return
        
        # Ask for save location
        file_path = filedialog.asksaveasfilename(
            title="Export Hasil Pengecekan NO KK",
            defaultextension=".xlsx",
            initialfile=f"cek_nokk_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            self.status_var.set("🔄 Export ke Excel...")
            self.root.update()
            
            # Prepare data
            df_data = []
            for idx, result in enumerate(self.results, 1):
                df_data.append({
                    'No': idx,
                    'NO_KK': result['nokk'],
                    'Status': 'VALID' if result['valid'] else 'INVALID',
                    'Panjang': result['panjang'],
                    'Format': result['format'],
                    'Keterangan': result['keterangan'],
                    'Nama': result.get('nama', '-'),
                    'Nomor_Center': result.get('nomor_center', '-'),
                    'Status_File': result.get('file_status', '-'),
                    'Path': result.get('path', '-')
                })
            
            df = pd.DataFrame(df_data)
            
            # Summary data
            valid_count = len([r for r in self.results if r['valid']])
            invalid_count = len([r for r in self.results if not r['valid']])
            
            summary_data = [{
                'Informasi': 'Total Data',
                'Value': len(self.results)
            }, {
                'Informasi': 'Valid',
                'Value': valid_count
            }, {
                'Informasi': 'Invalid',
                'Value': invalid_count
            }, {
                'Informasi': 'Persentase Valid',
                'Value': f"{(valid_count/len(self.results)*100):.2f}%" if self.results else "0%"
            }, {
                'Informasi': 'Waktu Export',
                'Value': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }]
            
            df_summary = pd.DataFrame(summary_data)
            
            # Export to Excel with multiple sheets
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name="Hasil Cek NO KK", index=False)
                df_summary.to_excel(writer, sheet_name="Summary", index=False)
            
            self.status_var.set(f"✅ Export berhasil: {os.path.basename(file_path)}")
            messagebox.showinfo(
                "Export Berhasil",
                f"Hasil pengecekan berhasil di-export!\n\n"
                f"File: {file_path}\n\n"
                f"Total: {len(self.results)} NO KK\n"
                f"✅ Valid: {valid_count}\n"
                f"❌ Invalid: {invalid_count}"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal export hasil:\n{str(e)}")
            self.status_var.set(f"❌ Error export")
    
    def back_to_menu(self):
        """Kembali ke menu utama"""
        if self.parent_window:
            self.root.destroy()
            self.parent_window.deiconify()
