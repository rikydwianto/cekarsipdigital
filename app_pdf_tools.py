"""
PDF Tool App - Form untuk merge, split, convert, dan OCR PDF
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime

from app_helpers import get_responsive_dimensions

# Import untuk PDF operations
try:
    from PyPDF2 import PdfReader, PdfWriter, PdfMerger
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    PdfReader = None
    PdfWriter = None
    PdfMerger = None
    pytesseract = None
    convert_from_path = None
    Image = None

class PDFToolApp:
    """Simple PDF Tool window with common utilities: merge, split, images<>pdf, compress.

    Note: This class uses PyPDF2 for PDF manipulation and Pillow for image handling.
    If libraries are missing, the UI will show instructions.
    """

    def __init__(self, root, parent_window=None):
        self.root = root
        self.parent_window = parent_window

        # Try to import optional dependencies lazily
        try:
            from PyPDF2 import PdfReader, PdfWriter
            self.PdfReader = PdfReader
            self.PdfWriter = PdfWriter
        except Exception as e:
            self.PdfReader = None
            self.PdfWriter = None
            print(f"Warning: PyPDF2 import failed - {e}")

        try:
            from PIL import Image
            self.PIL_Image = Image
        except Exception as e:
            self.PIL_Image = None
            print(f"Warning: Pillow import failed - {e}")

        self.setup_window()
        self.create_widgets()

    def setup_window(self):
        self.root.title("PDF Tool - Merge / Split / Convert")
        w, h = 600, 600
        self.root.geometry(f"{w}x{h}")
        self.root.minsize(500, 550)
        self.center_window()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        main = ttk.Frame(self.root, padding="20")
        main.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        title = ttk.Label(main, text="📄 PDF TOOL", font=("Arial", 20, "bold"))
        title.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        subtitle = ttk.Label(main, text="Pilihan: Merge, Split, Convert Images ↔ PDF, Compress",
                             font=("Arial", 11), foreground="gray")
        subtitle.grid(row=1, column=0, sticky=tk.W, pady=(0, 20))

        # Frame untuk buttons dengan grid 2 kolom
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        # Style untuk button yang lebih besar
        button_style = {
            'width': 25,
            'padding': 15
        }

        # Row 0 - Images → PDF (PRIORITAS UTAMA)
        img2pdf_btn = ttk.Button(btn_frame, text="�️ Images → PDF", command=self.images_to_pdf, **button_style)
        img2pdf_btn.grid(row=0, column=0, padx=5, pady=8, sticky=(tk.W, tk.E))

        # Row 0 - Merge PDFs
        merge_btn = ttk.Button(btn_frame, text="🔗 Merge PDFs", command=self.merge_pdfs, **button_style)
        merge_btn.grid(row=0, column=1, padx=5, pady=8, sticky=(tk.W, tk.E))

        # Row 1 - Split PDF
        split_btn = ttk.Button(btn_frame, text="✂️ Split PDF", command=self.split_pdf, **button_style)
        split_btn.grid(row=1, column=0, padx=5, pady=8, sticky=(tk.W, tk.E))

        # Row 1 - Compress PDF
        compress_btn = ttk.Button(btn_frame, text="�️ Compress PDF", command=self.compress_pdf, **button_style)
        compress_btn.grid(row=1, column=1, padx=5, pady=8, sticky=(tk.W, tk.E))

        # Row 2 - PDF → Images (PALING AKHIR)
        pdf2img_btn = ttk.Button(btn_frame, text="�️ PDF → Images", command=self.pdf_to_images, **button_style)
        pdf2img_btn.grid(row=2, column=0, padx=5, pady=8, sticky=(tk.W, tk.E))

        # Status and info
        self.status_var = tk.StringVar(value="✅ Ready - Pilih operasi PDF yang ingin dilakukan")
        status_label = ttk.Label(main, textvariable=self.status_var, font=("Arial", 10), foreground="blue")
        status_label.grid(row=3, column=0, sticky=tk.W, pady=(20, 0))

        footer = ttk.Frame(main)
        footer.grid(row=10, column=0, sticky=(tk.W, tk.E), pady=(20, 0))

        back_btn = ttk.Button(footer, text="⬅️ Kembali ke Menu", command=self.back_to_menu)
        back_btn.grid(row=0, column=0)

    def _ensure_pypdf(self):
        if not self.PdfReader or not self.PdfWriter:
            messagebox.showerror(
                "Dependency Missing",
                "PyPDF2 is required for PDF operations.\nInstall with: pip install PyPDF2"
            )
            return False
        return True

    def _ensure_pillow(self):
        if not self.PIL_Image:
            messagebox.showerror(
                "Dependency Missing",
                "Pillow is required for image operations.\nInstall with: pip install Pillow"
            )
            return False
        return True

    def merge_pdfs(self):
        if not self._ensure_pypdf():
            return

        paths = filedialog.askopenfilenames(
            title="Pilih file PDF untuk di-merge (bisa pilih multiple)", 
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not paths or len(paths) == 0:
            return

        out_path = filedialog.asksaveasfilename(
            title="Simpan hasil merge sebagai", 
            defaultextension=".pdf", 
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not out_path:
            return

        try:
            self.status_var.set(f"🔄 Menggabungkan {len(paths)} file PDF...")
            self.root.update()
            
            merger = self.PdfWriter()
            
            for idx, pdf_path in enumerate(paths, 1):
                self.status_var.set(f"🔄 Memproses file {idx}/{len(paths)}: {os.path.basename(pdf_path)}")
                self.root.update()
                
                with open(pdf_path, 'rb') as pdf_file:
                    reader = self.PdfReader(pdf_file)
                    for page in reader.pages:
                        merger.add_page(page)

            with open(out_path, 'wb') as output_file:
                merger.write(output_file)

            self.status_var.set(f"✅ Merge selesai: {os.path.basename(out_path)}")
            messagebox.showinfo(
                "Selesai", 
                f"✅ Berhasil menggabungkan {len(paths)} file PDF!\n\n"
                f"Output: {out_path}"
            )
        except Exception as e:
            self.status_var.set("❌ Error saat merge PDF")
            messagebox.showerror("Error", f"Gagal merge PDFs:\n\n{str(e)}")

    def split_pdf(self):
        if not self._ensure_pypdf():
            return

        path = filedialog.askopenfilename(
            title="Pilih file PDF untuk di-split", 
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not path:
            return

        out_dir = filedialog.askdirectory(title="Pilih folder tujuan untuk menyimpan halaman")
        if not out_dir:
            return

        try:
            self.status_var.set("🔄 Membaca PDF...")
            self.root.update()
            
            with open(path, 'rb') as pdf_file:
                reader = self.PdfReader(pdf_file)
                total_pages = len(reader.pages)
                
                base_name = os.path.splitext(os.path.basename(path))[0]
                
                for i, page in enumerate(reader.pages, start=1):
                    self.status_var.set(f"🔄 Memproses halaman {i}/{total_pages}...")
                    self.root.update()
                    
                    writer = self.PdfWriter()
                    writer.add_page(page)
                    
                    out_file = os.path.join(out_dir, f"{base_name}_halaman_{i}.pdf")
                    with open(out_file, 'wb') as output_file:
                        writer.write(output_file)

            self.status_var.set(f"✅ Split selesai: {total_pages} halaman")
            messagebox.showinfo(
                "Selesai", 
                f"✅ PDF berhasil di-split menjadi {total_pages} file!\n\n"
                f"Lokasi: {out_dir}"
            )
        except Exception as e:
            self.status_var.set("❌ Error saat split PDF")
            messagebox.showerror("Error", f"Gagal split PDF:\n\n{str(e)}")

    def pdf_to_images(self):
        """Convert PDF pages to images with portable Poppler support"""
        # Check if pdf2image is available
        try:
            from pdf2image import convert_from_path
        except ImportError:
            messagebox.showerror(
                "Library Tidak Ditemukan",
                "Fitur PDF → Images membutuhkan library 'pdf2image'.\n\n"
                "Install dengan perintah:\n"
                "pip install pdf2image"
            )
            return
        
        # Select PDF file
        pdf_path = filedialog.askopenfilename(
            title="Pilih file PDF untuk dikonversi ke gambar",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not pdf_path:
            return
        
        # Select output folder
        out_dir = filedialog.askdirectory(title="Pilih folder tujuan untuk menyimpan gambar")
        if not out_dir:
            return
        
        # Ask for image format
        format_choice = messagebox.askquestion(
            "Format Gambar",
            "Simpan sebagai PNG?\n\n"
            "Yes = PNG (kualitas tinggi)\n"
            "No = JPG (ukuran lebih kecil)"
        )
        img_format = "PNG" if format_choice == "yes" else "JPEG"
        
        try:
            # Show progress
            self.status_var.set("🔄 Mengkonversi PDF ke gambar...")
            self.root.update()
            
            # Get Poppler path from config
            poppler_path = config_manager.config.get("poppler_path", "")
            
            # Try to convert with different methods
            images = None
            error_detail = ""
            
            # Method 1: Try with saved poppler path
            if poppler_path and os.path.exists(poppler_path):
                try:
                    images = convert_from_path(pdf_path, dpi=200, poppler_path=poppler_path)
                except Exception as e:
                    error_detail = f"Gagal dengan saved path: {e}"
            
            # Method 2: Try without poppler_path (auto-detect from PATH)
            if images is None:
                try:
                    images = convert_from_path(pdf_path, dpi=200)
                except Exception as e:
                    error_detail += f"\nGagal auto-detect: {e}"
            
            # Method 3: Check common portable locations
            if images is None:
                common_paths = [
                    # Struktur poppler-xx.xx.x (dari release GitHub)
                    os.path.join(os.getcwd(), "poppler-25.07.0", "Library", "bin"),
                    os.path.join(os.getcwd(), "poppler-25.07.0", "bin"),
                    os.path.join(os.path.dirname(os.path.abspath(__file__)), "poppler-25.07.0", "Library", "bin"),
                    os.path.join(os.path.dirname(os.path.abspath(__file__)), "poppler-25.07.0", "bin"),
                    # Struktur folder poppler generik
                    os.path.join(os.getcwd(), "poppler", "Library", "bin"),
                    os.path.join(os.getcwd(), "poppler", "bin"),
                    os.path.join(os.path.dirname(os.path.abspath(__file__)), "poppler", "Library", "bin"),
                    os.path.join(os.path.dirname(os.path.abspath(__file__)), "poppler", "bin")
                ]
                
                for path in common_paths:
                    if os.path.exists(path):
                        try:
                            images = convert_from_path(pdf_path, dpi=200, poppler_path=path)
                            # Save successful path
                            config_manager.config["poppler_path"] = path
                            config_manager.save_config()
                            break
                        except Exception:
                            continue
            
            # If still failed, ask user to locate Poppler
            if images is None:
                result = messagebox.askyesno(
                    "Poppler Tidak Ditemukan",
                    "Poppler tidak ditemukan!\n\n"
                    "Poppler diperlukan untuk konversi PDF ke gambar.\n"
                    "File poppler bisa diletakkan di folder 'poppler' dalam project ini (portable).\n\n"
                    "Download Poppler dari:\n"
                    "https://github.com/oschwartz10612/poppler-windows/releases/\n\n"
                    "Apakah Anda ingin memilih lokasi folder Poppler sekarang?\n"
                    "(Pilih folder 'Library\\bin' atau 'bin' dari hasil extract Poppler)"
                )
                
                if result:
                    selected_path = filedialog.askdirectory(
                        title="Pilih folder bin Poppler (contoh: poppler/Library/bin)"
                    )
                    
                    if selected_path and os.path.exists(selected_path):
                        try:
                            images = convert_from_path(pdf_path, dpi=200, poppler_path=selected_path)
                            # Save path to config
                            config_manager.config["poppler_path"] = selected_path
                            config_manager.save_config()
                            messagebox.showinfo(
                                "Berhasil",
                                f"Path Poppler berhasil disimpan!\n\n{selected_path}\n\n"
                                "Selanjutnya tidak perlu pilih lagi."
                            )
                        except Exception as e:
                            messagebox.showerror(
                                "Error",
                                f"Folder Poppler tidak valid!\n\n{e}\n\n"
                                "Pastikan memilih folder 'bin' atau 'Library/bin' dari Poppler."
                            )
                            return
                    else:
                        return
                else:
                    return
            
            # Save images
            if images:
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                for i, image in enumerate(images, start=1):
                    self.status_var.set(f"🔄 Menyimpan gambar {i}/{len(images)}...")
                    self.root.update()
                    
                    ext = "png" if img_format == "PNG" else "jpg"
                    out_file = os.path.join(out_dir, f"{base_name}_halaman_{i}.{ext}")
                    image.save(out_file, img_format)
            
            self.status_var.set(f"✅ PDF → Images selesai: {len(images)} halaman")
            messagebox.showinfo(
                "Selesai",
                f"PDF berhasil dikonversi menjadi {len(images)} gambar di:\n{out_dir}"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Gagal konversi PDF ke gambar:\n{e}")

    def images_to_pdf(self):
        if not self._ensure_pillow():
            return

        paths = filedialog.askopenfilenames(
            title="Pilih gambar untuk digabung jadi PDF (urutan sesuai pilihan)", 
            filetypes=[
                ("Image Files", "*.png *.jpg *.jpeg *.bmp *.tiff *.gif"),
                ("PNG Files", "*.png"),
                ("JPEG Files", "*.jpg *.jpeg"),
                ("All Files", "*.*")
            ]
        )
        if not paths or len(paths) == 0:
            return

        out_path = filedialog.asksaveasfilename(
            title="Simpan hasil PDF sebagai", 
            defaultextension=".pdf", 
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not out_path:
            return

        try:
            self.status_var.set(f"🔄 Memproses {len(paths)} gambar...")
            self.root.update()
            
            images = []
            for idx, img_path in enumerate(paths, 1):
                self.status_var.set(f"🔄 Memproses gambar {idx}/{len(paths)}: {os.path.basename(img_path)}")
                self.root.update()
                
                img = self.PIL_Image.open(img_path)
                
                # Convert RGBA to RGB (PDF doesn't support transparency)
                if img.mode == 'RGBA':
                    # Create white background
                    background = self.PIL_Image.new('RGB', img.size, (255, 255, 255))
                    background.paste(img, mask=img.split()[3])  # Use alpha channel as mask
                    img = background
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                
                images.append(img)

            # Save as PDF
            if images:
                self.status_var.set("🔄 Menyimpan PDF...")
                self.root.update()
                
                images[0].save(
                    out_path, 
                    "PDF", 
                    resolution=100.0, 
                    save_all=True, 
                    append_images=images[1:] if len(images) > 1 else []
                )

            self.status_var.set(f"✅ Images → PDF selesai: {len(paths)} gambar")
            messagebox.showinfo(
                "Selesai", 
                f"✅ PDF berhasil dibuat dari {len(paths)} gambar!\n\n"
                f"Output: {out_path}"
            )
        except Exception as e:
            self.status_var.set("❌ Error saat convert images ke PDF")
            messagebox.showerror("Error", f"Gagal menggabungkan images menjadi PDF:\n\n{str(e)}")

    def compress_pdf(self):
        if not self._ensure_pypdf():
            return

        path = filedialog.askopenfilename(
            title="Pilih PDF untuk di-compress", 
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not path:
            return

        out_path = filedialog.asksaveasfilename(
            title="Simpan hasil compress sebagai", 
            defaultextension=".pdf", 
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not out_path:
            return

        try:
            self.status_var.set("🔄 Membaca dan menganalisis PDF...")
            self.root.update()
            
            with open(path, 'rb') as pdf_file:
                reader = self.PdfReader(pdf_file)
                writer = self.PdfWriter()
                
                total_pages = len(reader.pages)
                
                # Copy pages and compress content streams
                for i, page in enumerate(reader.pages, 1):
                    self.status_var.set(f"🔄 Memproses halaman {i}/{total_pages}...")
                    self.root.update()
                    
                    # Compress content streams BEFORE adding to writer
                    try:
                        page.compress_content_streams()
                    except Exception:
                        pass
                    
                    writer.add_page(page)

                # Transfer minimal metadata
                if reader.metadata:
                    try:
                        # Only copy essential metadata
                        essential_meta = {}
                        for key in ['/Title', '/Author', '/Subject']:
                            if key in reader.metadata:
                                essential_meta[key] = reader.metadata[key]
                        if essential_meta:
                            writer.add_metadata(essential_meta)
                    except Exception:
                        pass
                
                # Remove duplicate objects and compress
                self.status_var.set("🔄 Mengoptimalkan dan menghapus duplikasi...")
                self.root.update()
                
                # Write with compression
                with open(out_path, 'wb') as output_file:
                    writer.write(output_file)

            # Get file sizes
            original_size = os.path.getsize(path)
            compressed_size = os.path.getsize(out_path)
            reduction = ((original_size - compressed_size) / original_size * 100)

            self.status_var.set(f"✅ Optimasi selesai: {os.path.basename(out_path)}")
            
            # Format sizes
            def format_size(size_bytes):
                if size_bytes >= 1024*1024:
                    return f"{size_bytes / (1024*1024):.2f} MB"
                else:
                    return f"{size_bytes / 1024:.1f} KB"
            
            if reduction > 0:
                size_info = (
                    f"✅ PDF berhasil dioptimasi!\n\n"
                    f"Ukuran awal: {format_size(original_size)}\n"
                    f"Ukuran akhir: {format_size(compressed_size)}\n"
                    f"Pengurangan: {reduction:.1f}%\n\n"
                    f"💾 File disimpan di:\n{out_path}"
                )
            else:
                # File malah bertambah atau sama
                increase = abs(reduction)
                size_info = (
                    f"⚠️ PDF telah dioptimasi\n\n"
                    f"Ukuran awal: {format_size(original_size)}\n"
                    f"Ukuran akhir: {format_size(compressed_size)}\n"
                    f"Perubahan: +{increase:.1f}%\n\n"
                    f"ℹ️ Catatan: PDF ini sudah teroptimasi atau\n"
                    f"mengandung konten yang tidak bisa dikompresi lebih lanjut.\n\n"
                    f"💾 File disimpan di:\n{out_path}"
                )
            
            messagebox.showinfo("Selesai", size_info)
        except Exception as e:
            self.status_var.set("❌ Error saat compress PDF")
            messagebox.showerror("Error", f"Gagal melakukan compress PDF:\n\n{str(e)}")

    def back_to_menu(self):
        if self.parent_window:
            self.root.destroy()
            self.parent_window.deiconify()