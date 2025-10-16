# ARSIP OWNCLOUD - Aplikasi Scan Folder

## 📝 Deskripsi

Aplikasi Python untuk scanning dan analisis folder arsip digital Owncloud dengan fitur:

- ✅ Validasi 8 folder standar Owncloud
- 📊 Export ke Excel dengan 8 sheet terpisah
- 📈 File counting dan size calculation
- 🔗 Hyperlink path yang bisa diklik di Excel
- 📋 Breakdown lengkap hingga level file

## 🚀 Cara Menjalankan

### Opsi 1: Jalankan dengan Python (RECOMMENDED)

```bash
# 1. Aktifkan virtual environment
.venv\Scripts\activate

# 2. Jalankan aplikasi
python main.py
```

### Opsi 2: Build ke EXE (Opsional)

> **Catatan**: Build EXE membutuhkan cleanup dependencies yang conflict

```bash
# Cara manual:
# 1. Buat virtual environment baru yang bersih
python -m venv venv_clean

# 2. Aktifkan
venv_clean\Scripts\activate

# 3. Install hanya dependencies yang diperlukan
pip install pandas openpyxl

# 4. Install PyInstaller
pip install pyinstaller

# 5. Build EXE
pyinstaller --onefile --windowed --name "ArsipOwncloud" main.py

# 6. File EXE ada di: dist\ArsipOwncloud.exe
```

## 📦 Dependencies

- Python 3.10+
- tkinter (built-in)
- pandas
- openpyxl

## 📁 Struktur Folder yang Didukung

1. `01.SURAT_MENYURAT` - MASUK/KELUAR → Tahun → Bulan
2. `02.DATA_ANGGOTA` - Center (4-digit) → ID_NAMA
3. `03.DATA_ANGGOTA_KELUAR` - Tahun → Bulan → ID_NAMA
4. `04.DATA_DANA_RESIKO` - Tahun → Bulan → ID_NAMA → File
5. `05.BUKU_HARI_RAYA_ANGGOTA` - Tahun → File bulanan
6. `06.LAPORAN_BULANAN` - Tahun → Bulan → 12 jenis dokumen
7. `07.BUKU_BANK` - Tahun → Bulan → DD_BUKUBANK.XLSX
8. `08.DATA_LWK` - Tahun → Bulan → DD_CCCC.PDF

## ✨ Fitur Excel Export

- **8 Sheet Terpisah**: Satu sheet per folder standar
- **Kolom PATH**: Hyperlink yang bisa diklik untuk buka folder/file
- **Status Indicators**: FOLDER/FILE dengan ukuran (KB)
- **Auto-formatting**: Center 4-digit, tanggal, kode dokumen

## 💡 Tips

- Gunakan tombol "📋 Export Struktur Lengkap" untuk export komprehensif
- Klik path di Excel untuk langsung buka lokasi file
- Scan dilakukan rekursif hingga semua level folder

## 🔧 Troubleshooting

### "PyInstaller build gagal"

- Gunakan virtual environment bersih (tanpa torch/tensorflow)
- Atau jalankan langsung dengan `python main.py`

### "Module not found"

```bash
pip install pandas openpyxl
```

### "Tkinter error"

- Tkinter sudah built-in di Python Windows installer
- Reinstall Python dengan opsi "tcl/tk" dicentang

## 📞 Support

Untuk pertanyaan atau issue, silakan hubungi tim IT MIS.

---

**Version**: 1.0  
**Last Updated**: Oktober 2025
