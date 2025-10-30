# 📁 ARSIP OWNCLOUD - Aplikasi Manajemen Arsip Digital

> Aplikasi desktop berbasis Python untuk manajemen, scanning, dan analisis arsip digital Owncloud

**Versi**: 1.1.6 🆕
**Tanggal**: Oktober 2025
**Developer**: Riky Dwianto

---

## 🎉 What's New in v1.1.6

### **Major Refactoring** - Modular Structure

✅ **Code Organization**: Main.py dipecah menjadi 8 modul terpisah  
✅ **Maintainability**: Struktur kode lebih mudah dipelihara  
✅ **Arsip Category**: Semua fitur arsip digabung dalam `app_arsip.py`  
✅ **Zero Breaking Changes**: Semua fungsi tetap sama

� **Detail**: Lihat `REFACTORING_SUCCESS.md` untuk informasi lengkap

---

## �📋 Daftar Isi

1. [Deskripsi Umum](#-deskripsi-umum)
2. [Struktur Kode (NEW)](#-struktur-kode-new-)
3. [Instalasi & Setup](#-instalasi--setup)
4. [Menu Utama](#-menu-utama)
5. [Fitur Detail](#-fitur-detail)
   - [Cek Arsip Digital](#1-cek-arsip-digital)
   - [Scan Folder Arsip Digital](#2-scan-folder-arsip-digital)
   - [Universal Scan Database](#3-universal-scan-database)
   - [Scan File Besar](#4-scan-file-besar)
   - [Cek Pengajuan Dana](#5-cek-pengajuan-dana)
   - [Cek NO KK](#6-cek-no-kk)
   - [PDF Tool](#7-pdf-tool)
   - [Pengaturan](#8-pengaturan)
6. [Dependencies](#-dependencies)
7. [Build Executable](#-build-executable)
8. [Tips & Best Practices](#-tips--best-practices)
9. [Troubleshooting](#-troubleshooting)
10. [FAQ](#-faq)

---

## 🎯 Deskripsi Umum

Aplikasi **ARSIP OWNCLOUD** adalah sistem manajemen arsip digital yang dirancang untuk membantu organisasi dalam:

- ✅ **Validasi Struktur Arsip** - Memastikan folder arsip sesuai standar
- 📊 **Export ke Excel** - Membuat laporan lengkap dengan hyperlink
- 🔍 **Scanning File** - Mencari file besar dan format non-standar
- 💰 **Tracking Pengajuan Dana** - Inventarisasi dokumen pengajuan dana
- ⚙️ **Konfigurasi Fleksibel** - Default folder untuk efisiensi kerja

### Fitur Utama

| Fitur                         | Deskripsi                                          |
| ----------------------------- | -------------------------------------------------- |
| **Cek Arsip Digital**         | Matching data folder dengan database Excel anggota |
| **Scan Folder Arsip Digital** | Validasi 8 folder standar dengan export detail     |
| **Universal Scan Database**   | Scan komprehensif dengan multiple sheet output     |

---

## 📦 Struktur Kode (NEW) 🆕

Aplikasi sekarang menggunakan **modular structure** untuk maintainability yang lebih baik:

```
ARSIPOWNCLOUD/
├── main.py                  # 🎯 Entry point & Main Menu (300 lines)
├── app_helpers.py           # 🛠️ Helper functions & ConfigManager
├── app_settings.py          # ⚙️ Settings & web server
├── app_kk_checker.py        # 👨‍👩‍👧‍👦 NO KK validation with OCR
├── app_dana_checker.py      # 💰 Pengajuan Dana checker
├── app_pdf_tools.py         # 📃 PDF tools (merge/split/OCR)
├── app_arsip.py             # 📋 3 Arsip forms (grouped)
│   ├── ArsipDigitalApp      #    - Cek Arsip Digital
│   ├── ScanFolderApp        #    - Scan Folder Arsip
│   └── UniversalScanApp     #    - Universal Scan
├── app_scan_files.py        # 📊 Large file scanner
├── arsip_logic.py           # 🧠 Business logic
└── web_server.py            # 🌐 Web server
```

**Benefits**:

- ✅ **96% smaller main.py** (7,628 → 300 lines)
- ✅ **Easy to find code** - Each module has clear purpose
- ✅ **Better collaboration** - Multiple devs can work simultaneously
- ✅ **Faster development** - Isolated testing and debugging

📖 **Learn More**: `REFACTORING_SUCCESS.md`
| **Scan File Besar** | Deteksi file berukuran besar dengan threshold kustom |
| **Cek Pengajuan Dana** | Scan otomatis file PENGAJUAN_DANA.xlsm multi-tahun |
| **Cek NO KK** | Ekstrak dan validasi Nomor Kartu Keluarga dari PDF dengan OCR |
| **PDF Tool** | Konversi Images↔PDF, Merge, Split, Compress PDF |
| **Pengaturan** | Simpan default folder untuk semua form |

---

## 🚀 Instalasi & Setup

### Prasyarat

- Python 3.10 atau lebih baru
- Windows OS (untuk `os.startfile()`)
- Git (opsional, untuk clone repository)

### Langkah Instalasi

#### 1. Clone Repository

```bash
git clone https://github.com/rikydwianto/cekarsipdigital.git
cd cekarsipdigital
```

#### 2. Buat Virtual Environment

```bash
python -m venv .venv
```

#### 3. Aktivasi Virtual Environment

```bash
# Windows (PowerShell)
.venv\Scripts\activate

# Windows (CMD)
.venv\Scripts\activate.bat
```

#### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

Atau install manual:

```bash
pip install pandas openpyxl
```

#### 5. Jalankan Aplikasi

```bash
python main.py
```

### Menjalankan dari Executable (Opsional)

Jika sudah di-build ke `.exe`:

```bash
# Jalankan dari folder build
cd ArsipOwncloud_Portable
ArsipOwncloud.exe
```

---

## 💾 Lokasi Penyimpanan File

### File Data Aplikasi

Aplikasi menyimpan file database dan export di **AppData Local** sesuai Windows best practice:

```
C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\
├── database.xlsx                    # Database hasil scan (Arsip Digital)
├── file_export.xlsx                 # File export matching data
├── app_config.json                  # Konfigurasi aplikasi
└── universal_scan_database.xlsx     # Database universal scan
```

### Cara Akses Folder AppData

**Metode 1: Keyboard Shortcut**

1. Tekan `Win + R`
2. Ketik: `%LOCALAPPDATA%\ArsipDigitalOwnCloud`
3. Tekan Enter

**Metode 2: File Explorer**

1. Buka File Explorer
2. Enable "Show hidden files" di View options
3. Navigate ke: `C:\Users\[YourUsername]\AppData\Local\ArsipDigitalOwnCloud`

### Mengapa AppData?

✅ **Best Practice Windows**: Aplikasi tidak menulis ke Program Files
✅ **Tidak Perlu Admin**: User biasa bisa write tanpa elevated privileges
✅ **User Isolation**: Setiap Windows user punya data sendiri
✅ **Mudah Backup**: Terbackup otomatis dengan Windows Backup
✅ **Kompatibilitas**: Sesuai standar modern Windows applications

> 📝 **Note**: File `database.xlsx` dan `file_export.xlsx` dibuat otomatis saat scan pertama kali. Aplikasi akan menampilkan full path di error message jika file tidak ditemukan.

---

## 🏠 Menu Utama

Aplikasi memiliki menu utama dengan 8 tombol dalam layout 2 kolom:

```
┌─────────────────────────────────────┐
│      📁 APLIKASI ARSIP DIGITAL      │
│   Sistem Manajemen Arsip Digital   │
├─────────────────────────────────────┤
│  📋 Cek Arsip Digital              │
│  📂 Scan Folder Arsip Digital      │
│  🌐 Universal Scan Database        │
│  📊 Scan File Besar                │
│  💰 Cek Pengajuan Dana             │
│  👨‍👩‍👧‍👦 Cek NO KK                    │
│  📃 PDF Tool                       │
│  ⚙️  Pengaturan                     │
├─────────────────────────────────────┤
│           [Keluar]                  │
│     v1.1.5 - Developed by RD        │
└─────────────────────────────────────┘
```

---

## 📖 Fitur Detail

### 1. Cek Arsip Digital

**Fungsi**: Mencocokkan struktur folder data anggota dengan database Excel

#### Cara Menggunakan

1. **Pilih Folder Data Anggota**

   - Browse ke folder yang berisi data anggota
   - Bisa single anggota, center, atau root folder

2. **Pilih File Excel Database**

   - Browse file Excel dengan format header B3-Y3
   - File berisi data anggota lengkap

3. **Proses Arsip**

   - Klik "Proses Arsip"
   - Lihat preview matching
   - Pilih export option

#### Output

- **file_export.xlsx** dengan 4 sheet:
  - `databaseanggota` - Data dari Excel
  - `hasilscan` - Hasil scan folder
  - `datamatching` - Data yang match (ada di database DAN folder)
  - `belumdiarsip` - Data belum diarsip (ada di database tapi TIDAK ada di folder)

#### Fitur Matching

- Matching berdasarkan **Center + ID Anggota**
- Normalisasi ID ke format 6 digit
- Normalisasi Center ke format 4 digit
- Statistik lengkap jumlah match dan belum diarsip

---

### 2. Scan Folder Arsip Digital

**Fungsi**: Validasi struktur folder arsip sesuai 8 standar Owncloud

#### 8 Folder Standar

1. **01.SURAT_MENYURAT**

   - `01.SURAT_MASUK` → Tahun → Bulan
   - `02.SURAT_KELUAR` → Tahun → Bulan

2. **02.DATA_ANGGOTA**

   - Center (4-digit) → `IDIDID_NAMA`

3. **03.DATA_ANGGOTA_KELUAR**

   - Tahun → Bulan → `IDIDID_NAMA`

4. **04.DATA_DANA_RESIKO**

   - Tahun → Bulan → `IDIDID_NAMA` → File

5. **05.DATA_HARI_RAYA_ANGGOTA**

   - Tahun → File bulanan

6. **06.LAPORAN_BULANAN**

   - Tahun → Bulan → 12 jenis dokumen

7. **07.BUKU_BANK**

   - Tahun → Bulan → `DD_BUKUBANK.XLSX`

8. **08.DATA_LWK**

   - Tahun → Bulan → `DD_CCCC.PDF`

#### Export Excel

- **8 Sheet Terpisah** - Satu sheet per folder
- **Hyperlink PATH** - Klik untuk buka folder/file
- **Status Indicators** - FOLDER/FILE dengan ukuran
- **Auto-formatting** - Format angka dan tanggal otomatis

---

### 3. Universal Scan Database

**Fungsi**: Scan komprehensif struktur folder dan buat database lengkap dengan multiple sheet

#### Fitur Utama

- **Scan Folder Lengkap** - Deteksi otomatis semua folder dan file
- **Database Excel** - Output ke `database.xlsx` di AppData
- **Multiple Sheet** - Setiap folder punya sheet sendiri
- **Auto-detect Structure** - Tidak perlu config manual
- **Export Option** - Simpan dan Sinkron ke AppData

#### Cara Kerja

1. **Pilih Folder Root**

   - Browse ke folder arsip utama
   - Bisa root, center, atau anggota

2. **Proses Scan**

   - Deteksi struktur folder otomatis
   - Scan file di setiap subfolder
   - Collect metadata (nama, ukuran, path, type)

3. **Generate Database**
   - Buat `database.xlsx` di AppData
   - Lokasi: `C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\database.xlsx`
   - Multiple sheet sesuai struktur folder

#### Output Database Structure

**Sheet per Folder Type:**

- `01.SURAT_MENYURAT` - Data surat masuk/keluar
- `02.DATA_ANGGOTA` - Data anggota per center
- `03.DATA_ANGGOTA_KELUAR` - Data anggota keluar
- `04.DATA_DANA_RESIKO` - Data dana resiko
- `05.DATA_HARI_RAYA_ANGGOTA` - Data hari raya
- `06.LAPORAN_BULANAN` - Laporan bulanan
- `07.BUKU_BANK` - Data buku bank
- `08.DATA_LWK` - Data LWK

**Kolom di Setiap Sheet:**

- NOMOR_CENTER
- ID_NAMA_ANGGOTA
- NAMA_FILE
- TYPE (FOLDER/FILE)
- UKURAN_KB
- PATH

#### Integrasi dengan Fitur Lain

Database yang dibuat akan digunakan oleh:

- ✅ **Cek NO KK** - Baca list file PDF dari sheet `02.DATA_ANGGOTA`
- ✅ **Web Server** - API data anggota
- ✅ **Export Matching** - Matching dengan hasil scan

#### Use Cases

1. **Database Master** - Buat database lengkap untuk seluruh arsip
2. **Quick Scan** - Scan cepat tanpa perlu setup detail
3. **Data Integration** - Sumber data untuk fitur lain
4. **Inventory** - Inventarisasi lengkap file dan folder

---

### 4. Scan File Besar

**Fungsi**: Mencari file berdasarkan ukuran atau format file

#### 🔍 Mode 1: File Besar

Mencari file berukuran ≥ threshold tertentu (default: 10 MB)

**Parameter:**

- Ukuran minimum (MB) - dapat dikonfigurasi
- Filter file owncloud sync
- Treeview dengan kolom Ekstensi

**Use Cases:**

- Cleanup file besar untuk hemat storage
- Identifikasi file media untuk dipindah ke cold storage
- Audit penggunaan disk space

#### 📄 Mode 2: Format Non-Dokumen

Mencari file dengan ekstensi TIDAK standar

**Format Dokumen yang DIIZINKAN (ekstensi):**

- **Office**: `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`, `.odt`, `.ods`, `.odp`
- **PDF & Text**: `.pdf`, `.txt`, `.rtf`, `.csv`
- **File yang AKAN DITEMUKAN:**
- `.exe`, `.dll`, `.sys` - executables
- `.db`, `.sqlite`, `.mdb` - database files
- `.log`, `.tmp`, `.bak` - temporary files
- `.iso`, `.dmg`, `.img` - disk images
- File tanpa ekstensi

**Use Cases:**

- Audit keamanan (cari file `.exe`, `.dll`)
- Cari file database/backup (`.db`, `.sqlite`, `.bak`)
- Identifikasi file tidak standar

#### Export Excel

Kolom yang di-export:

- No
- Nama File
- **Ekstensi** (kolom baru!)
- Ukuran (MB)
- Ukuran (Bytes)
- Path Lengkap

#### Filter Otomatis

File berikut **SELALU DIABAIKAN**:

- `.owncloudsync.log`
- `.owncloudsync.log.1`
- `.sync_journal.db`
- `.sync_journal.db-wal`

---

### 5. Cek Pengajuan Dana

**Fungsi**: Scan dan inventarisasi file PENGAJUAN_DANA.xlsm dari Surat Keluar

#### Struktur Folder

```
${default_folder}/
└── 01.SURAT_MENYURAT/
    └── 02.SURAT_KELUAR/
        ├── 2020/
        │   ├── 01.JANUARI/
        │   │   ├── 001_PENGAJUAN_DANA.xlsm
        │   │   ├── 002_PENGAJUAN_DANA.xlsm
        │   └── 02.FEBRUARI/
        ├── 2021/
        ├── 2022/
        └── 2025/
            ├── 01.JANUARI/
            │   ├── 001_PENGAJUAN_DANA.xlsm
            │   └── 005_PENGAJUAN_DANA.xlsm
            └── 02.FEBRUARI/
```

#### Konvensi Penamaan

- **Format**: `{nomor_surat}_PENGAJUAN_DANA.xlsm`
- **Nomor Surat**: 3 digit (001, 025, 100)
- **Case-insensitive**: Terdeteksi baik huruf besar maupun kecil

#### Algoritma Scan

```python
1. Baca default_folder dari konfigurasi
2. Validasi path: {default_folder}/01.SURAT_MENYURAT/02.SURAT_KELUAR
3. Loop tahun: 2020 s/d (current_year + 1)
   Loop bulan: 01.JANUARI s/d 12.DESEMBER
   Scan file: *_PENGAJUAN_DANA.xlsm
   Filter: Skip temporary files (~*.xlsm)
4. Tampilkan hasil dengan 6 kolom:
   - No, Tahun, Bulan, Nomor Surat, Nama File, Path
5. Enable Export button dan Analisa Data button
```

#### Fitur Analisa Data 🔬

**Ekstraksi Data dari File Excel**

Tombol **🔬 Analisa Data** memungkinkan Anda mengambil data dari dalam setiap file PENGAJUAN_DANA.xlsm:

**Data yang Diekstrak:**

| No  | Data Field                  | Lokasi         | Cell | Keterangan                       |
| --- | --------------------------- | -------------- | ---- | -------------------------------- |
| 1   | **Nomor Surat**             | Sheet Surat    | F8   | Nomor surat dari dalam file      |
| 2   | **Nominal Input Kebutuhan** | Sheet Surat    | I8   | Nominal kebutuhan input          |
| 3   | **Nominal Kebutuhan**       | Sheet Laporan  | F68  | Total nominal kebutuhan          |
| 4   | **Status Balance**          | Sheet Laporan  | A4   | Status balance (BALANCE/SELISIH) |
| 5   | **Tanggal Disburse Awal**   | Sheet Lampiran | C3   | Tanggal mulai disburse           |
| 6   | **Tanggal Disburse Akhir**  | Sheet Lampiran | E3   | Tanggal akhir disburse           |
| 7   | **Nama BM**                 | Sheet Laporan  | A83  | Nama Branch Manager              |

**Detail Ekstraksi:**

- **Nomor Surat (F8)**: String - Nomor surat dari dalam file Excel
- **Nominal Input (I8)**: Number - Nominal yang diinput untuk kebutuhan
- **Nominal Kebutuhan (F68)**: Number - Total nominal kebutuhan dari laporan
- **Status Balance (A4)**: String - Extract text setelah "Ket. :"
  - Contoh: "Ket. : BALANCE" → ambil "BALANCE"
  - Contoh: "Ket. : NIHIL" → ambil "NIHIL"
- **Tanggal Disburse Awal (C3)**: Date - Tanggal mulai pencairan dana
- **Tanggal Disburse Akhir (E3)**: Date - Tanggal akhir pencairan dana
- **Nama BM (A83)**: String - Nama Branch Manager

**Proses Analisa:**

1. Klik tombol **🔬 Analisa Data** (aktif setelah scan)
2. Konfirmasi jumlah file yang akan dianalisa
3. Progress dialog menampilkan file yang sedang diproses
4. Treeview diupdate menjadi **14 kolom**:
   - No, Tahun, Bulan, Nomor Surat (Nama File)
   - Nomor di File (F8), Nominal Input (I8), Nominal Kebutuhan (F68)
   - Status Balance (A4), Tanggal Disburse Awal (C3), Tanggal Disburse Akhir (E3)
   - Nama BM (A83), Status, Nama File, Path
5. Status indikator: ✅ (sukses) atau ❌ (error)

**Handling Error:**

- File tanpa sheet 'Surat', 'Laporan', atau 'Lampiran' → Status ❌, data kosong
- Cell tidak ada atau kosong → value = None, tampil sebagai "-"
- Error pada satu file tidak mengganggu file lainnya
- Summary menampilkan jumlah sukses dan error

#### Quick Open Feature 🖱️

**💡 Tip: Double-click pada baris untuk membuka file Excel**

- Klik 2x pada row → file langsung terbuka di Excel
- Validasi otomatis jika file tidak ada
- Lebih cepat dari copy-paste path
- Bekerja dengan 6 kolom (sebelum analisa) atau 14 kolom (setelah analisa)

#### Export Excel

**Export Tanpa Analisa:**

- Sheet name: "Pengajuan Dana"
- Kolom: Tahun, Bulan, Kode Bulan, Nomor Surat, Nama File, Path Lengkap

**Export Dengan Analisa (14 Kolom):**

- Sheet name: "Pengajuan Dana"
- Kolom Export:
  1. Tahun
  2. Bulan
  3. Kode Bulan
  4. Nomor Surat (Nama File)
  5. Nomor Surat (F8)
  6. Nominal Input Kebutuhan (I8)
  7. Nominal Kebutuhan (F68)
  8. Status Balance (A4)
  9. Tanggal Disburse Awal (C3)
  10. Tanggal Disburse Akhir (E3)
  11. Nama BM (A83)
  12. Status Analisa
  13. Nama File
  14. Path Lengkap

**Filename**: `pengajuan_dana_YYYYMMDD_HHMMSS.xlsx`

#### Use Cases

1. **Inventarisasi Tahunan** - Lihat semua pengajuan dana 2024
2. **Audit Bulanan** - Cek pengajuan dana bulan tertentu
3. **Tracking Nomor Surat** - Cari file berdasarkan nomor surat
4. **Quick Access** - Buka file dengan double-click
5. **Backup/Migrasi** - Export daftar lengkap untuk dokumentasi

---

### 6. Cek NO KK

**Fungsi**: Ekstrak dan validasi Nomor Kartu Keluarga (NO KK) dari file PDF menggunakan OCR

#### Prasyarat

- **Tesseract OCR** harus terinstall di sistem
  - Download: https://github.com/UB-Mannheim/tesseract/wiki
  - Lokasi default: `C:\Program Files\Tesseract-OCR\tesseract.exe`
- **Poppler** (sudah dibundle dengan aplikasi)
- **Database**: File `database.xlsx` harus sudah ada (hasil scan folder arsip)

#### Cara Kerja

1. **Membaca Database**

   - Lokasi: `C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\database.xlsx`
   - Sheet: `02.DATA_ANGGOTA`
   - Filter: File PDF yang diawali dengan `02` (Data Kartu Keluarga)

2. **Ekstraksi OCR**

   - Konversi PDF ke Image (DPI 400)
   - Crop header 20% (fokus ke area NO KK)
   - Deskew otomatis (koreksi rotasi -20° hingga +20°)
   - Enhance kontras 3x dan ketajaman 2x
   - Binarisasi untuk OCR optimal
   - OCR dengan digit whitelist only

3. **Koreksi Karakter**

   - `b` → `6`, `B` → `8`
   - `O`, `o` → `0`
   - `l`, `I` → `1`
   - `S` → `5`, `Z` → `2`

4. **Validasi Format**
   - Panjang: 16 digit
   - Isi: Angka semua
   - Handle konversi 17→16 digit

#### Fitur

- ⏸️ **Pause/Resume** - OCR lambat? Pause dulu!
- 📊 **Treeview 10 Kolom**:
  1. No
  2. NO KK (hasil ekstraksi)
  3. Status (Valid/Invalid)
  4. Panjang (jumlah digit)
  5. Format (Numerik/Non-Numerik)
  6. Keterangan (detail error)
  7. Nama Anggota
  8. Nomor Center
  9. Status File (✅ Ada / ❌ Tidak Ada)
  10. Path
- 📥 **Export ke Excel** dengan 2 sheet:
  - `Data` - Hasil pengecekan lengkap
  - `Summary` - Statistik (Total, Valid, Invalid, %)

#### Status File

- ✅ **Ada** - File PDF ditemukan, OCR dijalankan
- ❌ **Tidak Ada** - File tidak ditemukan di PATH
- 🚫 **NO KK Tidak Ditemukan** - OCR gagal ekstrak NO KK

#### Use Cases

1. **Validasi Data KK** - Pastikan semua NO KK valid 16 digit
2. **Cek File Missing** - Identifikasi file yang hilang
3. **Quality Control** - Audit kualitas data sebelum upload
4. **Data Correction** - Export list NO KK invalid untuk perbaikan

#### Tips

- **File Missing**: Status File akan otomatis tracking file yang tidak ada
- **OCR Gagal**: Check kualitas scan PDF (resolusi minimal 300 DPI)
- **Proses Lama**: Gunakan Pause untuk istirahat, lanjut Resume
- **Rotasi**: Deskew otomatis handle KK yang miring hingga ±20°

---

### 7. PDF Tool

**Fungsi**: Toolbox lengkap untuk manipulasi file PDF

#### Fitur Utama (Layout 2 Kolom)

**Kolom Kiri:**

1. 🖼️ **Images to PDF** - Gabung multiple gambar jadi 1 PDF
2. 📄 **PDF to Images** - Extract semua halaman PDF ke gambar
3. ➕ **Merge PDF** - Gabung multiple PDF jadi 1 file
4. ✂️ **Split PDF** - Pisah PDF per halaman atau range

**Kolom Kanan:** 5. 🗜️ **Compress PDF** - Kurangi ukuran PDF 6. 🔄 **Rotate PDF** - Putar halaman PDF (90°, 180°, 270°) 7. 🔐 **Protect PDF** - Password protect PDF 8. 🔓 **Unlock PDF** - Hapus password dari PDF

#### Detail Fitur

##### 1. 🖼️ Images to PDF

- **Input**: Multiple gambar (JPG, PNG, BMP, TIFF, GIF)
- **Output**: 1 file PDF
- **Fitur**:
  - Drag & drop urutan gambar
  - Preview image
  - Custom page size (A4, Letter, Legal, atau Custom)
  - Quality control

##### 2. 📄 PDF to Images

- **Input**: 1 file PDF
- **Output**: Multiple gambar (PNG default)
- **Fitur**:
  - Pilih format output (PNG, JPG, TIFF, BMP)
  - Custom DPI (72-600)
  - Extract all pages atau specific range
  - Output folder selection
- **Membutuhkan**: Poppler (sudah dibundle)

##### 3. ➕ Merge PDF

- **Input**: Multiple file PDF
- **Output**: 1 file PDF merged
- **Fitur**:
  - Drag & drop untuk urutan
  - Preview setiap PDF
  - Bookmark otomatis per file

##### 4. ✂️ Split PDF

- **Input**: 1 file PDF
- **Output**: Multiple PDF
- **Mode**:
  - **Per Halaman**: 1 file = 1 halaman
  - **Range**: Tentukan range halaman (e.g., 1-5, 10-15)
  - **Custom Split**: Batch split dengan pattern

##### 5. 🗜️ Compress PDF

- **Input**: 1 file PDF
- **Output**: PDF terkompress
- **Level Kompresi**:
  - Low (kualitas tinggi, kompresi rendah)
  - Medium (balanced)
  - High (kualitas rendah, kompresi tinggi)
- **Preview**: Tampilkan ukuran before/after

##### 6. 🔄 Rotate PDF

- **Input**: 1 file PDF
- **Rotasi**: 90°, 180°, 270° (clockwise)
- **Mode**:
  - All Pages
  - Specific Range
  - Odd/Even Pages Only

##### 7. 🔐 Protect PDF

- **Input**: 1 file PDF
- **Output**: PDF dengan password
- **Password**: User password & Owner password
- **Permissions**: Set permission (print, copy, modify)

##### 8. 🔓 Unlock PDF

- **Input**: PDF yang diprotect
- **Requirement**: Masukkan password
- **Output**: PDF tanpa password

#### Dependencies

- **PyPDF2**: Merge, Split, Rotate, Protect, Unlock
- **Pillow**: Images to PDF
- **pdf2image + Poppler**: PDF to Images
- **ReportLab**: PDF generation untuk Images to PDF

#### Tips & Tricks

- **Compress**: Coba Medium dulu sebelum High
- **PDF to Images**: DPI 300 untuk print quality, 150 untuk web
- **Merge PDF**: Urutan penting! Cek preview sebelum merge
- **Poppler**: Sudah dibundle, tidak perlu install terpisah

---

### 8. Pengaturan

**Fungsi**: Konfigurasi default folder untuk semua form

#### Fitur

- **Set Default Folder** - Pilih folder yang sering digunakan
- **Auto-Load** - Semua browse dialog langsung ke folder ini
- **Hapus Default** - Reset ke current directory
- **Persistent Storage** - Disimpan di AppData `app_config.json`

#### Form yang Mendukung

1. ✅ Cek Arsip Digital
2. ✅ Scan Folder Arsip Digital
3. ✅ Scan File Besar

#### Cara Menggunakan

1. **Set Default Folder**

   ```
   Menu → ⚙️ Pengaturan → 📂 Pilih Folder Default
   ```

2. **Test di Form Lain**

   - Buka salah satu form
   - Klik "Browse Folder"
   - Dialog otomatis ke default folder

3. **Hapus Default** (jika diperlukan)

   ```
   Menu → ⚙️ Pengaturan → 🗑️ Hapus Default
   ```

#### File Konfigurasi

**app_config.json** (auto-generated di AppData):

```json
{
  "default_folder": "D:\\Data_Anggota_Owncloud",
  "web_server_enabled": false,
  "web_server_port": 1212
}
```

**Lokasi**: `C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\app_config.json`

**Security**:

- File ada di `.gitignore`
- Tidak terbawa ke Git repository
- Setiap komputer punya config sendiri

---

## 📂 Struktur Folder

### Workspace Structure

```
ARSIPOWNCLOUD/
├── main.py                    # Aplikasi utama (Main Menu)
├── app_helpers.py             # Helper functions & ConfigManager
├── app_settings.py            # Settings application
├── app_kk_checker.py          # KK checker with OCR
├── app_dana_checker.py        # Dana checker
├── app_pdf_tools.py           # PDF manipulation tools
├── app_arsip.py               # Arsip applications (3 classes)
├── app_scan_files.py          # Large file scanner
├── arsip_logic.py             # Business logic
├── web_server.py              # Web server for remote access
├── requirements.txt           # Dependencies
├── README.md                  # Dokumentasi (file ini)
├── .gitignore                 # Git ignore rules
├── .venv/                     # Virtual environment
├── build/                     # Build artifacts
├── src_web/                   # Web interface templates
└── ArsipOwncloud_Portable/   # Portable executable
```

**Data Files** (auto-generated di AppData):
- `app_config.json` - Konfigurasi aplikasi
- `database.xlsx` - Database hasil scan
- `file_export.xlsx` - File export matching data
- `universal_scan_database.xlsx` - Database universal scan

### Config Files

| File                             | Deskripsi                      | Lokasi   |
| -------------------------------- | ------------------------------ | -------- |
| `app_config.json`                | Konfigurasi default folder     | AppData  |
| `database.xlsx`                  | Database hasil scan            | AppData  |
| `file_export.xlsx`               | File export hasil scan         | AppData  |
| `universal_scan_database.xlsx`   | Database universal scan        | AppData  |
| `requirements.txt`               | Python dependencies            | Project  |
| `.gitignore`                     | Git ignore rules               | Project  |

---

## 📦 Dependencies

### Runtime Dependencies

```txt
# Core
pandas>=2.3.3
openpyxl>=3.1.5
numpy>=2.2.6

# PDF & OCR
PyPDF2>=3.0.1
pdf2image>=1.17.0
pytesseract>=0.3.10
Pillow>=12.0.0

# QR Code
qrcode>=8.0
```

### External Dependencies

| Tool              | Purpose                  | Status      | Download                                                  |
| ----------------- | ------------------------ | ----------- | --------------------------------------------------------- |
| **Poppler**       | PDF to Images conversion | ✅ Bundled  | Included in build                                         |
| **Tesseract OCR** | OCR engine for Cek NO KK | ⚠️ Required | [Download](https://github.com/UB-Mannheim/tesseract/wiki) |

**Note:**

- Poppler sudah dibundle dengan aplikasi di `poppler-25.07.0/`
- Tesseract OCR harus diinstall manual untuk fitur Cek NO KK
- Lokasi default Tesseract: `C:\Program Files\Tesseract-OCR\tesseract.exe`

### Built-in Modules

- `tkinter` - GUI framework (built-in Python Windows)
- `os`, `sys`, `json` - Standard library
- `datetime` - Date/time handling

### Development Dependencies (Opsional)

```txt
cx-Freeze==6.15.0     # Build EXE
pyinstaller==5.13.0   # Alternative build tool
```

### Install All

```bash
pip install -r requirements.txt
```

---

## 🔨 Build Executable

### Persiapan Build dengan Poppler (untuk PDF → Images)

**Penting!** Jika ingin fitur PDF → Images bekerja di exe, siapkan folder Poppler:

```bash
# 1. Download Poppler dari:
# https://github.com/oschwartz10612/poppler-windows/releases/

# 2. Extract dan letakkan di root project dengan nama: poppler-25.07.0
# Struktur folder harus seperti ini:
ARSIPOWNCLOUD/
├── main.py
├── poppler-25.07.0/           <-- Folder ini
│   └── Library/
│       └── bin/
│           ├── pdftoppm.exe
│           ├── pdfimages.exe
│           └── ...
└── ...

# 3. Folder ini akan otomatis di-bundle saat build!
```

### Opsi 1: cx_Freeze (RECOMMENDED)

```bash
# 1. Install cx_Freeze
pip install cx-Freeze

# 2. Build portable version
.\build_cxfreeze.bat

# Output: build\exe.win-amd64-3.10\ArsipOwncloud.exe
# Termasuk: src_web, app_config.json, poppler-25.07.0 (jika ada)
```

**Keuntungan cx_Freeze:**

- ✅ Compatible dengan pandas & numpy
- ✅ Portable folder (bisa di-copy)
- ✅ Include semua dependencies
- ✅ **Auto-bundle Poppler** untuk PDF → Images

**Isi Build:**

- `src_web/` → Web server templates & static files
- `app_config.json` → Default configuration
- `poppler-25.07.0/` → PDF to Images converter (jika folder ada)

### Opsi 2: PyInstaller

```bash
# 1. Buat virtual environment bersih
python -m venv venv_clean
venv_clean\Scripts\activate

# 2. Install dependencies minimal
pip install pandas openpyxl pyinstaller

# 3. Build
pyinstaller --onefile --windowed --name "ArsipOwncloud" main.py

# Output: dist\ArsipOwncloud.exe
```

**Catatan PyInstaller:**

- ⚠️ Mungkin conflict dengan numpy/torch
- ⚠️ Perlu cleanup dependencies
- ⚠️ Poppler tidak auto-bundle (perlu manual add)

### Build Scripts

| Script               | Deskripsi                                |
| -------------------- | ---------------------------------------- |
| `build_cxfreeze.bat` | Build dengan cx_Freeze + Poppler (BEST)  |
| `build_portable.bat` | Build dengan cx_Freeze (portable folder) |
| `build_simple.bat`   | Build sederhana cx_Freeze                |
| `build_exe.bat`      | Build dengan PyInstaller                 |

---

## 💡 Tips & Best Practices

### 1. Konsistensi Penamaan File

✅ **BENAR:**

```
001_PENGAJUAN_DANA.xlsm
025_PENGAJUAN_DANA.xlsm
100_PENGAJUAN_DANA.xlsm
```

❌ **SALAH:**

```
1_PENGAJUAN_DANA.xlsm      ← Harus 3 digit
001_pengajuan_dana.xlsm    ← Huruf kecil (masih terdeteksi)
001_DANA.xlsm              ← Tidak ada kata PENGAJUAN_DANA
```

### 2. Struktur Folder Standar

- Gunakan format bulan: `01.JANUARI`, `02.FEBRUARI`
- Folder tahun: `2025` (bukan `TAHUN_2025`)
- ID Anggota: 6 digit zero-padded
- Center Code: 4 digit zero-padded

### 3. Export Regular

- Export hasil scan secara berkala
- Gunakan timestamp pada nama file
- Simpan di lokasi backup yang aman

### 4. Penggunaan Default Folder

- Set default folder di menu Pengaturan
- Hemat waktu browse berulang kali
- Update jika ganti project

### 5. Scan File Besar

**Ukuran Minimum yang Disarankan:**

- **5 MB** - Folder kecil/personal
- **10 MB** - Default, cocok untuk kebanyakan kasus
- **50 MB** - Folder besar, fokus file sangat besar
- **100 MB** - Server/enterprise storage

### 6. Quick Access

- **Double-click** pada row untuk buka file
- Lebih cepat dari copy-paste path
- Validasi otomatis jika file tidak ada

---

## 🔧 Troubleshooting

### Build Executable Gagal

**Problem**: PyInstaller build error dengan pandas/numpy

**Solusi**:

1. Gunakan cx_Freeze sebagai alternatif
2. Atau buat venv bersih tanpa torch/tensorflow
3. Atau jalankan langsung dengan `python main.py`

### Module Not Found

**Problem**: `ModuleNotFoundError: No module named 'pandas'`

**Solusi**:

```bash
pip install pandas openpyxl
```

### Tkinter Error

**Problem**: `_tkinter.TclError` atau tkinter not found

**Solusi**:

- Tkinter sudah built-in di Python Windows installer
- Reinstall Python dengan opsi "tcl/tk" dicentang

### Default Folder Tidak Ada

**Problem**: Default folder sudah dihapus atau dipindah

**Solusi**:

- Sistem auto-fallback ke current directory
- Update default folder di menu Pengaturan

### Config File Corrupt

**Problem**: Error saat load `app_config.json`

**Solusi**:

- Hapus file di `C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\app_config.json`
- Aplikasi akan auto-create config baru
- Set ulang default folder di menu Pengaturan

### Database File Hilang

**Problem**: File `database.xlsx` atau `universal_scan_database.xlsx` tidak ditemukan

**Solusi**:

- Cek folder AppData: `%LOCALAPPDATA%\ArsipDigitalOwnCloud`
- Aplikasi akan auto-create file baru jika tidak ada
- Restore dari backup jika tersedia

### Excel Export Gagal

**Problem**: Permission denied saat export

**Solusi**:

- Tutup file Excel jika sedang dibuka
- Pastikan folder tujuan writable
- Gunakan nama file yang berbeda

### File Tidak Terbuka (Double-click)

**Problem**: Error saat double-click file di treeview

**Solusi**:

- Pastikan file masih ada di lokasi tersebut
- Check permission read file
- Pastikan Excel terinstall (untuk .xlsm)

---

## ❓ FAQ

### General

**Q: Apakah aplikasi ini gratis?**
A: Ya, aplikasi ini gratis untuk penggunaan internal organisasi.

**Q: Apakah perlu koneksi internet?**
A: Tidak, aplikasi berjalan full offline.

**Q: Apakah support macOS/Linux?**
A: Saat ini hanya Windows. Fitur `os.startfile()` Windows-specific.

### Fitur Scan

**Q: Berapa lama waktu scan folder besar?**
A: Tergantung jumlah file. Rata-rata beberapa detik untuk ribuan file.

**Q: Apakah scan recursive?**
A: Ya, scan dilakukan hingga semua level subfolder.

**Q: Apakah bisa scan network drive?**
A: Ya, selama drive ter-mapping dan accessible.

### File Format

**Q: Format Excel apa yang didukung?**
A: `.xlsx` dan `.xls` (dengan openpyxl dan pandas).

**Q: Apakah case-sensitive untuk nama file?**
A: Tidak. Aplikasi menggunakan case-insensitive comparison.

**Q: File tanpa ekstensi akan terdeteksi?**
A: Ya, di Mode Format Non-Dokumen.

### Export & Storage

**Q: Di mana hasil export disimpan?**
A: Lokasi default adalah root folder aplikasi, tapi bisa pilih lokasi lain.

**Q: Apakah hasil export otomatis terhapus?**
A: Tidak, hasil export persisten sampai dihapus manual.

**Q: Apakah bisa export ke format selain Excel?**
A: Saat ini hanya Excel (.xlsx). Format lain bisa ditambahkan di masa depan.

### Pengaturan

**Q: Apakah default folder wajib diset?**
A: Tidak, opsional. Tanpa default folder, dialog browse dari current directory.

**Q: Apakah bisa set default folder berbeda per form?**
A: Saat ini tidak. Satu default folder untuk semua form.

**Q: Apakah config terbawa saat copy aplikasi?**
A: Tidak, `app_config.json` di-gitignore. Setiap instalasi punya config sendiri.

### Cek Pengajuan Dana

**Q: Apakah bisa scan tahun sebelum 2020?**
A: Default range 2020-(current+1). Untuk tahun sebelumnya, perlu modifikasi kode.

**Q: Bagaimana jika file sudah dihapus?**
A: Saat double-click, muncul error "File Tidak Ditemukan". Scan ulang untuk refresh.

**Q: Apakah bisa filter hasil scan?**
A: Saat ini belum ada fitur filter built-in. Bisa filter manual di Excel hasil export.

### Cek NO KK

**Q: Kenapa OCR tidak mendeteksi NO KK?**
A: Pastikan:

- Tesseract OCR sudah terinstall
- File PDF bukan hasil scan dengan resolusi rendah (minimal 300 DPI)
- NO KK ada di header (20% bagian atas)
- Gunakan fitur Deskew jika gambar miring

**Q: File saya ada tapi Status File menunjukkan "Tidak Ada"?**
A: Path di database.xlsx mungkin sudah berubah. Scan ulang folder arsip untuk update path.

**Q: OCR lambat, apakah normal?**
A: Ya, OCR + preprocessing memakan waktu. Gunakan Pause/Resume untuk istirahat.

**Q: Di mana file database.xlsx disimpan?**
A: `C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\database.xlsx`

### PDF Tool

**Q: PDF to Images error "Poppler not found"?**
A: Poppler sudah dibundle di aplikasi portable. Jika run dari source, download Poppler dan letakkan di folder project.

**Q: Compress PDF tidak mengecilkan ukuran?**
A: Compress hanya rewrite PDF structure. Untuk image-heavy PDF, coba reduce image quality dulu sebelum buat PDF.

**Q: Kenapa Merge PDF urutan salah?**
A: Urutan sesuai pilihan file dialog. Pastikan pilih file dalam urutan yang benar (gunakan Ctrl+Click untuk multiple select berurutan).

### AppData Storage

**Q: Kenapa database.xlsx tidak ada di folder aplikasi?**
A: Sejak v1.1.5, semua file data disimpan di AppData (`C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud`) sesuai Windows best practice.

**Q: Bagaimana cara backup data?**
A: Backup folder `C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud` secara manual atau gunakan Windows Backup.

**Q: Apakah setiap user Windows perlu scan ulang?**
A: Ya, karena setiap user punya folder AppData sendiri (user isolation).

---

## 📞 Support & Contact

Untuk pertanyaan, bug report, atau feature request:

- **Developer**: Riky Dwianto
- **Email**: (hubungi tim IT MIS)
- **Repository**: https://github.com/rikydwianto/cekarsipdigital

---

## 📝 Version History

| Version   | Date     | Changes                                                |
| --------- | -------- | ------------------------------------------------------ |
| **1.0.0** | Sep 2024 | Initial release dengan Scan Folder                     |
| **1.0.1** | Sep 2024 | Fitur Scan File Besar (fixed)                          |
| **1.0.2** | Sep 2024 | Parameter ukuran minimum, filter owncloud sync         |
| **1.0.3** | Sep 2024 | Dual-mode: File Besar + Format Non-Dokumen             |
| **1.0.4** | Okt 2024 | Fitur Pengaturan & Default Folder                      |
| **1.0.5** | Okt 2024 | Fitur Cek Pengajuan Dana + Quick Open                  |
| **1.1.0** | Okt 2025 | PDF Tool (8 fitur), Universal Scan Database            |
| **1.1.5** | Okt 2025 | Cek NO KK (OCR), AppData storage, Inno Setup installer |

### Changelog v1.1.5

**🆕 Fitur Baru:**

- ✅ **Cek NO KK** - Ekstrak dan validasi Nomor Kartu Keluarga dari PDF menggunakan OCR
  - Auto deskew untuk KK yang miring
  - Character correction (b→6, O→0, etc.)
  - Pause/Resume untuk proses yang lama
  - Track file missing dengan Status File column
- ✅ **PDF Tool Lengkap** - 8 fitur PDF tools
  - Images to PDF, PDF to Images
  - Merge, Split, Compress
  - Rotate, Protect, Unlock PDF
- ✅ **Universal Scan Database** - Scan komprehensif dengan auto-detect structure
- ✅ **AppData Storage** - File database dan export disimpan di AppData Local
- ✅ **Inno Setup Script** - Installer profesional untuk distribusi

**🔧 Perbaikan:**

- ✅ Poppler bundled dengan aplikasi (tidak perlu install terpisah untuk PDF→Images)
- ✅ Auto-detect Tesseract path (3 lokasi default)
- ✅ Responsive UI dengan layout 2 kolom
- ✅ Error messages lebih informatif dengan full path

**📦 Dependencies Update:**

- numpy 2.2.6 (untuk deskew algorithm)
- pytesseract 0.3.10 (OCR engine)
- pdf2image 1.17.0 (PDF conversion)
- Pillow 12.0.0 (image processing)

---

## 📄 License

Copyright © 2025 Riky Dwianto

Aplikasi ini untuk penggunaan internal organisasi. Tidak untuk distribusi komersial.

---

## 🙏 Credits

- **Framework**: Python 3.10 + Tkinter
- **Data Processing**: pandas 2.3.3 + openpyxl 3.1.5 + numpy 2.2.6
- **PDF Tools**: PyPDF2 3.0.1 + Pillow 12.0.0 + pdf2image 1.17.0
- **OCR Engine**: pytesseract 0.3.10 + Tesseract OCR
- **PDF Renderer**: Poppler 25.07.0 (bundled)
- **QR Code**: qrcode 8.0
- **Build Tool**: cx_Freeze 8.4.1
- **Installer**: Inno Setup 6.x
- **Developer**: Riky Dwianto

---

**Last Updated**: Oktober 30, 2025
