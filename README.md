# 📁 ARSIP OWNCLOUD - Aplikasi Manajemen Arsip Digital

> Aplikasi desktop berbasis Python untuk manajemen, scanning, dan analisis arsip digital Owncloud

**Versi**: 1.0.5
**Tanggal**: Oktober 2025
**Developer**: Riky Dwianto

---

## 📋 Daftar Isi

1. [Deskripsi Umum](#-deskripsi-umum)
2. [Instalasi &amp; Setup](#-instalasi--setup)
3. [Menu Utama](#-menu-utama)
4. [Fitur Detail](#-fitur-detail)
   - [Cek Arsip Digital](#1-cek-arsip-digital)
   - [Scan Folder Arsip Digital](#2-scan-folder-arsip-digital)
   - [Scan File Besar &amp; Format Non-Dokumen](#3-scan-file-besar--format-non-dokumen)
   - [Cek Pengajuan Dana](#4-cek-pengajuan-dana)
   - [Pengaturan](#5-pengaturan)
5. [Struktur Folder](#-struktur-folder)
6. [Dependencies](#-dependencies)
7. [Build Executable](#-build-executable)
8. [Tips &amp; Best Practices](#-tips--best-practices)
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

| Fitur                        | Deskripsi                                            |
| ---------------------------- | ---------------------------------------------------- |
| **Cek Arsip Digital**  | Matching data folder dengan database Excel anggota   |
| **Scan Folder**        | Validasi 8 folder standar dengan export detail       |
| **Scan File Besar**    | Deteksi file berukuran besar dengan threshold kustom |
| **Format Non-Dokumen** | Identifikasi file dengan ekstensi tidak standar      |
| **Cek Pengajuan Dana** | Scan otomatis file PENGAJUAN_DANA.xlsm multi-tahun   |
| **Pengaturan**         | Simpan default folder untuk semua form               |

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

## 🏠 Menu Utama

Aplikasi memiliki menu utama dengan 5 tombol:

```
┌─────────────────────────────────────┐
│      📁 APLIKASI ARSIP DIGITAL      │
│   Sistem Manajemen Arsip Digital   │
├─────────────────────────────────────┤
│  📋 Cek Arsip Digital              │
│  📂 Scan Folder Arsip Digital      │
│  📊 Scan File Besar                │
│  💰 Cek Pengajuan Dana             │
│  ⚙️  Pengaturan                     │
├─────────────────────────────────────┤
│           [Keluar]                  │
│     v1.0.5 - Developed by RD        │
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
5. **05.BUKU_HARI_RAYA_ANGGOTA**

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

### 3. Scan File Besar & Format Non-Dokumen

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

### 4. Cek Pengajuan Dana

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

| No | Data Field                        | Lokasi         | Cell | Keterangan                       |
| -- | --------------------------------- | -------------- | ---- | -------------------------------- |
| 1  | **Nomor Surat**             | Sheet Surat    | F8   | Nomor surat dari dalam file      |
| 2  | **Nominal Input Kebutuhan** | Sheet Surat    | I8   | Nominal kebutuhan input          |
| 3  | **Nominal Kebutuhan**       | Sheet Laporan  | F68  | Total nominal kebutuhan          |
| 4  | **Status Balance**          | Sheet Laporan  | A4   | Status balance (BALANCE/SELISIH) |
| 5  | **Tanggal Disburse Awal**   | Sheet Lampiran | C3   | Tanggal mulai disburse           |
| 6  | **Tanggal Disburse Akhir**  | Sheet Lampiran | E3   | Tanggal akhir disburse           |
| 7  | **Nama BM**                 | Sheet Laporan  | A83  | Nama Branch Manager              |

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

### 5. Pengaturan

**Fungsi**: Konfigurasi default folder untuk semua form

#### Fitur

- **Set Default Folder** - Pilih folder yang sering digunakan
- **Auto-Load** - Semua browse dialog langsung ke folder ini
- **Hapus Default** - Reset ke current directory
- **Persistent Storage** - Disimpan di `app_config.json`

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

**app_config.json** (auto-generated):

```json
{
  "default_folder": "D:\\Data_Anggota_Owncloud"
}
```

**Lokasi**: Root folder aplikasi

**Security**:

- File ada di `.gitignore`
- Tidak terbawa ke Git repository
- Setiap komputer punya config sendiri

---

## 📂 Struktur Folder

### Workspace Structure

```
ARSIPOWNCLOUD/
├── main.py                    # Aplikasi utama
├── arsip_logic.py             # Business logic
├── requirements.txt           # Dependencies
├── README.md                  # Dokumentasi (file ini)
├── .gitignore                 # Git ignore rules
├── app_config.json           # Config (auto-generated, gitignored)
├── file_export.xlsx          # Export result (gitignored)
├── .venv/                     # Virtual environment
├── build/                     # Build artifacts
└── ArsipOwncloud_Portable/   # Portable executable
```

### Config Files

| File                 | Deskripsi                  | Git        |
| -------------------- | -------------------------- | ---------- |
| `app_config.json`  | Konfigurasi default folder | ❌ Ignored |
| `file_export.xlsx` | File export hasil scan     | ❌ Ignored |
| `requirements.txt` | Python dependencies        | ✅ Tracked |
| `.gitignore`       | Git ignore rules           | ✅ Tracked |

---

## 📦 Dependencies

### Runtime Dependencies

```txt
pandas>=2.0.0
openpyxl>=3.1.0
```

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

### Opsi 1: cx_Freeze (RECOMMENDED)

```bash
# 1. Install cx_Freeze
pip install cx-Freeze

# 2. Build portable version
.\build_portable.bat

# Output: ArsipOwncloud_Portable\ArsipOwncloud.exe
```

**Keuntungan cx_Freeze:**

- ✅ Compatible dengan pandas & numpy
- ✅ Portable folder (bisa di-copy)
- ✅ Include semua dependencies

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

### Build Scripts

| Script                 | Deskripsi                                |
| ---------------------- | ---------------------------------------- |
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

- Hapus file `app_config.json`
- Aplikasi akan auto-create config baru
- Set ulang default folder

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

---

## 📞 Support & Contact

Untuk pertanyaan, bug report, atau feature request:

- **Developer**: Riky Dwianto
- **Email**: (hubungi tim IT MIS)
- **Repository**: https://github.com/rikydwianto/cekarsipdigital

---

## 📝 Version History

| Version         | Date     | Changes                                        |
| --------------- | -------- | ---------------------------------------------- |
| **1.0.0** | Sep 2024 | Initial release dengan Scan Folder             |
| **1.0.1** | Sep 2024 | Fitur Scan File Besar (fixed)                  |
| **1.0.2** | Sep 2024 | Parameter ukuran minimum, filter owncloud sync |
| **1.0.3** | Sep 2024 | Dual-mode: File Besar + Format Non-Dokumen     |
| **1.0.4** | Okt 2024 | Fitur Pengaturan & Default Folder              |
| **1.0.5** | Okt 2024 | Fitur Cek Pengajuan Dana + Quick Open          |

---

## 📄 License

Copyright © 2024-2025 Riky Dwianto

Aplikasi ini untuk penggunaan internal organisasi. Tidak untuk distribusi komersial.

---

## 🙏 Credits

- **Framework**: Python + Tkinter
- **Data Processing**: pandas + openpyxl
- **Build Tool**: cx_Freeze
- **Developer**: Riky Dwianto

---

**Last Updated**: Oktober 17, 2025
