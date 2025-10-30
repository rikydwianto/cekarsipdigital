# Changelog - Penyimpanan File ke AppData

## 📌 Perubahan Penting (Breaking Changes)

### Lokasi Penyimpanan File

**Sebelum:** File `database.xlsx` dan `file_export.xlsx` disimpan di folder aplikasi yang sama dengan `main.exe`.

**Sekarang:** File-file tersebut disimpan di **AppData Local** untuk memenuhi best practice Windows:

```
C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\
├── database.xlsx
└── file_export.xlsx
```

### Alasan Perubahan

1. ✅ **Best Practice Windows**: Aplikasi tidak boleh menulis file ke `Program Files`
2. ✅ **Keamanan**: Tidak perlu admin privileges untuk write access
3. ✅ **User Data Isolation**: Setiap user Windows punya data sendiri
4. ✅ **Mudah Backup**: User tinggal backup folder AppData
5. ✅ **Kompatibilitas**: Sesuai dengan standar Windows modern apps

### Yang Berubah di Aplikasi

#### 1. **Scan Folder Arsip Digital**

- File `database.xlsx` sekarang dibuat di: `C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\database.xlsx`
- Tidak ada perubahan di UI atau workflow

#### 2. **Cek NO KK**

- Membaca `database.xlsx` dari AppData
- Info dialog menampilkan lokasi lengkap file
- Error message menampilkan path lengkap jika file tidak ditemukan

#### 3. **Export Data**

- File `file_export.xlsx` disimpan ke AppData
- Message box menampilkan lokasi lengkap file hasil export
- Opsi "Save As" masih tetap bisa memilih lokasi custom

#### 4. **Web Server**

- Web server membaca `database.xlsx` dari AppData
- Tidak ada perubahan di API atau endpoint
- Cache masih bekerja dengan baik

### Cara Akses File

#### Via File Explorer:

1. Tekan `Win + R`
2. Ketik: `%LOCALAPPDATA%\ArsipDigitalOwnCloud`
3. Tekan Enter

#### Via Aplikasi:

- Aplikasi otomatis membuat folder jika belum ada
- Semua operasi transparant untuk user
- File path ditampilkan di message box

### Migration dari Versi Lama

Jika Anda upgrade dari versi lama yang masih menyimpan file di folder aplikasi:

1. **Copy file lama ke lokasi baru:**

   ```powershell
   # Buka PowerShell
   $oldPath = "D:\path\to\old\database.xlsx"
   $newPath = "$env:LOCALAPPDATA\ArsipDigitalOwnCloud\database.xlsx"
   Copy-Item $oldPath $newPath
   ```

2. **Atau scan ulang:**
   - Buka aplikasi → Scan Folder Arsip Digital
   - Pilih folder → Simpan dan Singkron
   - File baru akan dibuat di lokasi AppData

### Backup Data

**Backup Otomatis Windows:**

- File di AppData terbackup otomatis jika Anda enable Windows Backup
- OneDrive dapat sync folder AppData (jika dikonfigurasi)

**Backup Manual:**

```powershell
# Backup ke folder lain
$source = "$env:LOCALAPPDATA\ArsipDigitalOwnCloud"
$backup = "D:\Backup\ArsipDigital_$(Get-Date -Format 'yyyyMMdd')"
Copy-Item $source $backup -Recurse
```

### Uninstall Aplikasi

File data TIDAK otomatis terhapus saat uninstall aplikasi.

**Hapus data manual:**

1. Tekan `Win + R`
2. Ketik: `%LOCALAPPDATA%`
3. Hapus folder `ArsipDigitalOwnCloud`

### Troubleshooting

**Q: File database.xlsx tidak ditemukan?**

- Pastikan sudah scan folder arsip terlebih dahulu
- Check lokasi: `C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\`
- Error message akan menampilkan path lengkap

**Q: Tidak bisa akses folder AppData?**

- Folder hidden by default di Windows
- Enable "Show hidden files" di File Explorer options
- Atau gunakan shortcut `Win + R` → `%LOCALAPPDATA%`

**Q: Setiap user Windows perlu scan ulang?**

- Ya, karena setiap user punya folder AppData sendiri
- Ini by design untuk user data isolation

**Q: Bisa ganti lokasi penyimpanan?**

- Tidak recommended, karena melanggar Windows best practice
- Jika perlu, bisa export manual ke lokasi lain

### Technical Details

**Implementasi:**

```python
# Helper functions di main.py dan web_server.py
def get_appdata_path():
    """Get AppData Local path untuk aplikasi"""
    appdata = os.getenv('LOCALAPPDATA')
    app_folder = os.path.join(appdata, 'ArsipDigitalOwnCloud')
    os.makedirs(app_folder, exist_ok=True)
    return app_folder

def get_database_path():
    return os.path.join(get_appdata_path(), 'database.xlsx')

def get_export_path():
    return os.path.join(get_appdata_path(), 'file_export.xlsx')
```

**Folder Creation:**

- Folder AppData dibuat otomatis saat import module
- Tidak perlu admin privileges
- Error handling jika gagal create folder

---

## 📝 Version History

**v1.1.0** (Current)

- ✅ Pindah penyimpanan ke AppData Local
- ✅ Update semua referensi file path
- ✅ Improved error messages dengan full path
- ✅ Auto-create folder AppData

**v1.0.0** (Old)

- ❌ Simpan file di folder aplikasi (tidak recommended)

---

**Dibuat:** 29 Oktober 2025
**Author:** Arsip Digital Development Team
