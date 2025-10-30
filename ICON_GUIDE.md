# 🎨 Panduan Mengganti Icon Aplikasi

## 📋 Daftar Isi

1. [Icon Window Aplikasi](#1-icon-window-aplikasi)
2. [Icon File EXE](#2-icon-file-exe)
3. [Cara Membuat/Convert Icon](#3-cara-membuatconvert-icon)
4. [Tips & Best Practices](#4-tips--best-practices)

---

## 1. Icon Window Aplikasi

Icon yang muncul di **taskbar** dan **title bar** saat aplikasi berjalan.

### Lokasi Code:

File: `main.py` baris 58

```python
try:
    self.root.iconbitmap("icon.ico")  # Icon window
except:
    pass
```

### Cara Pakai:

1. **Siapkan file icon.ico** di folder project
2. Code sudah ada, tinggal pastikan file `icon.ico` ada
3. Run aplikasi: `python main.py`

---

## 2. Icon File EXE

Icon yang muncul di **file explorer** untuk file `.exe` setelah di-build.

### Lokasi Code:

File: `setup.py` baris 52-58

```python
executables=[
    Executable(
        "main.py",
        base="Win32GUI",
        icon="icon.ico"  # Icon untuk EXE
    )
]
```

### Cara Build:

```bash
# Build dengan cx_Freeze
python setup.py build

# File exe akan ada di: build/exe.win-amd64-3.10/
# Icon akan muncul di file exe
```

---

## 3. Cara Membuat/Convert Icon

### Opsi A: Download Icon Siap Pakai

Website download icon gratis:

- 🔗 https://www.flaticon.com/
- 🔗 https://icons8.com/
- 🔗 https://www.iconfinder.com/

**Keyword search**: "folder", "archive", "document", "file manager"

### Opsi B: Convert dari PNG/JPG

**Gunakan script yang sudah dibuat:**

```bash
python convert_to_ico.py
```

**Atau manual:**

1. **Siapkan gambar PNG** (recommended: 256x256 atau 512x512)
2. **Jalankan script:**

   ```python
   from PIL import Image

   img = Image.open("logo.png")
   img.save("icon.ico", format='ICO',
            sizes=[(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)])
   ```

3. **Copy ke folder project**
   ```bash
   # icon.ico harus ada di folder yang sama dengan main.py
   ARSIPOWNCLOUD/
   ├── icon.ico          # <-- File icon di sini
   ├── main.py
   └── ...
   ```

### Opsi C: Online Converter

- 🔗 https://convertio.co/png-ico/
- 🔗 https://image.online-convert.com/convert-to-ico
- 🔗 https://www.icoconverter.com/

**Upload PNG → Download ICO → Copy ke project**

---

## 4. Tips & Best Practices

### ✅ Ukuran Icon Recommended:

| Size    | Purpose                         |
| ------- | ------------------------------- |
| 16x16   | Small icon (taskbar, title bar) |
| 32x32   | Standard icon                   |
| 48x48   | Desktop icon                    |
| 64x64   | Large icon                      |
| 128x128 | Extra large                     |
| 256x256 | Ultra large (Windows 7+)        |

### ✅ Format:

- **File format**: `.ico` (ICO)
- **Color depth**: 32-bit with alpha channel (RGBA)
- **Background**: Transparent atau solid color

### ✅ Design Tips:

1. **Simple & Clear** - Icon harus jelas di ukuran kecil (16x16)
2. **Consistent Style** - Sesuaikan dengan tema aplikasi
3. **Good Contrast** - Pastikan terlihat di background terang & gelap
4. **Professional** - Hindari terlalu banyak detail

### ✅ Testing:

```bash
# Test di aplikasi
python main.py

# Test di EXE
python setup.py build
cd build/exe.win-amd64-3.10
./main.exe
```

---

## 📝 Quick Start Guide

### Cara Tercepat:

1. **Download icon siap pakai** (format .ico)

   - Kunjungi: https://www.flaticon.com/
   - Search: "folder" atau "archive"
   - Download format ICO

2. **Rename menjadi `icon.ico`**

3. **Copy ke folder project**

   ```
   D:\PROJECT\PYTHON\ARSIPOWNCLOUD\icon.ico
   ```

4. **Test aplikasi**

   ```bash
   python main.py
   ```

5. **Build EXE** (optional)
   ```bash
   python setup.py build
   ```

**Done!** ✅ Icon akan muncul di window & file exe

---

## ⚠️ Troubleshooting

### Problem: Icon tidak muncul di window

**Solusi:**

```python
# Pastikan path benar
self.root.iconbitmap("icon.ico")  # Relative path
# atau
self.root.iconbitmap("D:/full/path/to/icon.ico")  # Absolute path
```

### Problem: Icon tidak muncul di EXE

**Solusi:**

1. Pastikan `icon.ico` ada saat build
2. Check `setup.py` sudah ada parameter `icon="icon.ico"`
3. Rebuild: `python setup.py build`

### Problem: Error "couldn't recognize data in image file"

**Solusi:**

- File bukan format ICO asli
- Convert ulang dengan PIL/Pillow:
  ```bash
  python convert_to_ico.py
  ```

### Problem: Icon blur/pecah

**Solusi:**

- Gunakan resolusi lebih tinggi (256x256)
- Pastikan multiple sizes di-include dalam ICO

---

## 🎨 Contoh Icon yang Cocok

Untuk aplikasi **Arsip Digital**, icon yang cocok:

### Style 1: Folder with Documents

```
📁 + 📄 = Folder dengan dokumen di dalamnya
```

### Style 2: Filing Cabinet

```
🗄️ = Cabinet arsip klasik
```

### Style 3: Cloud Archive

```
☁️ + 📦 = Cloud storage dengan arsip
```

### Style 4: Database/Storage

```
💾 = Database/storage icon
```

**Keyword search di icon website:**

- "archive folder"
- "document management"
- "file cabinet"
- "cloud storage"
- "database folder"

---

## 📚 Resources

### Tools:

- **PIL/Pillow**: Python library untuk image processing
- **cx_Freeze**: Build executable dengan icon

### Websites:

- **Flaticon**: Free icons (attribution required)
- **Icons8**: Free & premium icons
- **Iconfinder**: Large icon collection

### Documentation:

- Tkinter iconbitmap: https://docs.python.org/3/library/tkinter.html
- cx_Freeze options: https://cx-freeze.readthedocs.io/

---

**Updated**: October 30, 2025  
**Version**: 1.1.6  
**Author**: Riky Dwianto
