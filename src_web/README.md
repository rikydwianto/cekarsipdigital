# Arsip Owncloud Web Server

Sistem web server modular dengan template system untuk aplikasi Arsip Owncloud.

## 📁 Struktur Folder

```
src_web/
├── __init__.py                  # Package initializer
├── template_loader.py           # Template loader & renderer
├── templates/                   # HTML templates
│   ├── index.html              # Halaman utama
│   └── error.html              # Halaman error
└── static/                      # Static files
    ├── css/
    │   └── style.css           # Main stylesheet
    └── js/
        └── main.js             # Main JavaScript
```

## 🎨 Fitur

### Template System

- **Simple Variable Replacement**: Menggunakan `{{variable}}` syntax
- **Automatic Loading**: Load template dari folder `templates/`
- **Context Support**: Pass data ke template melalui context dictionary

### Static Files

- **CSS**: Styling terpisah di `static/css/`
- **JavaScript**: Script terpisah di `static/js/`
- **Auto Content-Type**: Deteksi otomatis berdasarkan ekstensi file
- **Security**: Prevention untuk directory traversal

### Responsive Design

- Mobile-friendly
- Gradient background
- Smooth animations
- Hover effects pada info boxes

## 🚀 Cara Menggunakan

### 1. Load Template

```python
from src_web.template_loader import TemplateLoader

loader = TemplateLoader()
html = loader.render('index.html', context={'name': 'World'})
```

### 2. Serve Static Files

```python
content, content_type = loader.get_static_file('css/style.css')
```

### 3. Get Default Context

```python
from src_web.template_loader import get_default_context

context = get_default_context(server_info={
    'local_ip': '192.168.1.100',
    'port': 8080,
    'default_folder': 'C:/Data'
})
```

## 📝 Template Variables

### index.html

- `{{server_time}}` - Waktu server
- `{{os_info}}` - Informasi OS
- `{{python_version}}` - Versi Python
- `{{hostname}}` - Nama host
- `{{local_ip}}` - IP Address lokal
- `{{port}}` - Port server
- `{{default_folder}}` - Document root folder

### error.html

- `{{error_message}}` - Pesan error

## 🎯 URL Routes

- `/` atau `/index.html` - Halaman utama
- `/static/css/style.css` - Main stylesheet
- `/static/js/main.js` - Main JavaScript

## 🔧 Menambah Template Baru

### 1. Buat file template

Buat file HTML di `src_web/templates/new_page.html`:

```html
<!DOCTYPE html>
<html>
  <head>
    <link rel="stylesheet" href="/static/css/style.css" />
  </head>
  <body>
    <h1>{{title}}</h1>
    <p>{{content}}</p>
  </body>
</html>
```

### 2. Render template

```python
context = {'title': 'Halaman Baru', 'content': 'Ini konten'}
html = loader.render('new_page.html', context)
```

## 💡 Tips

1. **Pisahkan CSS**: Semua styling di `style.css` agar mudah dimaintain
2. **Gunakan JavaScript**: Untuk fitur interaktif seperti real-time clock
3. **Context Reusable**: Gunakan `get_default_context()` untuk data umum
4. **Security**: Static file loader sudah prevent directory traversal
5. **Error Handling**: Selalu ada fallback jika template gagal load

## 🔒 Security Features

- Directory traversal prevention
- Safe path normalization
- Content-type validation
- UTF-8 encoding enforcement

## 📊 Content Types Support

- `.css` - text/css
- `.js` - application/javascript
- `.html` - text/html
- `.json` - application/json
- `.png`, `.jpg`, `.gif` - image/\*
- `.svg` - image/svg+xml
- `.ico` - image/x-icon

## 🎨 Customization

### Ubah Warna Theme

Edit `static/css/style.css`:

```css
body {
  background: linear-gradient(135deg, #your-color1 0%, #your-color2 100%);
}
```

### Tambah Animation

Edit `static/js/main.js`:

```javascript
// Custom animation code
```

## 📦 Dependencies

- Python 3.10+
- Standard library only (os, datetime, platform)

## 🌟 Keuntungan Sistem Template

✅ **Clean Code**: HTML, CSS, JS terpisah
✅ **Easy Maintenance**: Edit tanpa touch Python code
✅ **Reusable**: Template bisa dipakai ulang
✅ **Scalable**: Mudah tambah halaman baru
✅ **Professional**: Struktur seperti framework modern

---

**Made with ❤️ for Arsip Owncloud**
