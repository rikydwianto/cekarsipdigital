import sys
import os
from cx_Freeze import setup, Executable

# --- Blok untuk menyertakan C++ Redistributable ---
# Menemukan path folder redistributable di komputer build Anda
PYTHON_DIR = os.path.dirname(sys.executable)
redist_path = os.path.join(PYTHON_DIR, "Library", "bin")

# Jika tidak ditemukan (mungkin bukan venv anaconda), cari di System32
if not os.path.exists(redist_path):
    redist_path = os.path.join(os.environ.get("SystemRoot", "C:/Windows"), "System32")

# Daftar DLL yang sering dibutuhkan numpy/pandas
# Jika Anda tahu versi persisnya, Anda bisa lebih spesifik
# (misal 'vcruntime140.dll', 'msvcp140.dll')
# Menggunakan glob pattern (*) lebih aman untuk mencakup semua
include_files = [
    (os.path.join(redist_path, "msvcp140.dll"), "msvcp140.dll"),
    (os.path.join(redist_path, "vcruntime140.dll"), "vcruntime140.dll"),
    (os.path.join(redist_path, "vcruntime140_1.dll"), "vcruntime140_1.dll"),
]

# Filter file yang benar-benar ada di sistem Anda
include_files = [f for f in include_files if os.path.exists(f[0])]

# Tambahkan folder-folder yang perlu di-bundle
additional_includes = [
    ("src_web", "src_web"),                      # Web server templates & static files
    ("app_config.json", "app_config.json"),      # Config file
    ("poppler-25.07.0", "poppler-25.07.0"),      # Poppler untuk PDF → Images
]

# Tambahkan ke include_files
include_files.extend(additional_includes)
# ----------------------------------------------------


build_exe_options = {
    "packages": [
        "numpy", 
        "pandas"
    ],
    "include_files": include_files # <-- Tambahkan DLL di sini
}

setup(
    name="CekArsip",
    version="1.0",
    description="Cek Arsip Digital",
    options={"build_exe": build_exe_options},  
    executables=[
        Executable(
            "main.py",
            base="Win32GUI",
            icon="icon.ico"  # Icon untuk EXE file
        )
    ]
)