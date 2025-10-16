from cx_Freeze import setup, Executable

setup(
    name="CekArsip",
    version="1.0",
    description="Cek Arsip Digital",
    executables=[Executable("main.py", base="Win32GUI")]
)
