@echo off
echo ========================================
echo Building EXE with cx_Freeze...
echo ========================================

REM Cek apakah folder poppler-25.07.0 ada
if exist "poppler-25.07.0" (
    echo [OK] Poppler folder found: poppler-25.07.0
    echo      PDF to Images feature will work in built exe!
) else (
    echo [WARNING] Poppler folder NOT found!
    echo           PDF to Images feature will require manual setup.
    echo           Download from: https://github.com/oschwartz10612/poppler-windows/releases/
    echo.
)

REM Aktifkan virtual environment
call .venv\Scripts\activate

REM Hapus build lama
echo Cleaning old build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ArsipOwncloud.spec del /q ArsipOwncloud.spec

echo.
echo Building EXE file...
echo This may take a few minutes...
echo Bundling: src_web, app_config.json, poppler-25.07.0
echo.

python setup.py build_exe

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo Build SUCCESS!
    echo ========================================
    echo EXE file: build\exe.win-amd64-3.10\ArsipOwncloud.exe
    echo.
    if exist "poppler-25.07.0" (
        echo [OK] Poppler included in build!
        echo      Users can use PDF to Images without setup.
    )
    echo.
) else (
    echo.
    echo ========================================
    echo ERROR: Build failed!
    echo ========================================
)

pause
