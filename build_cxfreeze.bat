@echo off
echo ========================================
echo Building EXE with cx_Freeze...
echo ========================================

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

python setup.py build_exe

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo Build SUCCESS!
    echo ========================================
    echo EXE file is in: build\exe.win-amd64-3.10\ArsipOwncloud.exe
) else (
    echo.
    echo ========================================
    echo ERROR: Build failed!
    echo ========================================
)

pause
