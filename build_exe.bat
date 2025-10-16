@echo off
echo ========================================
echo Building Arsip Owncloud EXE
echo ========================================
echo.

REM Aktivasi virtual environment jika ada
if exist .venv\Scripts\activate.bat (
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
)

REM Install PyInstaller jika belum ada
echo.
echo Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
) else (
    echo PyInstaller already installed.
)

REM Hapus folder build dan dist lama
echo.
echo Cleaning old build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

REM Build EXE
echo.
echo Building EXE file...
echo This may take a few minutes...
echo.

pyinstaller --noconfirm --onefile --windowed ^
    --name "ArsipOwncloud" ^
    --icon=NONE ^
    --add-data ".venv/Lib/site-packages/pandas;pandas" ^
    --add-data ".venv/Lib/site-packages/openpyxl;openpyxl" ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.ttk" ^
    --hidden-import "tkinter.messagebox" ^
    --hidden-import "tkinter.filedialog" ^
    --hidden-import "csv" ^
    --hidden-import "datetime" ^
    --hidden-import "os" ^
    main.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR: Build failed!
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo.
echo EXE file location: dist\ArsipOwncloud.exe
echo.
echo You can now run the application by double-clicking:
echo %CD%\dist\ArsipOwncloud.exe
echo.
pause
