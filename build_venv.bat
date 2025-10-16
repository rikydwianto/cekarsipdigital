@echo off
echo ========================================
echo Building EXE with Virtual Environment
echo ========================================
echo.

REM Aktivasi virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo.
echo Installing PyInstaller...
pip install pyinstaller

echo.
echo Building EXE...
pyinstaller --onefile --windowed --name "ArsipOwncloud" ^
    --exclude-module torch ^
    --exclude-module tensorflow ^
    --exclude-module scipy ^
    --exclude-module matplotlib ^
    --exclude-module PIL ^
    --exclude-module cv2 ^
    main.py

if errorlevel 1 (
    echo.
    echo Build FAILED!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build SUCCESS!
echo ========================================
echo.
echo EXE Location: dist\ArsipOwncloud.exe
echo.
pause
