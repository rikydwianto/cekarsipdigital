@echo off
echo Installing PyInstaller...
pip install pyinstaller

echo.
echo Building EXE...
pyinstaller --onefile --windowed --name "ArsipOwncloud" main.py

echo.
echo Done! EXE file is in: dist\ArsipOwncloud.exe
pause
