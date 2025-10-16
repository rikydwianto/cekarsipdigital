@echo off
echo ========================================
echo Build Portable ArsipOwncloud
echo ========================================

REM Aktifkan virtual environment
call venv_exe\Scripts\activate

REM Build dengan cx_Freeze
echo Building with cx_Freeze...
python setup.py build

REM Buat folder distribusi
echo Creating portable package...
if exist "ArsipOwncloud_Portable" rmdir /s /q "ArsipOwncloud_Portable"
mkdir "ArsipOwncloud_Portable"

REM Copy semua file yang diperlukan
xcopy "build\exe.win-amd64-3.10\*.*" "ArsipOwncloud_Portable\" /E /I /Y

echo.
echo ========================================
echo Build Selesai!
echo ========================================
echo.
echo File EXE tersedia di:
echo %CD%\ArsipOwncloud_Portable\ArsipOwncloud.exe
echo.
echo Anda bisa copy folder 'ArsipOwncloud_Portable' 
echo ke komputer lain tanpa perlu install Python
echo.
pause
