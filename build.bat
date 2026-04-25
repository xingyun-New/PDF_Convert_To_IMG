@echo off
chcp 65001 >nul
setlocal

echo ============================================
echo  PDFtoJPG  -  Build standalone GUI app
echo ============================================
echo.

echo [1/2] Installing dependencies...
py -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo [2/2] Building windowed GUI bundle with PyInstaller...
echo (Using --onedir + --windowed: stable Tk window, fast startup;
echo  PyInstaller 6.x + Python 3.13 + --onefile has a known Tk visibility bug.)
py -m PyInstaller --onedir --windowed --name PDFtoJPG --clean --noconfirm pdf_to_jpg.py
if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Build finished.
echo  Output folder: dist\PDFtoJPG\
echo  Run by double-clicking: dist\PDFtoJPG\PDFtoJPG.exe
echo  Distribute by zipping the entire dist\PDFtoJPG\ folder.
echo ============================================
echo.
pause
