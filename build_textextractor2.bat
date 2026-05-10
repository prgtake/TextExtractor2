@echo off
echo ======================================================
echo  TextExtractor2 Pro - Build Script
echo ======================================================

echo.
echo [1/3] Installing dependencies...
pip install -r requirements.txt --quiet

echo.
echo [2/3] Creating executable with PyInstaller...
echo (Excluding pytest/py to avoid Python 3.13 compatibility issues)
pyinstaller --onefile --noconsole --name TextExtractor2 --exclude-module pytest --exclude-module py textextractor2.py

echo.
echo [3/3] Cleanup temporary files...
# No need for complex cleanup, dist/ will have the exe.

echo.
echo Build complete! Check the 'dist' folder for TextExtractor2.exe.
pause
