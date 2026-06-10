@echo off
echo ========================================
echo  PackingListGenerator - Windows Build
echo ========================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Install Python 3.9+
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Install dependencies
echo [1/3] Installing dependencies...
pip install pyinstaller openpyxl xlrd -q
if %errorlevel% neq 0 (
    echo [ERROR] Dependency install failed
    pause
    exit /b 1
)

:: Build
echo [2/3] Building (may take 3-5 minutes)...
pyinstaller ^
  --onedir ^
  --name "PackingListGenerator" ^
  --add-data "template.xlsx;." ^
  --hidden-import email.policy ^
  --hidden-import xlrd ^
  --hidden-import openpyxl ^
  web_gui.py

if %errorlevel% neq 0 (
    echo [ERROR] Build failed
    pause
    exit /b 1
)

:: Clean up
echo [3/3] Cleaning up build cache...
rmdir /s /q build 2>nul
del /q PackingListGenerator.spec 2>nul

echo.
echo ========================================
echo  Build complete!
echo  Output: dist\PackingListGenerator\
echo  EXE:    dist\PackingListGenerator\PackingListGenerator.exe
echo ========================================
pause
