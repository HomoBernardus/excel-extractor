@echo off
chcp 65001 >nul
echo ========================================
echo  分箱单生成工具 — Windows 构建脚本
echo ========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.9+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 安装依赖
echo [1/3] 安装依赖...
pip install pyinstaller openpyxl xlrd -q
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

:: 构建
echo [2/3] 开始构建 (约需 3-5 分钟)...
pyinstaller ^
  --onedir ^
  --name "分箱单生成工具" ^
  --add-data "分箱单.xlsx;." ^
  --hidden-import email.policy ^
  --hidden-import xlrd ^
  --hidden-import openpyxl ^
  web_gui.py

if %errorlevel% neq 0 (
    echo [错误] 构建失败
    pause
    exit /b 1
)

:: 清理
echo [3/3] 清理构建缓存...
rmdir /s /q build 2>nul
del /q 分箱单生成工具.spec 2>nul

echo.
echo ========================================
echo  构建完成!
echo  输出目录: dist\分箱单生成工具\
echo  可执行文件: dist\分箱单生成工具\分箱单生成工具.exe
echo ========================================
pause
