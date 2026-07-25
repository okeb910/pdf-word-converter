@echo off
chcp 65001 >nul
title PDF ^<-> Word 本地转换工具

cd /d "%~dp0"

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ============================================
echo   PDF ^<-> Word 本地转换工具
echo   优先本机 Word，可选图片模式 / LibreOffice
echo ============================================
echo.
echo [1/2] 检查并安装 Python 依赖...
pip install -r requirements.txt -q 2>nul
if %errorlevel% neq 0 (
    echo [警告] 依赖安装可能未完全成功，可手动执行: pip install -r requirements.txt
)
echo [2/2] 启动转换工具...
echo.
echo 提示: 推荐安装 Microsoft Word；图片模式需 PyMuPDF；也可使用 LibreOffice。
echo.
python pdf_word_converter.py
if %errorlevel% neq 0 (
    echo.
    echo [错误] 程序异常退出，错误码: %errorlevel%
    pause
    exit /b %errorlevel%
)

exit /b 0
