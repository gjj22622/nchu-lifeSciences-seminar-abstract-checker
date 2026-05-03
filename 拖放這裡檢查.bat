@echo off
chcp 65001 >nul
REM 中興生科所專討摘要檢查器 (Windows 拖放版)
REM 將 .docx 拖到本檔上即可自動檢查

if "%~1"=="" (
    echo.
    echo  ============================================================
    echo   中興生科所博士班專題討論 ─ 摘要檢查器
    echo  ============================================================
    echo.
    echo   用法：把你的摘要 .docx 檔拖到這個 .bat 檔上即可
    echo.
    echo   首次使用前需安裝 python-docx：
    echo     pip install python-docx
    echo.
    pause
    exit /b 1
)

cd /d "%~dp0"
python check_abstract.py "%~1"
echo.
pause
