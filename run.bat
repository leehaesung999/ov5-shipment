@echo off
chcp 65001 > nul
cd /d "%~dp0"
python app.py
if errorlevel 1 (
    echo.
    echo [ERROR] 실행 중 오류가 발생했습니다.
    pause
)
