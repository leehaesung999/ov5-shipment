@echo off
chcp 65001 > nul
cd /d "%~dp0"
title OV5 Designated Shipment

rem --- prevent Streamlit first-run email prompt ---
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    > "%USERPROFILE%\.streamlit\credentials.toml" echo [general]
    >> "%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)

rem --- pick Python (bundled if present) ---
if exist "%~dp0python\python.exe" (
    set "PY=%~dp0python\python.exe"
) else (
    set "PY=python"
)

echo.
echo =============================================
echo   OV5 Designated Shipment Matcher
echo   Starting... browser opens automatically
echo   ( http://localhost:8501 )
echo   Do NOT close this window while using the app.
echo =============================================
echo.

"%PY%" -m streamlit run app_streamlit.py
pause
