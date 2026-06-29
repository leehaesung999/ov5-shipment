@echo off
title Build - OV5 지정출고
chcp 65001 > nul
cd /d "%~dp0"

if not exist "%~dp0python\python.exe" (
    echo.
    echo [ERROR] python 폴더가 없습니다.
    echo 입고_파레트구분기.zip 의 python 폴더를 이 위치에 복사하세요.
    echo.
    pause
    exit /b 1
)

echo === 패키지 설치 (동봉 Python) ===
"%~dp0python\python.exe" -m pip install --upgrade pip
"%~dp0python\python.exe" -m pip install streamlit pandas openpyxl
echo.

echo === 배포판 zip 생성 ===
"%~dp0python\python.exe" build.py
echo.
pause
