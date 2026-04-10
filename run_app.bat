@echo off
echo ========================================
echo   p-net OrderReply Tool
echo ========================================
echo.

cd /d "%~dp0"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python이 설치되어 있지 않습니다.
    echo Python을 설치한 후 다시 실행해주세요.
    pause
    exit /b 1
)

REM Check if required packages are installed
python -c "import pandas, openpyxl, tkinter" >nul 2>&1
if errorlevel 1 (
    echo 필요한 패키지를 설치하는 중...
    python -m pip install pandas openpyxl
    if errorlevel 1 (
        echo ERROR: 패키지 설치에 실패했습니다.
        pause
        exit /b 1
    )
)

echo GUI 애플리케이션을 시작합니다...
python main.py

pause