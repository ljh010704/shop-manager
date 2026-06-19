@echo off
chcp 65001 >nul
echo ========================================
echo   Douyin Shop Profit Tracker
echo ========================================
echo.

cd /d "%~dp0"

echo [1/2] Installing dependencies...
pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    pip3 install -r requirements.txt -q
)

echo [2/2] Starting server...
echo.
echo Open browser: http://localhost:8000
echo Press Ctrl+C to stop
echo.
python app.py
pause
