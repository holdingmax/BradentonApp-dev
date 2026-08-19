@echo off
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Install failed. Ensure Python is on PATH, then retry.
    pause
    exit /b 1
)
echo.
echo Dependencies installed successfully.
pause
