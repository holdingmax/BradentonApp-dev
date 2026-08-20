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

if not exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo.
    echo NOTE: Tesseract OCR is not installed. It is required to read
    echo photographed/scanned Gettel-Toyota reports and is a separate
    echo program, not a Python package. Install it with:
    echo   winget install UB-Mannheim.TesseractOCR
)

pause
