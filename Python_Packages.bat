@echo off
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python ist nicht installiert. Bitte installiere Python, bevor du fortfährst.
    pause
    exit /b
)

pip --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo pip ist nicht installiert. Bitte installiere pip, bevor du fortfährst.
    pause
    exit /b
)

echo Installiere erforderliche Pakete...
pip install tk
pip install pillow
pip install pymupdf

echo.
echo Überprüfung der Installation:
pip show tk >nul 2>&1 && echo - tk installiert || echo - tk NICHT installiert
pip show pillow >nul 2>&1 && echo - Pillow installiert || echo - Pillow NICHT installiert
pip show pymupdf >nul 2>&1 && echo - PyMuPDF installiert || echo - PyMuPDF NICHT installiert

echo.
echo Installation abgeschlossen.
pause