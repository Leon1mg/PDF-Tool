@echo off

:: Python-Version überprüfen
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python ist nicht installiert oder nicht im PATH. Bitte installiere Python und stelle sicher, dass es im PATH ist.
    pause
    exit /b
)

:: pip-Version überprüfen
pip --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo pip ist nicht installiert. Bitte installiere pip, bevor du fortfährst.
    pause
    exit /b
)

:: Erforderliche Pakete installieren
echo Installiere erforderliche Pakete...
pip install --upgrade pip
pip install tk pillow pymupdf pywin32 pypdf2

:: Nach der Installation pywin32 konfigurieren
echo Konfiguriere pywin32...
python -m pywin32_postinstall -install >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Fehler bei der Konfiguration von pywin32. Bitte manuell pruefen.
)

:: Überprüfen, ob Pakete korrekt installiert wurden
echo.
echo Überprüfung der Installation:
pip show tk >nul 2>&1 && echo - tk installiert || echo - tk NICHT installiert
pip show pillow >nul 2>&1 && echo - Pillow installiert || echo - Pillow NICHT installiert
pip show pymupdf >nul 2>&1 && echo - PyMuPDF installiert || echo - PyMuPDF NICHT installiert
pip show pywin32 >nul 2>&1 && echo - pywin32 installiert || echo - pywin32 NICHT installiert
pip show pypdf2 >nul 2>&1 && echo - PyPDF2 installiert || echo - PyPDF2 NICHT installiert

echo.
echo Installation abgeschlossen. Druecke eine beliebige Taste, um fortzufahren.
pause
