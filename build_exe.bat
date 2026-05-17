@echo off
:: build_exe.bat
:: =============
:: Installiert Abhängigkeiten und baut eine Windows-EXE mit PyInstaller.
::
:: Voraussetzungen:
::   - Python 3.10+ muss installiert und im PATH sein
::   - Internetverbindung für pip install (nur beim ersten Mal nötig)
::
:: Ergebnis:
::   dist\PS_Rechnungstool\PS_Rechnungstool.exe
::
:: Aufruf:
::   build_exe.bat
:: oder als Administrator für systemweite Installation.

setlocal

echo ============================================
echo  PS Rechnungstool – Windows-EXE-Build
echo ============================================
echo.

:: Abhängigkeiten installieren
echo [1/3] Installiere Python-Pakete ...
pip install --upgrade pyinstaller pdfplumber openpyxl
if errorlevel 1 (
    echo.
    echo [FEHLER] pip install fehlgeschlagen. Bitte Python und pip prüfen.
    pause
    exit /b 1
)

echo.
echo [2/3] Bereinige alte Build-Artefakte ...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

echo.
echo [3/3] Baue EXE mit PyInstaller ...
pyinstaller PS_Rechnungstool.spec
if errorlevel 1 (
    echo.
    echo [FEHLER] PyInstaller-Build fehlgeschlagen.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Build erfolgreich!
echo  EXE: dist\PS_Rechnungstool\PS_Rechnungstool.exe
echo ============================================
pause
endlocal
