# PS_Rechnungstool.spec
# =====================
# PyInstaller-Spec-Datei für den Windows-EXE-Build des PS-Rechnungstools.
#
# Build-Befehl (Windows-Kommandozeile oder Skript):
#   pyinstaller PS_Rechnungstool.spec
#
# Alternativ:  build_exe.bat ausführen (installiert Abhängigkeiten + baut EXE).
#
# Ergebnis: dist/PS_Rechnungstool/PS_Rechnungstool.exe  (onedir)
#       oder dist/PS_Rechnungstool.exe                  (onefile, siehe unten)

import sys
from pathlib import Path

block_cipher = None

# Alle Quell-Python-Dateien
_src = Path(SPECPATH)

a = Analysis(
    [str(_src / "gui.py")],
    pathex=[str(_src)],
    binaries=[],
    datas=[
        # Füge hier ggf. weitere Ressourcen hinzu, z. B. Icons:
        # (str(_src / "icon.ico"), "."),
    ],
    hiddenimports=[
        "pdfplumber",
        "pdfminer",
        "pdfminer.high_level",
        "pdfminer.layout",
        "openpyxl",
        "openpyxl.styles",
        "Rechnung_umbenennen_1",
        "Rechnungsdatenübernahme",
        "compare_excel_customer_and_partial_company_email_update",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PS_Rechnungstool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # Kein Konsolenfenster – GUI-Anwendung
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="icon.ico",      # Icon-Datei eintragen, falls vorhanden
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="PS_Rechnungstool",
)
