"""
Rechnung_umbenennen_1.py
========================
Benennt Rechnungs-PDFs nach dem Schema

    Rechnung_<Rechnungsnummer>_<Projektname>.pdf

um, basierend auf dem im PDF enthaltenen Text.

Unterstützte Formate in der PDF-Zeile:
  "Rechnung <Projektname> REPS1234"   ← neu (bevorzugt)
  "Rechnung <Projektname> RE12345678" ← alt (abwärtskompatibel)

Dateien, die bereits dem Muster entsprechen, werden übersprungen.
"""

import os
import re
import sys

import pdfplumber

# ---------------------------------------------------------------------------
# Muster
# ---------------------------------------------------------------------------

# Gültiger Zieldateiname: Rechnung_REPS1234_Projektname.pdf
#                     oder Rechnung_RE12345678_Projektname.pdf
_valid_name_pattern = re.compile(
    r"^Rechnung_(?:REPS\d{4}|RE\d{8})_.+\.pdf$",
    re.IGNORECASE,
)

# Rechnungsblock in einer Zeile: "Rechnung <Projektname> REPS1234"
_invoice_line_pattern = re.compile(
    r"Rechnung\s+(.+?)\s+((?:REPS\d{4}|RE\d{8}))\b",
    re.IGNORECASE,
)

# Bereinigung unerwünschter Zeichen im Dateinamen
_unsafe_chars = re.compile(r'[\\/*?:"<>|]')


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def _sanitize_name(name: str) -> str:
    """Entfernt dateisystemfremde Zeichen und normalisiert Leerzeichen."""
    name = _unsafe_chars.sub("", name)
    # Mehrfache Leerzeichen/Tabs → Unterstrich
    name = re.sub(r"\s+", "_", name.strip())
    return name


def _extract_text_from_pdf(pdf_path: str) -> str:
    """Gibt den vollständigen Textinhalt der PDF zurück."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text


def _parse_invoice_info(text: str) -> tuple[str, str]:
    """Parst Projektname und Rechnungsnummer aus dem PDF-Text.

    Sucht nach der Zeile  'Rechnung <Projektname> REPS1234'.
    Gibt (rechnungsnummer, projektname) oder ("", "") zurück.
    """
    m = _invoice_line_pattern.search(text)
    if m:
        projekt = _sanitize_name(m.group(1))
        rechnung_nr = m.group(2).upper()
        return rechnung_nr, projekt
    return "", ""


# ---------------------------------------------------------------------------
# Umbenennungslogik
# ---------------------------------------------------------------------------

def rename_invoice_pdf(pdf_path: str, dry_run: bool = False) -> str | None:
    """Benennt eine einzelne PDF-Datei um.

    Returns:
        Neuer Pfad, wenn die Datei umbenannt wurde.
        None, wenn die Datei übersprungen wurde (bereits korrekt benannt
        oder kein Rechnungsmuster erkannt).
    """
    directory = os.path.dirname(os.path.abspath(pdf_path))
    filename = os.path.basename(pdf_path)

    # Bereits korrekt benannt?
    if _valid_name_pattern.match(filename):
        print(f"[SKIP] Bereits korrekt benannt: {filename}")
        return None

    # PDF-Text lesen
    try:
        text = _extract_text_from_pdf(pdf_path)
    except OSError as exc:
        print(f"[FEHLER] Dateizugriff: {filename} – {exc}")
        return None
    except Exception as exc:  # noqa: BLE001 – pdfplumber-Fehler (korrupte PDF etc.)
        print(f"[FEHLER] Kann PDF nicht lesen: {filename} – {exc}")
        return None

    rechnung_nr, projekt = _parse_invoice_info(text)

    if not rechnung_nr:
        print(f"[SKIP] Kein Rechnungsmuster gefunden: {filename}")
        return None

    if not projekt:
        projekt = "Unbekannt"

    new_filename = f"Rechnung_{rechnung_nr}_{projekt}.pdf"
    new_path = os.path.join(directory, new_filename)

    if os.path.abspath(pdf_path) == os.path.abspath(new_path):
        print(f"[SKIP] Name ist bereits korrekt: {filename}")
        return None

    if dry_run:
        print(f"[DRY-RUN] {filename}  →  {new_filename}")
        return new_path

    try:
        os.rename(pdf_path, new_path)
        print(f"[OK] {filename}  →  {new_filename}")
        return new_path
    except OSError as exc:
        print(f"[FEHLER] Umbenennen fehlgeschlagen: {filename} – {exc}")
        return None


def process_folder(folder: str, dry_run: bool = False) -> list[str]:
    """Benennt alle PDFs im angegebenen Ordner um.

    Returns:
        Liste der neuen Pfade aller umbenannten Dateien.
    """
    if not os.path.isdir(folder):
        print(f"[FEHLER] Ordner nicht gefunden: {folder}")
        return []

    pdf_files = [
        os.path.join(folder, f)
        for f in sorted(os.listdir(folder))
        if f.lower().endswith(".pdf")
    ]

    if not pdf_files:
        print(f"Keine PDF-Dateien in: {folder}")
        return []

    renamed = []
    for pdf_path in pdf_files:
        result = rename_invoice_pdf(pdf_path, dry_run=dry_run)
        if result:
            renamed.append(result)

    print(f"\n{len(renamed)} Datei(en) {'würden umbenannt' if dry_run else 'umbenannt'}.")
    return renamed


# ---------------------------------------------------------------------------
# Direktaufruf
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    args = sys.argv[1:]
    dry = "--dry-run" in args
    paths = [a for a in args if not a.startswith("--")]

    if not paths:
        print("Verwendung: python Rechnung_umbenennen_1.py <Ordner|PDF> [--dry-run]")
        sys.exit(1)

    target = paths[0]

    if os.path.isdir(target):
        process_folder(target, dry_run=dry)
    elif os.path.isfile(target) and target.lower().endswith(".pdf"):
        rename_invoice_pdf(target, dry_run=dry)
    else:
        print(f"[FEHLER] Ungültiger Pfad: {target}")
        sys.exit(1)
