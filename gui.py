"""
gui.py
======
Grafische Benutzeroberfläche für das PS-Rechnungstool.

Funktionen:
  - PDF-Rechnungen umbenennen (Rechnung_umbenennen_1)
  - Rechnungsdaten in Excel exportieren (Rechnungsdatenübernahme)
  - Kunden-E-Mails mit Firmenstamm abgleichen
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# Interne Module
import Rechnung_umbenennen_1 as renamer
import Rechnungsdatenübernahme as extractor
import compare_excel_customer_and_partial_company_email_update as email_updater


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def _run_in_thread(fn, *args, **kwargs):
    """Führt *fn* in einem Daemon-Thread aus, damit die GUI nicht blockiert."""
    t = threading.Thread(target=fn, args=args, kwargs=kwargs, daemon=True)
    t.start()


# ---------------------------------------------------------------------------
# Haupt-GUI-Klasse
# ---------------------------------------------------------------------------

class PSRechnungstoolApp(tk.Tk):
    """Hauptfenster des PS-Rechnungstools."""

    APP_TITLE = "PS Rechnungstool"
    MIN_WIDTH = 750
    MIN_HEIGHT = 550

    def __init__(self):
        super().__init__()
        self.title(self.APP_TITLE)
        self.minsize(self.MIN_WIDTH, self.MIN_HEIGHT)
        self.resizable(True, True)

        # Zustandsvariablen
        self._pdf_folder = tk.StringVar()
        self._excel_file = tk.StringVar()
        self._customer_excel = tk.StringVar()
        self._company_excel = tk.StringVar()
        self._output_excel = tk.StringVar()

        self._build_ui()

    # ------------------------------------------------------------------
    # UI-Aufbau
    # ------------------------------------------------------------------

    def _build_ui(self):
        """Baut die Benutzeroberfläche auf."""
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Tab 1: Umbenennen
        tab_rename = ttk.Frame(notebook)
        notebook.add(tab_rename, text="PDF umbenennen")
        self._build_rename_tab(tab_rename)

        # Tab 2: Datenübernahme
        tab_extract = ttk.Frame(notebook)
        notebook.add(tab_extract, text="Daten → Excel")
        self._build_extract_tab(tab_extract)

        # Tab 3: E-Mail-Abgleich
        tab_email = ttk.Frame(notebook)
        notebook.add(tab_email, text="E-Mail-Abgleich")
        self._build_email_tab(tab_email)

        # Log-Bereich (gemeinsam)
        log_frame = ttk.LabelFrame(self, text="Protokoll")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self._log = scrolledtext.ScrolledText(
            log_frame, height=10, state=tk.DISABLED, wrap=tk.WORD
        )
        self._log.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        btn_clear = ttk.Button(log_frame, text="Protokoll leeren", command=self._clear_log)
        btn_clear.pack(anchor=tk.E, padx=4, pady=(0, 4))

    # --- Tab: Umbenennen ---

    def _build_rename_tab(self, parent: ttk.Frame):
        pad = {"padx": 8, "pady": 4}

        ttk.Label(parent, text="PDF-Ordner:").grid(row=0, column=0, sticky=tk.W, **pad)
        ttk.Entry(parent, textvariable=self._pdf_folder, width=55).grid(
            row=0, column=1, sticky=tk.EW, **pad
        )
        ttk.Button(parent, text="Durchsuchen…", command=self._browse_pdf_folder).grid(
            row=0, column=2, **pad
        )

        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=8)

        ttk.Button(
            btn_frame,
            text="Umbenennen (Vorschau)",
            command=lambda: _run_in_thread(self._run_rename, dry_run=True),
        ).pack(side=tk.LEFT, padx=4)

        ttk.Button(
            btn_frame,
            text="Umbenennen (ausführen)",
            command=lambda: _run_in_thread(self._run_rename, dry_run=False),
        ).pack(side=tk.LEFT, padx=4)

        parent.columnconfigure(1, weight=1)

    # --- Tab: Datenübernahme ---

    def _build_extract_tab(self, parent: ttk.Frame):
        pad = {"padx": 8, "pady": 4}

        ttk.Label(parent, text="PDF-Ordner:").grid(row=0, column=0, sticky=tk.W, **pad)
        ttk.Entry(parent, textvariable=self._pdf_folder, width=55).grid(
            row=0, column=1, sticky=tk.EW, **pad
        )
        ttk.Button(parent, text="Durchsuchen…", command=self._browse_pdf_folder).grid(
            row=0, column=2, **pad
        )

        ttk.Label(parent, text="Excel-Zieldatei:").grid(row=1, column=0, sticky=tk.W, **pad)
        ttk.Entry(parent, textvariable=self._excel_file, width=55).grid(
            row=1, column=1, sticky=tk.EW, **pad
        )
        ttk.Button(parent, text="Speichern unter…", command=self._browse_excel_save).grid(
            row=1, column=2, **pad
        )

        ttk.Button(
            parent,
            text="Daten extrahieren",
            command=lambda: _run_in_thread(self._run_extract),
        ).grid(row=2, column=0, columnspan=3, pady=8)

        parent.columnconfigure(1, weight=1)

    # --- Tab: E-Mail-Abgleich ---

    def _build_email_tab(self, parent: ttk.Frame):
        pad = {"padx": 8, "pady": 4}

        ttk.Label(parent, text="Kunden-Excel:").grid(row=0, column=0, sticky=tk.W, **pad)
        ttk.Entry(parent, textvariable=self._customer_excel, width=55).grid(
            row=0, column=1, sticky=tk.EW, **pad
        )
        ttk.Button(parent, text="Öffnen…", command=self._browse_customer_excel).grid(
            row=0, column=2, **pad
        )

        ttk.Label(parent, text="Firmen-Excel:").grid(row=1, column=0, sticky=tk.W, **pad)
        ttk.Entry(parent, textvariable=self._company_excel, width=55).grid(
            row=1, column=1, sticky=tk.EW, **pad
        )
        ttk.Button(parent, text="Öffnen…", command=self._browse_company_excel).grid(
            row=1, column=2, **pad
        )

        ttk.Label(parent, text="Ausgabe-Excel (opt.):").grid(
            row=2, column=0, sticky=tk.W, **pad
        )
        ttk.Entry(parent, textvariable=self._output_excel, width=55).grid(
            row=2, column=1, sticky=tk.EW, **pad
        )
        ttk.Button(parent, text="Speichern unter…", command=self._browse_output_excel).grid(
            row=2, column=2, **pad
        )

        ttk.Button(
            parent,
            text="E-Mails abgleichen",
            command=lambda: _run_in_thread(self._run_email_update),
        ).grid(row=3, column=0, columnspan=3, pady=8)

        parent.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    # Dateidialoge
    # ------------------------------------------------------------------

    def _browse_pdf_folder(self):
        path = filedialog.askdirectory(title="PDF-Ordner wählen")
        if path:
            self._pdf_folder.set(path)

    def _browse_excel_save(self):
        path = filedialog.asksaveasfilename(
            title="Excel-Datei speichern",
            defaultextension=".xlsx",
            filetypes=[("Excel-Datei", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self._excel_file.set(path)

    def _browse_customer_excel(self):
        path = filedialog.askopenfilename(
            title="Kunden-Excel öffnen",
            filetypes=[("Excel-Datei", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self._customer_excel.set(path)

    def _browse_company_excel(self):
        path = filedialog.askopenfilename(
            title="Firmen-Excel öffnen",
            filetypes=[("Excel-Datei", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self._company_excel.set(path)

    def _browse_output_excel(self):
        path = filedialog.asksaveasfilename(
            title="Ausgabe-Excel speichern",
            defaultextension=".xlsx",
            filetypes=[("Excel-Datei", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self._output_excel.set(path)

    # ------------------------------------------------------------------
    # Aktionen (laufen in Threads)
    # ------------------------------------------------------------------

    def _run_rename(self, dry_run: bool = False):
        folder = self._pdf_folder.get().strip()
        if not folder:
            messagebox.showwarning("Hinweis", "Bitte einen PDF-Ordner auswählen.")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Fehler", f"Ordner nicht gefunden:\n{folder}")
            return

        self._log_redirect(lambda: renamer.process_folder(folder, dry_run=dry_run))

    def _run_extract(self):
        folder = self._pdf_folder.get().strip()
        excel = self._excel_file.get().strip()
        if not folder or not excel:
            messagebox.showwarning("Hinweis", "Bitte Ordner und Excel-Datei angeben.")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Fehler", f"Ordner nicht gefunden:\n{folder}")
            return

        self._log_redirect(lambda: extractor.process_folder(folder, excel))

    def _run_email_update(self):
        cust = self._customer_excel.get().strip()
        comp = self._company_excel.get().strip()
        out = self._output_excel.get().strip() or None
        if not cust or not comp:
            messagebox.showwarning("Hinweis", "Bitte Kunden- und Firmen-Excel angeben.")
            return

        self._log_redirect(lambda: email_updater.compare_and_update_emails(cust, comp, out))

    # ------------------------------------------------------------------
    # Log-Hilfsmethoden
    # ------------------------------------------------------------------

    def _log_redirect(self, fn):
        """Leitet print-Ausgaben von *fn* ins Log-Widget um."""
        import io
        import contextlib

        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf):
                fn()
        except Exception as exc:  # noqa: BLE001
            buf.write(f"\n[FEHLER] {exc}\n")
            messagebox.showerror("Fehler", str(exc))
        finally:
            self._append_log(buf.getvalue())

    def _append_log(self, text: str):
        """Hängt *text* ans Log-Widget an (thread-safe via after)."""
        def _do():
            self._log.configure(state=tk.NORMAL)
            self._log.insert(tk.END, text)
            self._log.see(tk.END)
            self._log.configure(state=tk.DISABLED)

        self.after(0, _do)

    def _clear_log(self):
        self._log.configure(state=tk.NORMAL)
        self._log.delete("1.0", tk.END)
        self._log.configure(state=tk.DISABLED)


# ---------------------------------------------------------------------------
# Einstiegspunkt
# ---------------------------------------------------------------------------

def main():
    app = PSRechnungstoolApp()
    app.mainloop()


if __name__ == "__main__":
    main()
