"""
Forestly — Mixin: TabNazwiskaMietekMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import numpy as np

from app.core.word_worker import (
    get_resource_path,
)

class TabNazwiskaMietekMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_nazwiska_mietek_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)

        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)

        card = ctk.CTkFrame(scroll_frame, fg_color="#252526", corner_radius=8, border_width=1, border_color="#333333")
        card.grid(row=0, column=0, padx=20, pady=(15, 15), sticky="new")
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card, text="1. Folder z plikami XLS (Ewidencja):", font=font_label, text_color="#E0E0E0").grid(
            row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.nazwiska_bazowy_entry = ctk.CTkEntry(card,
                                                  placeholder_text="Stąd program pobierze nazwy wsi i wszystkich właścicieli...",
                                                  height=36)
        self.nazwiska_bazowy_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(card, text="Przeglądaj", image=self.icon_folder,
                      command=lambda: self.select_dir(self.nazwiska_bazowy_entry), width=110, height=36, font=font_btn,
                      fg_color="#333333", hover_color="#444444").grid(row=0, column=2, padx=15, pady=(15, 8))

        ctk.CTkLabel(card, text="2. Folder docelowy zapisu:", font=font_label, text_color="#E0E0E0").grid(row=1,
                                                                                                          column=0,
                                                                                                          padx=15,
                                                                                                          pady=(8, 15),
                                                                                                          sticky="w")
        self.nazwiska_out_entry = ctk.CTkEntry(card, placeholder_text="Gdzie wygenerować struktury MIETEK (W*.DBF)?",
                                               height=36)
        self.nazwiska_out_entry.grid(row=1, column=1, padx=5, pady=(8, 15), sticky="ew")
        ctk.CTkButton(card, text="Przeglądaj", image=self.icon_folder,
                      command=lambda: self.select_dir(self.nazwiska_out_entry), width=110, height=36, font=font_btn,
                      fg_color="#333333", hover_color="#444444").grid(row=1, column=2, padx=15, pady=(8, 15))

        # Osobna ramka dla nagłówków DBF
        wsie_frame = ctk.CTkFrame(card, fg_color="#1E1E1E", border_width=1, border_color="#333333")
        wsie_frame.grid(row=2, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="ew")
        wsie_frame.grid_columnconfigure(1, weight=1)
        wsie_frame.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(wsie_frame, text="Dane nagłówka WSIE.DBF (stałe dla całego uruchomienia):",
                     font=font_label, text_color="#A0A0A0").grid(row=0, column=0, columnspan=4, padx=10, pady=(8, 6),
                                                                 sticky="w")

        def _wsie_row(r, c_label, c_entry, label, default, placeholder):
            ctk.CTkLabel(wsie_frame, text=label, font=font_btn, text_color="#E0E0E0").grid(row=r, column=c_label,
                                                                                           padx=(10, 6), pady=4,
                                                                                           sticky="e")
            e = ctk.CTkEntry(wsie_frame, height=30, placeholder_text=placeholder)
            if default:
                e.insert(0, default)
            e.grid(row=r, column=c_entry, padx=(0, 12), pady=4, sticky="ew")
            return e

        self.nm_wsie_wojew_entry = _wsie_row(1, 0, 1, "Województwo (kod):", "10", "np. 10")
        self.nm_wsie_powiat_entry = _wsie_row(1, 2, 3, "Powiat:", "", "np. WYSZKOWSKI")
        self.nm_wsie_stan_entry = _wsie_row(2, 0, 1, "Stan na:", "01.01.2023", "DD.MM.RRRR")
        self.nm_wsie_obod_entry = _wsie_row(2, 2, 3, "Obowiązuje od:", "01.01.2023", "DD.MM.RRRR")
        self.nm_wsie_obdo_entry = _wsie_row(3, 0, 1, "Obowiązuje do:", "31.12.2032", "DD.MM.RRRR")
        self.nm_wsie_nrws_entry = _wsie_row(3, 2, 3, "Nr wsi:", "1", "np. 1")
        self.nm_wsie_rokz_entry = _wsie_row(4, 0, 1, "Rok zal.:", "19", "np. 19")

        _wsie_defaults = {
            "wsie_wojew": "10", "wsie_powiat": "", "wsie_stan": "01.01.2023",
            "wsie_obod": "01.01.2023", "wsie_obdo": "31.12.2032",
            "wsie_nrws": "1", "wsie_rokz": "19",
        }
        for _attr, _default in _wsie_defaults.items():
            _entry = getattr(self, f"nm_{_attr}_entry", None)
            if _entry is not None:
                _saved = self.get_setting(f"wsie_{_attr}", _default)
                if _saved:
                    _entry.delete(0, "end")
                    _entry.insert(0, _saved)
        ctk.CTkLabel(wsie_frame, text="(NAZWA i GMINA = nazwa obrębu, wpisywane automatycznie)",
                     font=ctk.CTkFont(size=11), text_color="#777777").grid(row=4, column=2, columnspan=2, padx=(0, 12),
                                                                           pady=4, sticky="w")

        self.nazwiska_mietek_start_btn = ctk.CTkButton(
            scroll_frame, text="Generuj struktury (tylko Ewidencja)", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0", hover_color="#005A9E", height=44, corner_radius=6,
            command=self.start_nazwiska_mietek_pipeline
        )
        self.nazwiska_mietek_start_btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    def start_nazwiska_mietek_pipeline(self):
        baz_dir = self.nazwiska_bazowy_entry.get().strip() if hasattr(self,
                                                                      'nazwiska_bazowy_entry') and self.nazwiska_bazowy_entry else ""
        out_dir = self.nazwiska_out_entry.get().strip() if hasattr(self,
                                                                   'nazwiska_out_entry') and self.nazwiska_out_entry else ""

        if not baz_dir or not Path(baz_dir).exists():
            messagebox.showwarning("Błąd", "Wybierz główny folder z plikami XLS (Ewidencja).")
            return
        if not out_dir:
            messagebox.showwarning("Błąd", "Wybierz folder docelowy dla nowych obrębów.")
            return

        baz_path = Path(baz_dir)
        xls_files = [
            f.stem for f in baz_path.iterdir()
            if f.is_file() and f.suffix.lower() in {'.xls', '.xlsx'} and not f.name.startswith("~$")
        ]

        if not xls_files:
            messagebox.showwarning("Błąd",
                                   "We wskazanym folderze XLS Ewidencji nie znaleziono żadnych plików, z których można by pobrać nazwy obrębów.")
            return

        names_list = sorted(list(set(xls_files)))

        base_dir = get_resource_path("pusty")
        if not Path(base_dir).exists() or not Path(base_dir).is_dir():
            messagebox.showerror("Błąd",
                                 f"Nie znaleziono wbudowanego folderu 'pusty' w plikach programu!\nŚcieżka: {base_dir}")
            return

        if self.running: return
        self.last_output_dir = Path(out_dir)
        self._disable_ui_for_process()
        self.set_progress(0)

        self.set_setting("wsie_wsie_wojew", self.nm_wsie_wojew_entry.get().strip())
        self.set_setting("wsie_wsie_powiat", self.nm_wsie_powiat_entry.get().strip())
        self.set_setting("wsie_wsie_stan", self.nm_wsie_stan_entry.get().strip())
        self.set_setting("wsie_wsie_obod", self.nm_wsie_obod_entry.get().strip())
        self.set_setting("wsie_wsie_obdo", self.nm_wsie_obdo_entry.get().strip())
        self.set_setting("wsie_wsie_nrws", self.nm_wsie_nrws_entry.get().strip())
        self.set_setting("wsie_wsie_rokz", self.nm_wsie_rokz_entry.get().strip())

        wsie_meta = {
            'WOJEW': self.nm_wsie_wojew_entry.get().strip(),
            'POWIAT': self.nm_wsie_powiat_entry.get().strip(),
            'STAN_NA': self.nm_wsie_stan_entry.get().strip(),
            'OBOW_OD': self.nm_wsie_obod_entry.get().strip(),
            'OBOW_DO': self.nm_wsie_obdo_entry.get().strip(),
            'NR_WSI': self.nm_wsie_nrws_entry.get().strip() or "1",
            'ROK_ZAL': self.nm_wsie_rokz_entry.get().strip(),
        }
        if not wsie_meta['POWIAT']:
            self.log("[UWAGA] Pole 'Powiat' w danych WSIE.DBF jest puste — uzupełnij je, jeśli MIETEK go wymaga.")

        # Wywołanie znanego wątku, ale z pominięciem parametru rozl_dir
        threading.Thread(
            target=self.run_tworzenie_mietkow_thread,
            args=(base_dir, out_dir, names_list, baz_dir, None, wsie_meta),
            daemon=True
        ).start()

    # ==========================================
    # ZAKŁADKA: WYKAZ ROZBIEŻNOŚCI (MIETEK D*.DBF vs EXCEL)
    # ==========================================
