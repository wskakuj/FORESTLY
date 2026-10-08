"""
Forestly — Mixin: TabGenerowanieValMixin

Zakładka „Generowanie VAL" (sekcja ROZLICZANIE): tworzy pliki .VAL
(rozliczenie geodezyjne) z map GEO-MAP (.MAP).

Metoda: dla każdej działki (pole A2 mapy) szuka wydzieleń (pole A1) leżących
na niej i liczy CZĘŚĆ WSPÓLNĄ (przecięcie poligonów). Tak liczy to GEO-MAP,
a wynik czyta rozliczanie (excel_tasks.wczytaj_i_przetworz_val).

UWAGA: kod działa w OBU wersjach programu — klasycznej (CustomTkinter)
i webowej (PyWebView). Dlatego kontrolki czytamy obronnie (mogą być None),
a komunikaty idą przez self.log()/self.update_status(), które są w obu.
"""

from app.core.leniwe_importy import leniwy_modul
ctk = leniwy_modul("customtkinter")
from pathlib import Path
import threading
import traceback


class TabGenerowanieValMixin:
    """Mixin dla ModernApp oraz WebBackend."""

    # ------------------------------------------------------------------ UI
    def setup_generowanie_val_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)
        font_hint = ctk.CTkFont(family="Segoe UI", size=12)

        card = ctk.CTkFrame(scroll, fg_color="#252526", corner_radius=8,
                            border_width=1, border_color="#333333")
        card.grid(row=0, column=0, padx=20, pady=(20, 15), sticky="ew")
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card, text="Folder z mapami GEO-MAP (.MAP):",
                     font=font_label, text_color="#E0E0E0").grid(
            row=0, column=0, padx=15, pady=(20, 10), sticky="w")
        self.genval_source_entry = ctk.CTkEntry(
            card, placeholder_text="Wskaż folder z plikami .MAP (albo jeden plik)",
            height=36, border_width=1)
        self.genval_source_entry.grid(row=0, column=1, padx=5, pady=(20, 10), sticky="ew")
        _saved = self.get_setting("folder_genval_source")
        if _saved:
            self.genval_source_entry.insert(0, _saved)
        ctk.CTkButton(card, text="Wybierz folder", width=130, height=36,
                      command=lambda: self.select_dir(self.genval_source_entry)
                      ).grid(row=0, column=2, padx=10, pady=(20, 10))

        self.genval_recursive_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(card, text="Uwzględnij podfoldery",
                        variable=self.genval_recursive_var,
                        font=font_hint).grid(row=1, column=1, padx=5, pady=(0, 8), sticky="w")

        ctk.CTkLabel(card, text="Pliki .VAL zapisują się obok każdej mapy "
                                "(ta sama nazwa, końcówka .VAL).",
                     font=font_hint, text_color="#9AA0A6", justify="left").grid(
            row=2, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="w")

        btn_row = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_row.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="ew")
        self.genval_start_btn = ctk.CTkButton(
            btn_row, text="Generuj pliki VAL", height=38, width=200, font=font_btn,
            command=self.start_generowanie_val)
        self.genval_start_btn.pack(side="left")

        self.genval_log = ctk.CTkTextbox(scroll, height=260, font=("Consolas", 11),
                                         fg_color="#1E1E1E", text_color="#D4D4D4")
        self.genval_log.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.genval_log.configure(state="disabled")

    # ------------------------------------------------------------- pomocnicze
    def _genval_pisz(self, tekst):
        """Wypisuje komunikat — do okna logu, jeśli istnieje, i do logu programu.

        Działa w obu wersjach: w webowej self.genval_log nie istnieje.
        """
        try:
            self.log(tekst)
        except Exception:                                   # noqa: BLE001
            pass
        okno = getattr(self, "genval_log", None)
        if okno is None or not hasattr(okno, "insert"):
            return
        def _dopisz():
            try:
                okno.configure(state="normal")
                okno.insert("end", tekst + "\n")
                okno.see("end")
                okno.configure(state="disabled")
            except Exception:                               # noqa: BLE001
                pass
        try:
            self.after(0, _dopisz)
        except Exception:                                   # noqa: BLE001
            _dopisz()

    def _genval_przycisk(self, wlaczony):
        """Włącza/wyłącza przycisk, jeśli istnieje (w webie go nie ma)."""
        btn = getattr(self, "genval_start_btn", None)
        if btn is None or not hasattr(btn, "configure"):
            return
        try:
            btn.configure(state="normal" if wlaczony else "disabled")
        except Exception:                                   # noqa: BLE001
            pass

    # ------------------------------------------------------------- uruchomienie
    def start_generowanie_val(self):
        pole = getattr(self, "genval_source_entry", None)
        zrodlo = pole.get().strip() if pole else ""
        if not zrodlo:
            self._genval_pisz("Najpierw wskaż folder z mapami .MAP.")
            return
        p = Path(zrodlo)
        if not p.exists():
            self._genval_pisz("Ścieżka nie istnieje: %s" % zrodlo)
            return
        self.set_setting("folder_genval_source", zrodlo)
        self._genval_przycisk(False)
        okno = getattr(self, "genval_log", None)
        if okno is not None and hasattr(okno, "delete"):
            try:
                okno.configure(state="normal")
                okno.delete("1.0", "end")
                okno.configure(state="disabled")
            except Exception:                               # noqa: BLE001
                pass
        threading.Thread(target=self._genval_praca, args=(p,), daemon=True).start()

    def _genval_praca(self, sciezka):
        try:
            from app.core import generowanie_val as gv
            self._genval_pisz("=== Generowanie plików VAL z map GEO-MAP ===")
            if sciezka.is_file() and sciezka.suffix.lower() == ".map":
                gv.generuj_val(sciezka, log=self._genval_pisz)
                ile = 1
            else:
                wyniki = gv.generuj_val_folder(sciezka, log=self._genval_pisz)
                ile = sum(1 for w in wyniki if w.get("ok"))
            self._genval_pisz("")
            self._genval_pisz("Zakończono. Utworzono plików .VAL: %d" % ile)
            self.update_status("Generowanie VAL zakończone.", "#28A745", animate=False)
        except Exception as e:                              # noqa: BLE001
            self._genval_pisz("BŁĄD: %s" % e)
            self._genval_pisz(traceback.format_exc())
            try:
                self.update_status("Błąd generowania VAL", "#D83B01", animate=False)
            except Exception:                               # noqa: BLE001
                pass
        finally:
            self._genval_przycisk(True)
