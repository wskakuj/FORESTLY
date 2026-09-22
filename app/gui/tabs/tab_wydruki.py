"""
Forestly — Mixin: TabWydrukiMixin

Zakładka WYDRUKI — generuje pliki wydrukowe MIETEKA bezpośrednio
z plików DBF (bez uruchamiania programu MS-DOS).
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import traceback

from app.core.wydruki import (generuj_halizny_txt, czytaj_agencje,
                              generuj_wszystkie_po_przeniesieniu)
from app.config import add_tooltip

# Nagłówek wydruków (agencja) — stała, jak w MIETEKU
AGENCJA_NAGLOWKA = 'AGENCJA "CEZAR"'


class TabWydrukiMixin:
    """Mixin dla ModernApp — wydruki generowane z plików DBF mietka."""
    pass

    def _wybrane_wydruki(self):
        """Zaznaczone pliki TXT albo None (= wszystkie)."""
        zm = getattr(self, "wydruki_filter_vars", None)
        if not zm:
            return None
        if zm["Wszystkie"].get():
            return None
        wybrane = [n for n, v in zm.items() if v.get() and n != "Wszystkie"]
        if not wybrane or len(wybrane) == len(zm) - 1:
            return None
        return {f"{n}.TXT" for n in wybrane}


    def on_wydruki_filter_change(self, zmienione):
        """'Wszystkie' wyklucza i BLOKUJE pozostałe check-boxy.

        - zaznaczenie 'Wszystkie' (albo kliknięcie go przy innych zaznaczonych)
          -> pozostałe są odznaczane i blokowane (szare, nieklikalne),
        - odznaczenie 'Wszystkie' -> pozostałe się odblokowują do wyboru,
        - zaznaczenie konkretnego pliku -> 'Wszystkie' się odznacza (jeśli było).
        """
        zm = getattr(self, "wydruki_filter_vars", None)
        cbx = getattr(self, "wydruki_filter_checkboxes", None) or {}
        if not zm:
            return

        def _blokuj(stan):
            for n, cb in cbx.items():
                if n != "Wszystkie":
                    cb.configure(state=stan)

        if zmienione == "Wszystkie":
            if zm["Wszystkie"].get():
                # zaznaczono 'Wszystkie' — odznacz i zablokuj pozostałe
                for n, v in zm.items():
                    if n != "Wszystkie":
                        v.set(False)
                _blokuj("disabled")
            else:
                # odznaczono 'Wszystkie' — odblokuj pozostałe do wyboru
                _blokuj("normal")
        else:
            if zm[zmienione].get():
                # zaznaczono konkretny plik — 'Wszystkie' przestaje obowiązywać
                zm["Wszystkie"].set(False)

    def setup_wydruki_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)
        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)

        card = ctk.CTkFrame(
            scroll_frame, fg_color="#252526", corner_radius=8,
            border_width=1, border_color="#333333",
        )
        card.grid(row=0, column=0, padx=20, pady=(15, 15), sticky="new")
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            card, text="Folder z Mietkami (obręby):",
            font=font_label, text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.wydruki_mietki_entry = ctk.CTkEntry(
            card,
            placeholder_text="Folder, w którym leżą foldery obrębów "
                             "(np. CHORZEWO\\WOL.001\\...DBF)",
            height=36,
        )
        self.wydruki_mietki_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=lambda: self.select_dir(self.wydruki_mietki_entry),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))

        ctk.CTkLabel(
            card, text="Generuj tylko:", font=font_label, text_color="#E0E0E0",
        ).grid(row=1, column=0, padx=15, pady=(8, 8), sticky="nw")
        options_frame = ctk.CTkFrame(card, fg_color="transparent")
        options_frame.grid(row=1, column=1, columnspan=2, padx=5, pady=(8, 8), sticky="w")
        wybory = ["Wszystkie", "OPTAX", "TAB_KLW3", "ZEST1", "REJESTR1",
                  "WSKAZ1", "WYK_NEG", "WSK_ZB", "HALIZNY"]
        self.wydruki_filter_vars = {}
        self.wydruki_filter_checkboxes = {}
        for idx, wyb in enumerate(wybory):
            var = ctk.BooleanVar(value=(wyb == "Wszystkie"))
            cb = ctk.CTkCheckBox(
                options_frame, text=wyb, variable=var,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                fg_color="#0067C0", hover_color="#005A9E",
                state=("normal" if wyb == "Wszystkie" else "disabled"),
                command=lambda c=wyb: self.on_wydruki_filter_change(c),
            )
            cb.grid(row=idx // 4, column=idx % 4, padx=(0, 14), pady=4, sticky="w")
            self.wydruki_filter_vars[wyb] = var
            self.wydruki_filter_checkboxes[wyb] = cb
        add_tooltip(
            options_frame,
            "Możesz zaznaczyć wiele plików naraz. Opcja 'Wszystkie' wyklucza pozostałe.\n"
            "HALIZNY.TXT zaznaczaj tylko PRZED przeniesieniem halizn (zakładka 'Halizny');\n"
            "pozostałe wydruki generuj PO przeniesieniu.",
        )

        ctk.CTkLabel(
            card,
            text="Pliki TXT trafiają tam, gdzie pliki DBF obrębu. HALIZNY.TXT generuj "
                 "PRZED przeniesieniem halizn, pozostałe wydruki PO (zakładka 'Halizny'). "
                 "Dalej przetwarzaj je zakładką 'Konwersja: MIETEK -> Word'.",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#888888",
        ).grid(row=2, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="w")

        self.wydruki_all_btn = ctk.CTkButton(
            scroll_frame,
            text="Generuj wydruki TXT — zaznaczone w 'Generuj tylko:'",
            image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#107C10", hover_color="#0B6A0B", height=44, corner_radius=6,
            command=self.start_wydruki_all,
        )
        self.wydruki_all_btn.grid(row=1, column=0, padx=20, pady=(8, 20), sticky="ew")
        add_tooltip(
            self.wydruki_all_btn,
            "Generuje z plików DBF mietka komplet wydruków MIETEKA:\n"
            "  • OPTAX.TXT — opis lasów i gruntów\n"
            "  • TAB_KLW3.TXT — zestawienie wg klas i podklas wieku\n"
            "  • ZEST1.TXT — skorowidz działek\n"
            "  • REJESTR1.TXT — rejestr działek wg właścicieli\n"
            "  • WSKAZ1.TXT — wykaz wskaźników\n"
            "  • WYK_NEG.TXT — zestawienie drzewostanów negatywnych\n"
            "    i źle produkujących (pusty, gdy brak takich wydzieleń)\n"
            "  • WSK_ZB.TXT — zestawienie czynności gospodarczych na 10-lecie\n"
            "Uruchom PO przeniesieniu halizn (zakładka 'Halizny').\n"
            "Numeracja stron TAB_KLW3 kontynuuje OPTAX — jak w MIETEKU.",
        )

    # ---------------------------------------------------------- POZOSTAŁE WYDRUKI

    def start_wydruki_all(self):
        mietki_dir = self.wydruki_mietki_entry.get().strip() if self.wydruki_mietki_entry else ""
        if not mietki_dir or not Path(mietki_dir).exists():
            messagebox.showwarning("Błąd", "Wybierz istniejący folder z Mietkami.")
            return
        if self.running:
            return
        agencja = AGENCJA_NAGLOWKA
        tylko = self._wybrane_wydruki()
        self.last_output_dir = Path(mietki_dir)
        self._disable_ui_for_process()
        self.log("[WYDRUKI] " + (", ".join(sorted(tylko)) if tylko else "komplet wydruków")
                 + f"\nMIETKI: {mietki_dir}")
        self.set_progress(0)
        threading.Thread(
            target=self.run_wydruki_all_thread,
            args=(mietki_dir, agencja, tylko),
            daemon=True,
        ).start()

    def run_wydruki_all_thread(self, mietki_dir_str, agencja=None, tylko=None):
        try:
            self.update_status("Generowanie wydruków z DBF-ów...", "#0078D7")
            mietki_dir = Path(mietki_dir_str)
            obraby = sorted([d for d in mietki_dir.iterdir() if d.is_dir()])
            if not obraby:
                raise Exception("Brak podfolderów obrębów we wskazanym folderze.")
            total = len(obraby)
            self.start_progress_tracking(total, "Generowanie wydruków MIETEK")

            stat_ok, stat_blad = 0, 0
            lista_ok, lista_brak, lista_blad = [], [], []
            for idx, obr in enumerate(obraby, start=1):
                self.check_stop()
                self.progress_current_file = obr.name
                try:
                    ag = agencja if agencja else (czytaj_agencje(obr) or '')
                    out = {}
                    if tylko is None or "HALIZNY.TXT" in (tylko or set()):
                        try:
                            hp, hn = generuj_halizny_txt(obr, agencja=ag or None)
                            if hp:
                                out["HALIZNY.TXT"] = hp
                            else:
                                self.log(f"  ℹ️ {obr.name}: brak wydzieleń niezalesionych — "
                                         f"HALIZNY.TXT nie powstał.")
                        except Exception:
                            traceback.print_exc()
                    out.update(generuj_wszystkie_po_przeniesieniu(
                        obr, agencja=ag or None, tylko=tylko))
                    if out:
                        stat_ok += 1
                        lista_ok.append(obr.name)
                        for nazwa, p in out.items():
                            self.log(f"  ✅ {obr.name}: {nazwa} → {p}")
                    elif tylko is not None:
                        self.log(f"  ⚠️ {obr.name}: żaden z zaznaczonych plików nie powstał.")
                    else:
                        self.log(f"  ⚠️ {obr.name}: brak O*.DBF — pomijam.")
                        lista_brak.append(obr.name)
                except Exception as e:
                    stat_blad += 1
                    lista_blad.append(obr.name)
                    self.log(f"  ❌ {obr.name}: błąd — {e}")
                    traceback.print_exc()

                self.set_progress(idx / total, current_file=obr.name, current=idx)

            self.log(
                f"\n✅ WYDRUKI: wygenerowano komplet wydruków dla {stat_ok} obrębów"
                + (f"; błędy: {stat_blad}." if stat_blad else ".")
            )
            self.update_status("Gotowe — wydruki wygenerowane.", "#108C4C")

            czesci = [f"Wygenerowano komplet wydruków: {stat_ok} obrębów."]
            if lista_ok:
                czesci.append("• " + ", ".join(lista_ok))
            if lista_brak:
                czesci.append(f"\nBrak pliku O*.DBF (pominięto): {len(lista_brak)}\n"
                              "• " + "\n• ".join(lista_brak))
            if lista_blad:
                czesci.append(f"\nBłędy: {len(lista_blad)} (szczegóły w logu)\n"
                              "• " + "\n• ".join(lista_blad))
            podsumowanie = "\n".join(czesci)
            self.after(0, lambda: messagebox.showinfo(
                "Wydruki — generowanie z DBF", podsumowanie))
        except InterruptedError:
            self.log("\n⛔ ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01")
        except Exception as e:
            self.log(f"\n❌ Błąd WYDRUKI: {e}")
            traceback.print_exc()
            self.update_status("Błąd podczas generowania wydruków.", "#C42B1C")
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)
