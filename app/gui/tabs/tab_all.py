"""
Forestly — Mixin: TabAllMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import os
import re
import traceback
import shutil
import tempfile
import numpy as np
from docx import Document
import win32com.client
import pythoncom

from app.config import (
    is_file_locked, load_margins, save_margins, add_tooltip,
)

from app.core.word_worker import (
    get_resource_path,
)
from app.core.wydruki import (
    generuj_wszystkie_po_przeniesieniu, generuj_halizny_txt,
)
from app.gui.tabs.tab_wydruki import AGENCJA_NAGLOWKA

class TabAllMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def _setup_all_extras(self, card_frame, row_idx):
        from app.config import TERRITORY_DATA
        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")

        # --- KREATOR STRONY TYTUŁOWEJ (rozwijany) — strona tytułowa i daty ---
        # Strona tytułowa ZAWSZE generowana z tych ustawień (jak w zakładce
        # „Kreator Stron tytułowych", ale bez wsi konkretnej i wiersza powierzchni).
        woj_list = sorted(TERRITORY_DATA.keys()) if TERRITORY_DATA else ["BRAK DANYCH"]
        default_woj = ("KUJAWSKO-POMORSKIE" if "KUJAWSKO-POMORSKIE" in TERRITORY_DATA
                       else (woj_list[0] if woj_list else ""))
        powiat_list = sorted(TERRITORY_DATA.get(default_woj, {}).keys()) or ["BRAK DANYCH"]
        default_powiat = ("TUCHOLSKI" if "TUCHOLSKI" in TERRITORY_DATA.get(default_woj, {})
                          else (powiat_list[0] if powiat_list else ""))
        gmina_list = list(TERRITORY_DATA.get(default_woj, {}).get(default_powiat, []))
        default_gmina = ("LUBIEWO" if "LUBIEWO" in gmina_list
                        else (gmina_list[0] if gmina_list else ""))

        tpl_frame = ctk.CTkFrame(card_frame, fg_color="#1E1E1E",
                                border_width=1, border_color="#333333")
        tpl_frame.grid(row=row_idx, column=0, columnspan=3, padx=15,
                       pady=(0, 10), sticky="ew")
        tpl_frame.grid_columnconfigure(0, weight=1)

        self.all_tpl_open = False
        self.all_tpl_header = ctk.CTkButton(
            tpl_frame, text="▸  Kreator strony tytułowej (strona tytułowa i daty)",
            fg_color="transparent", anchor="w", height=34,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#E0E0E0", hover_color="#252526",
            command=self._toggle_all_tpl_ui)
        self.all_tpl_header.grid(row=0, column=0, padx=5, pady=3, sticky="ew")

        body = ctk.CTkFrame(tpl_frame, fg_color="transparent")
        body.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")
        body.grid_columnconfigure(1, weight=1)
        self.all_tpl_body = body

        def _row(r, label, widget):
            ctk.CTkLabel(body, text=label, font=font_label,
                         text_color="#E0E0E0").grid(
                row=r, column=0, padx=(10, 10), pady=4, sticky="w")
            widget.grid(row=r, column=1, padx=(0, 10), pady=4, sticky="ew")

        self.all_tpl_doc_var = ctk.StringVar(value="UPUL")
        _row(0, "Typ dokumentu:", ctk.CTkOptionMenu(
            body, values=["UPUL", "ISL"], variable=self.all_tpl_doc_var, height=30))
        self.all_tpl_prefix_var = ctk.StringVar(value="położonych na terenie obrębu")
        _row(1, "Prefiks obrębu:", ctk.CTkOptionMenu(
            body, values=["położonych na terenie obrębu", "Obręb:"],
            variable=self.all_tpl_prefix_var, height=30))
        self.all_tpl_woj_var = ctk.StringVar(value=default_woj)
        self.all_tpl_woj_box = ctk.CTkComboBox(
            body, values=woj_list, variable=self.all_tpl_woj_var, height=30,
            command=lambda _v: self._all_tpl_refresh_powiat())
        _row(2, "Województwo (można wpisać własne):", self.all_tpl_woj_box)
        self.all_tpl_powiat_var = ctk.StringVar(value=default_powiat)
        self.all_tpl_powiat_box = ctk.CTkComboBox(
            body, values=powiat_list, variable=self.all_tpl_powiat_var,
            height=30, command=lambda _v: self._all_tpl_refresh_gmina())
        _row(3, "Powiat (można wpisać własny):", self.all_tpl_powiat_box)
        self.all_tpl_gmina_var = ctk.StringVar(value=default_gmina)
        self.all_tpl_gmina_box = ctk.CTkComboBox(
            body, values=gmina_list, variable=self.all_tpl_gmina_var, height=30)
        _row(4, "Gmina (można wpisać własną):", self.all_tpl_gmina_box)

        self.all_tpl_stan_na_entry = ctk.CTkEntry(body, height=30)
        self.all_tpl_stan_na_entry.insert(0, "30.06.2026 r.")
        _row(5, "Stan na (także data we wszystkich Wordach):",
             self.all_tpl_stan_na_entry)
        self.all_tpl_okres_entry = ctk.CTkEntry(body, height=30)
        self.all_tpl_okres_entry.insert(0, "01.01.2027 – 31.12.2036 r.")
        _row(6, "Na okres (strona tytułowa):", self.all_tpl_okres_entry)

        daty_frame = ctk.CTkFrame(body, fg_color="#252526", corner_radius=6)
        daty_frame.grid(row=7, column=0, columnspan=2, padx=10, pady=(4, 2), sticky="ew")
        daty_frame.grid_columnconfigure((1, 3), weight=1)
        ctk.CTkLabel(daty_frame, text="WSK_ZB — 10-lecie:", font=font_label,
                     text_color="#888888").grid(row=0, column=0, padx=(10, 6),
                                                pady=6, sticky="w")
        self.all_wsk_od_entry = ctk.CTkEntry(daty_frame, height=28)
        self.all_wsk_od_entry.insert(0, "01-01-2027")
        self.all_wsk_od_entry.grid(row=0, column=1, padx=4, pady=6, sticky="ew")
        ctk.CTkLabel(daty_frame, text="do:", font=font_label,
                     text_color="#888888").grid(row=0, column=2, padx=4, sticky="w")
        self.all_wsk_do_entry = ctk.CTkEntry(daty_frame, height=28)
        self.all_wsk_do_entry.insert(0, "31-12-2036")
        self.all_wsk_do_entry.grid(row=0, column=3, padx=(4, 10), pady=6, sticky="ew")

        # --- SKRÓTY (własny plik, opcjonalnie) ---
        self.all_custom_skroty_var = ctk.BooleanVar(value=False)
        cb_skroty = ctk.CTkCheckBox(
            card_frame,
            text="Użyj własnego pliku 'Skróty i symbole' (zamiast domyślnego z programu)",
            variable=self.all_custom_skroty_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            command=self._toggle_all_skroty_ui,
        )
        cb_skroty.grid(row=row_idx + 1, column=0, columnspan=3, padx=15,
                       pady=(5, 5), sticky="w")
        self.all_skroty_frame = ctk.CTkFrame(
            card_frame, fg_color="#1E1E1E", border_width=1, border_color="#333333")
        self.all_skroty_frame.grid(row=row_idx + 2, column=0, columnspan=3,
                                    padx=15, pady=(0, 10), sticky="ew")
        self.all_skroty_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.all_skroty_frame, text="Własny plik:", font=font_label,
                     text_color="#E0E0E0").grid(row=0, column=0, padx=(10, 10),
                                                pady=8, sticky="w")
        self.all_skroty_entry = ctk.CTkEntry(
            self.all_skroty_frame, placeholder_text="Wskaż własny plik ze skrótami...",
            height=32)
        self.all_skroty_entry.grid(row=0, column=1, padx=5, pady=8, sticky="ew")
        ctk.CTkButton(
            self.all_skroty_frame, text="Wybierz",
            command=lambda: self.select_file(
                self.all_skroty_entry,
                [("Word/PDF", "*.docx *.doc *.pdf"), ("Wszystkie pliki", "*.*")],
            ),
            width=90, height=32, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=(5, 10), pady=8)

        # --- OPISY OGÓLNE (zawsze z gotowych WSK_ZB.doc z pipeline) ---
        self.all_gen_opis_og_var = ctk.BooleanVar(value=True)
        cb_opis = ctk.CTkCheckBox(
            card_frame,
            text="Generuj opisy ogólne (opis og_<wieś>.docx) po plikach Word",
            variable=self.all_gen_opis_og_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            command=self._toggle_all_gdos_ui,
        )
        cb_opis.grid(row=row_idx + 3, column=0, columnspan=3, padx=15,
                     pady=(5, 5), sticky="w")
        self.all_gdos_frame = ctk.CTkFrame(
            card_frame, fg_color="#1E1E1E", border_width=1, border_color="#333333")
        self.all_gdos_frame.grid(row=row_idx + 4, column=0, columnspan=3,
                                padx=15, pady=(0, 10), sticky="ew")
        self.all_gdos_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.all_gdos_frame,
                     text="Folder z wynikami GDOŚ (opcjonalny):", font=font_label,
                     text_color="#E0E0E0").grid(row=0, column=0, padx=(10, 10),
                                                pady=5, sticky="w")
        self.all_gdos_entry = ctk.CTkEntry(
            self.all_gdos_frame,
            placeholder_text="Formy ochrony przyrody (NN_WIEŚ_wynik.xlsx) — puste = bez GDOŚ",
            height=32)
        self.all_gdos_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(
            self.all_gdos_frame, text="Wybierz",
            command=lambda: self.select_dir(self.all_gdos_entry),
            width=90, height=32, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=(5, 10), pady=5)

        self._toggle_all_skroty_ui()
        self._toggle_all_gdos_ui()
        self._toggle_all_tpl_ui()

        # --- TABELA MARGINESÓW (zwijana) ---
        self._build_margins_ui(card_frame, row_idx + 5, "ALL")

    def _toggle_all_tpl_ui(self):
        self.all_tpl_open = not getattr(self, "all_tpl_open", False)
        if self.all_tpl_open:
            self.all_tpl_body.grid()
            self.all_tpl_header.configure(
                text="▾  Kreator strony tytułowej (strona tytułowa i daty)")
        else:
            self.all_tpl_body.grid_remove()
            self.all_tpl_header.configure(
                text="▸  Kreator strony tytułowej (strona tytułowa i daty)")

    def _all_tpl_refresh_powiat(self):
        from app.config import TERRITORY_DATA
        lista = sorted(TERRITORY_DATA.get(self.all_tpl_woj_var.get(), {}).keys()) \
            or ["BRAK DANYCH"]
        self.all_tpl_powiat_box.configure(values=lista)
        if self.all_tpl_powiat_var.get() not in lista:
            self.all_tpl_powiat_var.set(lista[0])
        self._all_tpl_refresh_gmina()

    def _all_tpl_refresh_gmina(self):
        from app.config import TERRITORY_DATA
        lista = list(TERRITORY_DATA.get(self.all_tpl_woj_var.get(), {})
                     .get(self.all_tpl_powiat_var.get(), [])) or ["BRAK DANYCH"]
        self.all_tpl_gmina_box.configure(values=lista)
        if self.all_tpl_gmina_var.get() not in lista:
            self.all_tpl_gmina_var.set(lista[0])

    def _zbuduj_szablon_str_tyt_dla_all(self):
        """Buduje tymczasowy szablon STR_TYT z ustawień kreatora w 1-Click."""
        def _v(attr, default=""):
            e = getattr(self, attr, None)
            return e.get().strip() if e is not None else default
        try:
            fd, tmp = tempfile.mkstemp(suffix=".docx", prefix="STR_TYT_ALL_")
            os.close(fd)
            out = self._build_str_tyt_doc(
                _v("all_tpl_doc_var", "UPUL") or "UPUL",
                _v("all_tpl_prefix_var") or "położonych na terenie obrębu",
                _v("all_tpl_woj_var").upper(),
                _v("all_tpl_powiat_var").upper(),
                _v("all_tpl_gmina_var").upper(),
                _v("all_tpl_stan_na_entry"),
                _v("all_tpl_okres_entry"),
                tmp, village="NAZWA WSI", keep_area=False)
            self.log("[STR_TYT] Zbudowano szablon bazowy z ustawień kreatora (1-Click).")
            return out
        except Exception:
            self.log("[STR_TYT] Nie udało się zbudować szablonu strony tytułowej:\n"
                     + traceback.format_exc())
            return None

    def _zamien_daty_txt(self, txt_dir):
        """Zamienia daty w wyczyszczonych TXT (przed konwersją na Word).

        'Stan na: <data>' — we wszystkich plikach (pole kreatora „Stan na");
        'w 10-leciu od <data> do <data>' — w plikach WSK_ZB (pola od/do)."""
        def _v(attr):
            e = getattr(self, attr, None)
            return e.get().strip() if e is not None else ""
        stan_na, wsk_od, wsk_do = (_v("all_tpl_stan_na_entry"),
                                   _v("all_wsk_od_entry"), _v("all_wsk_do_entry"))
        if not stan_na and not (wsk_od and wsk_do):
            return
        n_stan = n_wsk = 0
        for f in Path(txt_dir).rglob("*"):
            if not f.is_file() or f.suffix.lower() != ".txt":
                continue
            try:
                raw = f.read_bytes()
            except OSError:
                continue
            nowy = raw
            if stan_na:
                val = stan_na.encode("cp852", errors="replace")
                nowy, k = re.subn(
                    rb"(Stan\s+na:?\s*)[0-9]{2}[-./][0-9]{2}[-./][0-9]{4}(\s*r\.?)?",
                    lambda m, v=val: m.group(1) + v, nowy, flags=re.IGNORECASE)
                n_stan += k
            if wsk_od and wsk_do and f.stem.upper().startswith("WSK_ZB"):
                od = wsk_od.encode("cp852", errors="replace")
                do = wsk_do.encode("cp852", errors="replace")
                nowy, k = re.subn(
                    rb"(w\s+10-leciu\s+od\s+)[0-9][0-9.\-]*[0-9](\s+do\s+)[0-9][0-9.\-]*[0-9]",
                    lambda m, o=od, d=do: m.group(1) + o + m.group(2) + d, nowy)
                n_wsk += k
            if nowy != raw:
                try:
                    f.write_bytes(nowy)
                except OSError:
                    pass
        if n_stan or n_wsk:
            self.log(f"[DATY] Zamieniono \"Stan na\" {n_stan}×, "
                     f"okres 10-lecia WSK_ZB {n_wsk}×.")

    def _toggle_all_gdos_ui(self):
        wlasny = bool(self.all_gen_opis_og_var.get())
        if wlasny:
            self.all_gdos_frame.grid()
        else:
            self.all_gdos_frame.grid_remove()

    def _toggle_all_skroty_ui(self):
        if getattr(self, "all_custom_skroty_var", None) and getattr(self, "all_skroty_frame", None):
            if self.all_custom_skroty_var.get():
                # Jeśli checkbox jest zaznaczony, przywracamy ramkę z powrotem na ekran
                self.all_skroty_frame.grid()
            else:
                # Jeśli checkbox jest odznaczony, całkowicie ukrywamy ramkę
                self.all_skroty_frame.grid_remove()

    def task_generuj_txt(self, in_root):
        """Etap 0: generuje pliki TXT MIETEKA z DBF (jak 'Generowanie: MIETEK -> TXT').

        Szuka folderów z plikami O*.DBF (np. WOL.001) w drzewie źródłowym
        i generuje komplet wydruków w miejscu, obok plików DBF.
        Zwraca liczbę obrębów, dla których coś wygenerowano.
        """
        in_root = Path(in_root)
        # katalogi bezpośrednio zawierające O*.DBF
        kat_o = sorted({p.parent for p in in_root.rglob("*.DBF")
                        if p.name[:1].upper() == "O"})
        if not kat_o:
            self.log("[TXT] Nie znaleziono plików O*.DBF — pomijam generowanie z DBF.")
            return 0
        # obręb = folder nadrzędny (np. CHORZEWO nad WOL.001); gdy DBF-y leżą
        # bezpośrednio w folderze źródłowym, traktujemy go jako obręb
        obreby = sorted({d.parent if d.parent != in_root.parent else d for d in kat_o})
        n = 0
        for obr in obreby:
            self.check_stop()
            try:
                # 1) HALIZNY.TXT (przed przeniesieniem halizn)
                hp, hn = generuj_halizny_txt(obr, agencja=AGENCJA_NAGLOWKA)
                if hn:
                    self.log(f"  [TXT] {obr.name}: HALIZNY.TXT ({hn} wydzieleń)")
                    st, np_ = self.przenies_halizny_obreb(obr)
                    if st == "ok":
                        self.log(f"  [TXT] {obr.name}: przeniesiono halizny "
                                 f"w D*.DBF ({np_} rekordów)")
                    elif st == "blad":
                        self.log(f"  [TXT] {obr.name}: błąd przenoszenia halizn — "
                                 f"wynik może być niepełny.")
                # 2) komplet wydruków (po przeniesieniu halizn)
                out = generuj_wszystkie_po_przeniesieniu(obr, agencja=AGENCJA_NAGLOWKA)
                if out:
                    n += 1
                    self.log(f"  [TXT] {obr.name}: {', '.join(sorted(out))}")
                else:
                    self.log(f"  [TXT] {obr.name}: brak O*.DBF — pomijam.")
            except Exception as e:
                self.log(f"  [TXT] {obr.name}: błąd — {e}")
                traceback.print_exc()
        return n

    def start_pipeline(self, mode):
        src_path = self.entries[mode]["src"].get()
        dst_path = self.entries[mode]["dst"].get()
        remove_names_flag = self.remove_names_var.get()
        if not src_path or not os.path.exists(src_path):
            messagebox.showwarning(
                "Nieprawidłowa ścieżka", "Wybierz istniejący folder źródłowy."
            )
            return
        if not dst_path:
            messagebox.showwarning("Nieprawidłowa ścieżka", "Wybierz folder docelowy.")
            return
        if self.running:
            return
        self.last_output_dir = Path(dst_path)
        self._disable_ui_for_process()
        self.log(f"[{mode}] URUCHOMIENIE ZADANIA\nZ: {src_path}\nDo: {dst_path}")
        self.set_progress(0)

        # Pobieranie i zapis marginesów
        margins_dict = {}
        if mode in ["ALL", "WORD"] and hasattr(self, "margin_vars") and mode in self.margin_vars:
            saved_config = load_margins()  # Odczytujemy stary plik, żeby nie nadpisać układu drugiej zakładki
            if mode not in saved_config:
                saved_config[mode] = {}

            for ftype, entries in self.margin_vars[mode].items():
                try:
                    t = float(entries["T"].get().replace(',', '.'))
                    b = float(entries["B"].get().replace(',', '.'))
                    l = float(entries["L"].get().replace(',', '.'))
                    r = float(entries["R"].get().replace(',', '.'))
                    margins_dict[ftype] = [t, b, l, r]

                    # Wrzucamy do struktury do zapisania na dysku
                    saved_config[mode][ftype] = {"T": t, "B": b, "L": l, "R": r}
                except ValueError:
                    messagebox.showwarning("Błąd", f"Marginesy dla {ftype} muszą być liczbami (np. 1.5).")
                    self.restore_all_buttons()
                    return

            # Zapis całego słownika z marginesami na dysk
            save_margins(saved_config)

        threading.Thread(
            target=self.run_logic_thread,
            args=(src_path, dst_path, mode, remove_names_flag, margins_dict),
            daemon=True,
        ).start()

    def task_generate_str_tyt(self, word_dir, template_path, village_ph, area_ph):
        word_dir = Path(word_dir)
        optax_files = sorted(
            [
                p
                for p in word_dir.rglob("OPTAX*.doc*")
                if p.is_file() and not p.name.startswith("~$")
            ]
        )
        if not optax_files:
            self.log(
                "[STR_TYT] Nie znaleziono plików OPTAX w folderze Word. Pomijam generowanie."
            )
            return

        self.log(
            f"[STR_TYT] Rozpoczynam generowanie stron tytułowych dla {len(optax_files)} wsi..."
        )
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0

        try:
            for optax_path in optax_files:
                self.check_stop()
                if is_file_locked(optax_path):
                    self.log(f"  [Pominięto] Plik zablokowany: {optax_path.name}")
                    continue

                try:
                    doc_word = word_app.Documents.Open(str(optax_path), ReadOnly=True)
                    text_content = doc_word.Content.Text
                    doc_word.Close(SaveChanges=False)

                    village_match = re.search(
                        r"Obiekt:\s*(.+?)(?=\s{2,}|\t|\r|\n|$)",
                        text_content,
                        re.IGNORECASE,
                    )
                    area_match = re.search(
                        r"Razem\s*[^0-9A-Za-z]*([\d\s]+(?:[\.,]\d+)?)",
                        text_content,
                        re.IGNORECASE,
                    )

                    village_name = (
                        village_match.group(1).strip().upper()
                        if village_match
                        else "NIEZNANA_WIES"
                    )
                    area_str = (
                        area_match.group(1).replace(" ", "").strip()
                        if area_match
                        else "[BRAK_DANYCH]"
                    )

                    doc = Document(template_path)
                    self.replace_text_robust(doc, village_ph, village_name)
                    self.replace_text_robust(doc, area_ph, area_str)

                    target_path = optax_path.parent / "STR_TYT.docx"
                    doc.save(str(target_path))
                    self.log(
                        f"  └─ Utworzono: {target_path.parent.name}/STR_TYT.docx (Wieś: {village_name})"
                    )

                except Exception as e:
                    self.log(
                        f"  [Błąd] Nie udało się wygenerować STR_TYT dla {optax_path.parent.name}: {e}"
                    )
        finally:
            try:
                word_app.Quit()
            except:
                pass

    # NOWA METODA: Wstrzykiwanie Skrótów i Symboli do pakietów wsi
    def _resolve_skroty_path(self):
        """Plik ze skrótami: własny (checkbox w Pełnym automacie) albo domyślny z zasobów."""
        if getattr(self, "all_custom_skroty_var", None) and self.all_custom_skroty_var.get():
            entry = getattr(self, "all_skroty_entry", None)
            if entry is not None:
                p = str(entry.get()).strip()
                if p:
                    return p
        # domyślny plik z zasobów programu
        domyslne = get_resource_path("Skroty.pdf")
        if not domyslne.exists():
            domyslne = get_resource_path("Skroty.docx")
        return str(domyslne) if domyslne.exists() else None

    def _inject_skroty_step(self, pdf_dir):
        """Dołącza 'Skróty i symbole' (skroty.pdf) do każdego folderu z PDF-ami."""
        skroty_path = self._resolve_skroty_path()
        if skroty_path and Path(skroty_path).exists():
            c = self.task_inject_skroty(pdf_dir, skroty_path)
            self.log(f"[SKROTY] Dodano plik do {c} folderów wsi.")
            return c
        self.log("[UWAGA] Nie znaleziono pliku ze skrótami (ani domyślnego, ani własnego). Pomijam.")
        return 0

    def task_inject_skroty(self, pdf_dir, skroty_source_path):
        pdf_dir = Path(pdf_dir)
        skroty_source_path = Path(skroty_source_path)

        if not pdf_dir.exists():
            self.log("[SKROTY] Brak folderu PDF — pomijam dołączanie skrótów.")
            return 0

        if not skroty_source_path.exists():
            self.log("[SKROTY] Plik nie istnieje. Pomijam.")
            return 0

        ext = skroty_source_path.suffix.lower()
        temp_skroty_pdf = None
        skroty_pdf_to_copy = None

        if ext in {".doc", ".docx"}:
            self.log("[SKROTY] Konwertuję plik Word na PDF...")
            word_app = None
            try:
                word_app = win32com.client.DispatchEx("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0
                doc = word_app.Documents.Open(str(skroty_source_path.resolve()), AddToRecentFiles=False)
                temp_skroty_pdf = Path(tempfile.gettempdir()) / "skroty_temp.pdf"
                doc.ExportAsFixedFormat(
                    OutputFileName=str(temp_skroty_pdf),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    Range=0,
                    Item=0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )
                doc.Close(False)
                skroty_pdf_to_copy = temp_skroty_pdf
            except Exception as e:
                self.log(f"[SKROTY] Błąd konwersji: {e}")
                return 0
            finally:
                if word_app is not None:
                    try:
                        word_app.Quit()
                    except:
                        pass
        elif ext == ".pdf":
            skroty_pdf_to_copy = skroty_source_path
        else:
            self.log(f"[SKROTY] Nieobsługiwany format: {ext}")
            return 0

        # Znajdź wszystkie foldery które zawierają pliki PDF (oprócz skroty.pdf)
        pdf_folders = set()
        for p in pdf_dir.rglob("*.pdf"):
            if p.name.lower() != "skroty.pdf":
                pdf_folders.add(p.parent)

        count = 0

        if skroty_pdf_to_copy is not None:
            for folder in pdf_folders:
                target_skroty = folder / "skroty.pdf"
                try:
                    shutil.copy2(skroty_pdf_to_copy, target_skroty)
                    count += 1
                except Exception as e:
                    self.log(f"[SKROTY] Błąd kopiowania do {folder.name}: {e}")

        if temp_skroty_pdf and temp_skroty_pdf.exists():
            try:
                temp_skroty_pdf.unlink()
            except:
                pass

        return count

    def _flatten_001_subfolders(self, root_dir):
        """Wyciąga pliki z podfolderów *.001 do folderu nadrzędnego i usuwa puste podfoldery."""
        root_dir = Path(root_dir)
        if not root_dir.exists():
            return
        for sub in sorted(root_dir.rglob("*.001"), reverse=True):
            if not sub.is_dir():
                continue
            for f in sub.iterdir():
                if f.is_file():
                    target = sub.parent / f.name
                    if target.exists():
                        target.unlink()
                    shutil.move(str(f), str(target))
            try:
                sub.rmdir()
                self.log(f"  [SPŁASZCZONO] {sub.parent.name}/{sub.name} → {sub.parent.name}/")
            except Exception:
                pass

    def run_logic_thread(self, src_str, out_str, mode, remove_names, margins_dict=None):
        # --- INICJALIZACJA ZMIENNYCH ---
        in_root = None
        out_root = None
        dir_01, dir_02, dir_03, dir_04, dir_05 = None, None, None, None, None
        # -------------------------------

        pythoncom.CoInitialize()
        try:
            in_root = Path(src_str)
            out_root = Path(out_str)
            out_root.mkdir(parents=True, exist_ok=True)

            if mode == "ALL":
                dir_01, dir_02, dir_03, dir_04, dir_05 = (
                    out_root / "TXT",
                    out_root / "Word",
                    out_root / "PDF",
                    out_root / "PDF Polaczone",
                    out_root / "PDF bez pustych stron",
                )

                self.reset_dashboard()

                # === ETAP 0: GENEROWANIE TXT Z DBF MIETEKA ===
                if getattr(self, "all_gen_txt_var", None) is None or self.all_gen_txt_var.get():
                    self.update_dashboard(0, "running", "Generowanie TXT...")
                    self.check_stop()
                    c0 = self.task_generuj_txt(in_root)
                    self.update_dashboard(0, "done", f"{c0} obrębów")
                else:
                    self.update_dashboard(0, "done", "Pominięto")
                self.set_progress(0.05)

                self.update_dashboard(1, "running", "Czyszczenie...")
                self.check_stop()
                c1 = self.task_clean_txt(in_root, dir_01)
                self._flatten_001_subfolders(dir_01)
                self.update_dashboard(1, "done", f"{c1} plików")
                self.set_progress(0.15)

                if c1 == 0:
                    # nic do roboty — nie kontynuuj (Word/PDF nie mają na czym pracować)
                    self.update_dashboard(1, "error", "Brak TXT")
                    self.log(
                        "\n[BŁĄD] Brak plików TXT do przetworzenia.\n"
                        "Folder źródłowy nie zawiera plików TXT ani plików DBF mietka\n"
                        "(O*.DBF itd.), albo odznaczono 'Generuj pliki TXT z DBF mietka'.")
                    self.update_status("Błąd — brak plików TXT", "#D83B01", animate=False)
                    return

                self.update_dashboard(2, "running", "Kompilacja...")
                self.check_stop()
                self.task_word_processing_subprocess(dir_01, dir_02, remove_names, margins_dict=margins_dict)
                self._flatten_001_subfolders(dir_02)
                self.update_dashboard(2, "done", "Gotowe")
                self.set_progress(0.30)

                # === GENEROWANIE STR_TYT (zawsze — z kreatora w 1-Click) ===
                self.update_status(
                    "Generowanie stron tytułowych (STR_TYT)...", "#0078D7"
                )
                tpl_tmp = self._zbuduj_szablon_str_tyt_dla_all()
                if tpl_tmp:
                    try:
                        self.task_generate_str_tyt(dir_02, tpl_tmp,
                                                   "NAZWA WSI", "wielkość")
                    finally:
                        try:
                            Path(tpl_tmp).unlink()
                        except OSError:
                            pass
                self.set_progress(0.45)

                # === OPISY OGÓLNE (po plikach Word, przed konwersją do PDF) ===
                if (getattr(self, "all_gen_opis_og_var", None) is None
                        or self.all_gen_opis_og_var.get()):
                    self.update_status(
                        "Generowanie opisów ogólnych (opis og_<wieś>.docx)...",
                        "#0078D7",
                    )
                    self.check_stop()
                    gdos_raw = ""
                    ent = getattr(self, "all_gdos_entry", None)
                    if ent is not None:
                        gdos_raw = ent.get().strip()
                    try:
                        self._opis_og_generuj(
                            dir_02, Path(gdos_raw) if gdos_raw else None,
                            tylko_istniejace=True,
                        )
                    except Exception:
                        self.log("[OPIS OG] Błąd generowania opisów ogólnych:\n"
                                 + traceback.format_exc())

                self.update_dashboard(3, "running", "Konwersja...")
                self.check_stop()
                c3 = self.task_convert_to_pdf(dir_02, dir_03)
                self._flatten_001_subfolders(dir_03)
                self.update_dashboard(3, "done", f"{c3} plików")
                self.set_progress(0.60)

                # === WSTRZYKIWANIE SKROTÓW (ZAWSZE WŁĄCZONE) ===
                self.update_status("Dołączanie 'Skrótów i symboli' do pakietów...", "#0078D7")
                self._inject_skroty_step(dir_03)

                self.update_dashboard(4, "running", "Scalanie...")
                self.check_stop()
                c4 = self.task_merge_pdfs(dir_03, dir_04, mode_key="ALL")
                self.update_dashboard(4, "done", f"{c4} pakietów")
                self.set_progress(0.80)

                # usuwanie pustych stron — bez osobnego kroku na dashboardzie
                self.check_stop()
                c5 = self.task_remove_blank_pages(dir_04, dir_05)

                # === PORZĄDKI: zostaje tylko finalny folder "PDF polaczone" ===
                try:
                    if dir_04 and dir_04.exists():
                        shutil.rmtree(dir_04)
                        self.log("[PORZĄDKI] Usunięto folder pośredni 'PDF Polaczone'.")
                except Exception as e:
                    self.log(f"[PORZĄDKI] Nie udało się usunąć 'PDF Polaczone': {e}")
                try:
                    if dir_05 and dir_05.exists():
                        dir_05.rename(out_root / "PDF polaczone")
                        self.log("[PORZĄDKI] Folder 'PDF bez pustych stron' "
                                 "przemianowano na 'PDF polaczone'.")
                except Exception as e:
                    self.log(f"[PORZĄDKI] Nie udało się zmienić nazwy folderu: {e}")

            elif mode == "WORD":
                dir_01, dir_02 = out_root / "TXT", out_root / "Word"
                file_filter = self.get_selected_word_filters()
                filter_label = ", ".join(self.get_selected_word_filters())
                self.check_stop()
                self.update_status(
                    f"ETAP 1/2: Oczyszczanie plików TXT ({filter_label})", "#0078D7"
                )
                cw = self.task_clean_txt(in_root, dir_01, file_filter)
                if cw == 0:
                    self.log(
                        "\n[BŁĄD] Brak plików TXT do przetworzenia "
                        f"dla filtru: {filter_label}.")
                    self.update_status("Błąd — brak plików TXT", "#D83B01", animate=False)
                    return
                self.set_progress(0.5)
                self.check_stop()
                self.update_status(
                    f"ETAP 2/2: Przetwarzanie i konwersja Word ({filter_label})",
                    "#0078D7",
                )
                self.task_word_processing_subprocess(
                    dir_01, dir_02, remove_names, file_filter, margins_dict=margins_dict
                )
                self._flatten_001_subfolders(dir_01)
                self._flatten_001_subfolders(dir_02)

            elif mode == "PDF":
                dir_03, dir_04, dir_05 = (
                    out_root / "PDF",
                    out_root / "PDF Polaczone",
                    out_root / "PDF bez pustych stron",
                )
                # UWAGA: domyślna wartość getattr jest tworzona ZAWSZE, nawet gdy
                # atrybut istnieje — ctk.BooleanVar bez okna Tk (web GUI) rzuca
                # RuntimeError "Too early to create variable", dlatego None + .get().
                _pmv = getattr(self, "pdf_merge_var", None)
                do_merge = True if _pmv is None else bool(_pmv.get())
                # przełącznik: czy dołączać 'Skróty i symbole' (checkbox w GUI)
                _skr = getattr(self, "pdf_skroty_var", None)
                add_skroty = True if _skr is None else bool(_skr.get())
                _total = 4 if (do_merge and add_skroty) else 3

                self.check_stop()
                (
                    self.update_status(
                        f"ETAP 1/{_total}: Zmiana formatu z Word na PDF", "#0078D7"
                    )
                    if do_merge
                    else self.update_status(
                        "Trwa zmiana formatu z Word na PDF...", "#0078D7"
                    )
                )
                self.task_convert_to_pdf(in_root, dir_03)
                self._flatten_001_subfolders(dir_03)
                self.set_progress(0.4 if do_merge else 1.0)
                if do_merge:
                    _e = 1
                    if add_skroty:
                        self.check_stop()
                        _e = 2
                        self.update_status(
                            "ETAP 2/4: Dołączanie 'Skrótów i symboli'", "#0078D7"
                        )
                        self._inject_skroty_step(dir_03)
                        self.set_progress(0.5)
                    self.check_stop()
                    self.update_status(
                        f"ETAP {_e + 1}/{_total}: Logiczna integracja dokumentacji",
                        "#0078D7",
                    )
                    self.task_merge_pdfs(dir_03, dir_04, mode_key="PDF")
                    self.set_progress(0.8)
                    self.check_stop()
                    self.update_status(
                        f"ETAP {_e + 2}/{_total}: Usuwanie anomalii", "#0078D7"
                    )
                    self.task_remove_blank_pages(dir_04, dir_05)

            self.log("\nZAKOŃCZONO POMYŚLNIE.")
            self.set_progress(1.0)
            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            self.after(0, lambda: messagebox.showinfo("Sukces", "Zadanie zakończone."))
        except InterruptedError:
            self.update_status("Przerwano", "#D83B01", animate=False)
            self.log("\nZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
        except Exception as e:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            pythoncom.CoUninitialize()
            self.running = False
            self.after(0, self.restore_all_buttons)

