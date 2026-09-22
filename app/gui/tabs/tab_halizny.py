"""
Forestly — Mixin: TabHaliznyMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import re
import traceback
import numpy as np

class TabHaliznyMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def _classify_halizna(self, rodzaj):
        """Zwraca nazwę kolumny docelowej dla danego rodzaju powierzchni, albo None."""
        r = (rodzaj or '').lower()
        if 'bagno' in r:
            return 'POW_N_ZAL'
        if 'energetyczna' in r:        # "Linia energetyczna" (odporne na kodowanie)
            return 'POW_INNE'
        if 'halizna' in r:
            return 'POW_L_NZAL'
        if 'azowina' in r:             # "Płazowina"/"plazowina" (odporne na kodowanie)
            return 'POW_L_NZAL'
        return None

    def parse_halizny_txt(self, text, sep):
        """Parsuje HALIZNY.TXT. Zwraca listę słowników:
        {'oddzial','pododdz','kolumna','pow_txt','rodzaj'}.
        Pomija ramki, nagłówki oraz wiersze sum (R.oddz. / Razem)."""
        results = []
        for line in text.splitlines():
            if sep not in line:
                continue
            parts = line.split(sep)
            if len(parts) < 4:
                continue
            col1 = parts[1].strip()   # oddzial+poddz, np. "1gx", "12tx", "16bx"
            col2 = parts[2].strip()   # powierzchnia [ha], np. "0.1915"
            col3 = parts[3].strip()   # rodzaj, np. "241-Halizna"
            if not re.match(r'^\d+[a-zA-Z]*$', col1):
                continue              # odrzuca "R.oddz.", "Razem", nagłówki
            if not re.match(r'^\d+\.\d+$', col2):
                continue              # odrzuca "Pow.", puste kom. sum
            if not re.match(r'^\d+-', col3):
                continue              # odrzuca puste kom. sum / nagłówki
            oddzial = ''.join(ch for ch in col1 if ch.isdigit())
            pododdz = ''.join(ch for ch in col1 if ch.isalpha())
            rodzaj = col3.split('-', 1)[1].strip()
            kolumna = self._classify_halizna(rodzaj)
            if kolumna is None:
                continue
            results.append({
                'oddzial': oddzial, 'pododdz': pododdz,
                'kolumna': kolumna, 'pow_txt': col2, 'rodzaj': rodzaj,
            })
        return results

    def przenies_halizny_obreb(self, obr):
        """Przenosi halizny w D*.DBF JEDNEGO obrębu na podstawie HALIZNY.TXT.

        Używane przez zakładkę 'Halizny' (logika) oraz przez Pełny Automat.
        Zwraca (status, liczba przeniesień):
          'ok'       — zmodyfikowano i zapisano D*.DBF,
          'brak_txt' — brak HALIZNY.TXT,
          'puste'    — HALIZNY.TXT bez wierszy danych,
          'brak_dbf' — brak D*.DBF,
          'nic'      — brak rekordów do przeniesienia (np. już przeniesione),
          'blad'     — błąd odczytu/zapisu.
        Rekordy z POW_L_ZAL równym 0 są pomijane, dzięki czemu
        wielokrotne uruchomienie nie niszczy danych.
        """
        obr = Path(obr)
        # --- 1. Znajdź HALIZNY.TXT ---
        hal_path = None
        for cand in obr.rglob("HALIZNY.TXT"):
            hal_path = cand
            break
        if hal_path is None:
            for cand in obr.rglob("HALIZNY.*"):
                hal_path = cand
                break
        if hal_path is None:
            return ("brak_txt", 0)
        # --- 2. Odczyt z fallbackiem kodowania ---
        raw = hal_path.read_bytes()
        text = raw.decode('cp852', errors='replace')
        wiersze = self.parse_halizny_txt(text, '│')
        if not wiersze:
            text = raw.decode('cp1250', errors='replace')
            wiersze = self.parse_halizny_txt(text, 'ł')
        if not wiersze:
            return ("puste", 0)
        # --- 3. Mapa (oddzial, pododdz) -> (kolumna, pow_txt, rodzaj) ---
        hal_map = {}
        for w in wiersze:
            key = (w['oddzial'], w['pododdz'])
            if key in hal_map and hal_map[key][2] != w['rodzaj']:
                self.log(
                    f"  ⚠️ {obr.name}: pododdział {w['oddzial']}{w['pododdz']} "
                    f"występuje w HALIZNY.TXT wielokrotnie z różnym rodzajem — używam ostatniego."
                )
            hal_map[key] = (w['kolumna'], w['pow_txt'], w['rodzaj'])
        # --- 4. Znajdź D*.DBF ---
        d_dbfs, seen = [], set()
        for pat in ("D*.DBF", "D*.dbf", "d*.DBF", "d*.dbf"):
            for q in obr.rglob(pat):
                k = str(q).upper()
                if k not in seen:
                    seen.add(k)
                    d_dbfs.append(q)
        if not d_dbfs:
            return ("brak_dbf", 0)
        target_dbf = d_dbfs[0]
        # --- 5. Odczyt DBF i indeks rekordów ---
        try:
            fields, records = self.read_dbf(str(target_dbf))
        except Exception as e:
            self.log(f"  ❌ {obr.name}: błąd odczytu {target_dbf.name}: {e}")
            return ("blad", 0)
        idx_map = {}
        for ri, rec in enumerate(records):
            key = (str(rec.get('ODDZIAL', '')).strip(),
                   str(rec.get('PODODDZ', '')).strip())
            idx_map.setdefault(key, []).append(ri)
        # --- 6. Przeniesienie wartości z POW_L_ZAL do właściwej kolumny ---
        przeniesione = 0
        for key, (kolumna, pow_txt, rodzaj) in hal_map.items():
            if key not in idx_map:
                self.log(
                    f"  ⚠️ {obr.name}: halizna {key[0]}{key[1]} ({rodzaj}) "
                    f"nie ma rekordu w {target_dbf.name} — pomijam.")
                continue
            ri_list = idx_map[key]
            for ri in ri_list:
                rec = records[ri]
                val = str(rec.get('POW_L_ZAL', '')).strip()
                if not val:
                    continue
                try:
                    if float(val) == 0.0:
                        continue  # już przeniesione / puste — nie ruszaj
                except ValueError:
                    pass
                rec[kolumna] = val
                rec['POW_L_ZAL'] = '0.0000'
                przeniesione += 1
        if przeniesione == 0:
            return ("nic", 0)
        self.write_dbf(str(target_dbf), fields, records)
        return ("ok", przeniesione)

    def setup_halizny_tab(self, parent):
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
            card, text="Folder z utworzonymi Mietkami:",
            font=font_label, text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.halizny_mietki_entry = ctk.CTkEntry(
            card, placeholder_text="Gdzie leżą foldery obrębów (np. BIAŁCZ\\WOL.001\\HALIZNY.TXT)?",
            height=36,
        )
        self.halizny_mietki_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=lambda: self.select_dir(self.halizny_mietki_entry),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))
        ctk.CTkLabel(
            card,
            text="Reguły: Halizna/Płazowina -> POW_L_NZAL | Bagno -> POW_N_ZAL | "
                 "Linia energetyczna -> POW_INNE.  Wartość brana z POW_L_ZAL (i tam czyszczona).",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#888888",
        ).grid(row=1, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="w")
        self.halizny_start_btn = ctk.CTkButton(
            scroll_frame, text="Przenieś halizny w D*.DBF", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0", hover_color="#005A9E", height=44, corner_radius=6,
            command=self.start_halizny_pipeline,
        )
        self.halizny_start_btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    def start_halizny_pipeline(self):
        mietki_dir = self.halizny_mietki_entry.get().strip() if self.halizny_mietki_entry else ""
        if not mietki_dir or not Path(mietki_dir).exists():
            messagebox.showwarning("Błąd", "Wybierz istniejący folder z utworzonymi Mietkami.")
            return
        if self.running:
            return
        self.last_output_dir = Path(mietki_dir)
        self._disable_ui_for_process()
        self.log(f"[HALIZNY] URUCHOMIENIE\nMIETKI: {mietki_dir}")
        self.set_progress(0)
        threading.Thread(
            target=self.run_halizny_thread, args=(mietki_dir,), daemon=True,
        ).start()

    def run_halizny_thread(self, mietki_dir_str):
        try:
            self.update_status("Przenoszenie halizn w plikach D*.DBF...", "#0078D7")
            mietki_dir = Path(mietki_dir_str)
            obraby = sorted([d for d in mietki_dir.iterdir() if d.is_dir()])
            if not obraby:
                raise Exception("Brak podfolderów obrębów we wskazanym folderze.")
            total = len(obraby)
            self.start_progress_tracking(total, "Przetwarzanie halizn")
            stat_ok = 0
            stat_brak_txt = 0
            stat_brak_dbf = 0
            stat_puste = 0

            for idx, obr in enumerate(obraby, start=1):
                self.check_stop()
                self.progress_current_file = obr.name
                # --- 1. Znajdź HALIZNY.TXT (rekurencyjnie; nazwa podkatalogu bywa różna: WOL.001 / KAM.001 / ...) ---
                hal_path = None
                for cand in obr.rglob("HALIZNY.TXT"):
                    hal_path = cand
                    break
                if hal_path is None:
                    for cand in obr.rglob("HALIZNY.*"):
                        hal_path = cand
                        break
                if hal_path is None:
                    self.log(f"  ⚠️ {obr.name}: brak HALIZNY.TXT — pomijam.")
                    stat_brak_txt += 1
                    self.set_progress(idx / total, current_file=obr.name, current=idx)
                    continue

                # --- 2. Odczyt z fallbackiem kodowania (cp852 -> cp1250) ---
                raw = hal_path.read_bytes()
                text = raw.decode('cp852', errors='replace')
                wiersze = self.parse_halizny_txt(text, '│')
                if not wiersze:
                    text = raw.decode('cp1250', errors='replace')
                    wiersze = self.parse_halizny_txt(text, 'ł')
                if not wiersze:
                    self.log(f"  ℹ️ {obr.name}: HALIZNY.TXT nie zawiera wierszy danych — pomijam.")
                    stat_puste += 1
                    self.set_progress(idx / total, current_file=obr.name, current=idx)
                    continue

                # --- 3. Mapa (oddzial,poddz) -> (kolumna, pow_txt, rodzaj) ---
                hal_map = {}
                for w in wiersze:
                    key = (w['oddzial'], w['pododdz'])
                    if key in hal_map and hal_map[key][2] != w['rodzaj']:
                        self.log(
                            f"  ⚠️ {obr.name}: pododdział {w['oddzial']}{w['pododdz']} "
                            f"występuje w HALIZNY.TXT wielokrotnie z różnym rodzajem — używam ostatniego.")
                    hal_map[key] = (w['kolumna'], w['pow_txt'], w['rodzaj'])

                # --- 4. Znajdź D*.DBF (rekurencyjnie) ---
                d_dbfs = []
                seen = set()
                for p in (list(obr.rglob("D*.DBF")) + list(obr.rglob("D*.dbf")) +
                          list(obr.rglob("d*.DBF")) + list(obr.rglob("d*.dbf"))):
                    k = str(p).upper()
                    if k not in seen:
                        seen.add(k)
                        d_dbfs.append(p)
                if not d_dbfs:
                    self.log(f"  ⚠️ {obr.name}: brak pliku D*.DBF — pomijam.")
                    stat_brak_dbf += 1
                    self.set_progress(idx / total, current_file=obr.name, current=idx)
                    continue
                target_dbf = d_dbfs[0]

                # --- 5. Odczyt DBF i indeks rekordów wg (ODDZIAL,PODODDZ) ---
                try:
                    fields, records = self.read_dbf(str(target_dbf))
                except Exception as e:
                    self.log(f"  ❌ {obr.name}: błąd odczytu {target_dbf.name}: {e}")
                    self.set_progress(idx / total, current_file=obr.name, current=idx)
                    continue
                idx_map = {}
                for ri, rec in enumerate(records):
                    key = (str(rec.get('ODDZIAL', '')).strip(),
                           str(rec.get('PODODDZ', '')).strip())
                    idx_map.setdefault(key, []).append(ri)

                # --- 6. Przeniesienie wartości z POW_L_ZAL do właściwej kolumny ---
                przeniesione = 0
                for key, (kolumna, pow_txt, rodzaj) in hal_map.items():
                    if key not in idx_map:
                        self.log(
                            f"  ⚠️ {obr.name}: halizna {key[0]}{key[1]} ({rodzaj}) "
                            f"nie ma rekordu w {target_dbf.name} — pomijam.")
                        continue
                    ri_list = idx_map[key]
                    # HALIZNY.TXT podaje powierzchnię SUMARYCZNĄ dla pododdziału,
                    # a w DBF ten pododdział może być rozbity na kilka rekordów.
                    # Dlatego diagnostykę robimy na SUMIE, a nie na pojedynczym rekordzie.
                    suma_dbf = 0.0
                    for ri in ri_list:
                        v = str(records[ri].get('POW_L_ZAL', '')).strip()
                        if v:
                            try:
                                suma_dbf += float(v)
                            except Exception:
                                pass
                    # Ostrzeżenie TYLKO przy prawdziwej rozbieżności sum
                    # (czyli gdy pododdział jest tylko CZĘŚCIOWO halizną / danymi niezgodnymi).
                    try:
                        if pow_txt and abs(suma_dbf - float(pow_txt)) > 0.0011:
                            self.log(
                                f"  ⚠️ {obr.name}: rozbieżność SUMY pow. dla {key[0]}{key[1]}: "
                                f"suma DBF={suma_dbf:.4f} vs HALIZNY={pow_txt}")
                    except Exception:
                        pass
                    # Przeniesienie rekord-po-rekordzie (cała POW_L_ZAL -> kolumna docelowa)
                    for ri in ri_list:
                        rec = records[ri]
                        val = str(rec.get('POW_L_ZAL', '')).strip()
                        if not val:
                            self.log(
                                f"  ⚠️ {obr.name}: rekord {key[0]}{key[1]} ma pustą "
                                f"POW_L_ZAL — nie ma czego przenieść.")
                            continue
                        rec[kolumna] = val
                        rec['POW_L_ZAL'] = '0.0000'
                        przeniesione += 1

                if przeniesione == 0:
                    self.log(
                        f"  ℹ️ {obr.name}: brak rekordów do przeniesienia "
                        f"(halizny nie pokrywają się z {target_dbf.name}).")
                    self.set_progress(idx / total, current_file=obr.name, current=idx)
                    continue

                self.write_dbf(str(target_dbf), fields, records)
                self.log(
                    f"  ✅ {obr.name}: przeniesiono {przeniesione} wartości halizn w {target_dbf.name}.")
                stat_ok += 1
                self.set_progress(idx / total, current_file=obr.name, current=idx)

            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            self.log(
                f"\n✅ HALIZNY: zmodyfikowano {stat_ok} obrębów; brak TXT {stat_brak_txt}, "
                f"brak DBF {stat_brak_dbf}, puste {stat_puste} (z {total}).")
            self.after(
                0, lambda: messagebox.showinfo("Sukces", f"Halizny: zmodyfikowano {stat_ok} obrębów."))
        except InterruptedError:
            self.update_status("Przerwano", "#D83B01", animate=False)
            self.log("\nZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
        except Exception as e:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)

    # ==========================================
    # ZAKŁADKA: EXCEL Z MDB
    # ==========================================
