"""
Forestly — Mixin: TabKrzyzowkiMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import re
import warnings
warnings.filterwarnings("ignore", message=".*OLE2 inconsistency.*")
warnings.filterwarnings("ignore", message=".*file size.*not.*sector size.*")
warnings.filterwarnings("ignore", message=".*SSCS size.*")
import traceback
import pandas as pd
import numpy as np

class TabKrzyzowkiMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_krzyzowki_tab(self, parent):
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
            card, text="1. Folder z poprawionymi XLSX (rozliczonymi):",
            font=font_label, text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.krzyz_xls_entry = ctk.CTkEntry(
            card, placeholder_text="Wskaż folder z ręcznie zredagowanymi plikami *_Rozliczone.xlsx...",
            height=36,
        )
        self.krzyz_xls_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=lambda: self.select_dir(self.krzyz_xls_entry),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))

        ctk.CTkLabel(
            card, text="2. Folder z utworzonymi Mietkami:",
            font=font_label, text_color="#E0E0E0",
        ).grid(row=1, column=0, padx=15, pady=(8, 8), sticky="w")
        self.krzyz_mietki_entry = ctk.CTkEntry(
            card, placeholder_text="Gdzie leżą foldery obrębów (np. BIAŁCZ\\WOL.001)?",
            height=36,
        )
        self.krzyz_mietki_entry.grid(row=1, column=1, padx=5, pady=(8, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=lambda: self.select_dir(self.krzyz_mietki_entry),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=1, column=2, padx=15, pady=(8, 8))

        ctk.CTkLabel(
            card,
            text="Do D*.DBF trafią: J. rej. -> NRREJ, nr_dz -> NR_DZIAL, kolumna F -> POW i POW_L_ZAL, "
                 "cyfry z 'litery' -> ODDZIAL, litery z 'litery' -> PODODDZ.",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#888888",
        ).grid(row=2, column=0, columnspan=3, padx=15, pady=(0, 5), sticky="w")

        self.krzyz_usun_puste_jrej_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            card, text="Usuń wiersze jeśli w krzyżówce brakuje wartości w J. rej. (usuwanie wydzieleń bez właścicieli)",
            variable=self.krzyz_usun_puste_jrej_var, font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#8B0000", hover_color="#A52A2A"
        ).grid(row=3, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="w")

        self.krzyz_start_btn = ctk.CTkButton(
            scroll_frame, text="Wstrzyknij krzyżówki do DBF", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0", hover_color="#005A9E", height=44, corner_radius=6,
            command=self.start_krzyzowki_pipeline,
        )
        self.krzyz_start_btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    def start_krzyzowki_pipeline(self):
        xls_dir = self.krzyz_xls_entry.get().strip() if self.krzyz_xls_entry else ""
        mietki_dir = self.krzyz_mietki_entry.get().strip() if self.krzyz_mietki_entry else ""
        if not xls_dir or not Path(xls_dir).exists():
            messagebox.showwarning("Błąd", "Wybierz istniejący folder z poprawionymi plikami XLSX.")
            return
        if not mietki_dir or not Path(mietki_dir).exists():
            messagebox.showwarning("Błąd", "Wybierz istniejący folder z utworzonymi Mietkami.")
            return
        if self.running:
            return

        usun_puste_jrej = self.krzyz_usun_puste_jrej_var.get()

        self.last_output_dir = Path(mietki_dir)
        self._disable_ui_for_process()
        self.log(f"[KRZYŻÓWKI] URUCHOMIENIE\nXLSX: {xls_dir}\nMIETKI: {mietki_dir}")
        self.set_progress(0)
        threading.Thread(
            target=self.run_krzyzowki_thread, args=(xls_dir, mietki_dir, usun_puste_jrej), daemon=True,
        ).start()

    def run_krzyzowki_thread(self, xls_dir_str, mietki_dir_str, usun_puste_jrej=False):
        try:
            self.update_status("Wstrzykiwanie krzyżówek do plików D*.DBF...", "#0078D7")
            xls_dir = Path(xls_dir_str)
            mietki_dir = Path(mietki_dir_str)
            xls_files = sorted([
                f for f in xls_dir.iterdir()
                if f.is_file() and f.suffix.lower() in {".xls", ".xlsx"} and not f.name.startswith("~$")
            ])
            if not xls_files:
                raise Exception("Brak plików Excel we wskazanym folderze.")

            total = len(xls_files)
            self.start_progress_tracking(total, "Wpisywanie krzyżówek")

            # Struktura D*.DBF wg MIETEK.EXE — NRREJ MUSI być pierwsze, inaczej program nie ruszy!
            dbf_fields = [
                ('NRREJ', 'N', 5, 0),  # J. rej.  (kolumna B w XLSX)
                ('NR_DZIAL', 'C', 9, 0),  # nr_dz
                ('POW', 'N', 9, 4),  # kolumna F
                ('POW_L_ZAL', 'N', 9, 4),  # kolumna F
                ('POW_L_NZAL', 'N', 8, 4),  # puste
                ('POW_N_ZAL', 'N', 9, 4),  # puste
                ('POW_INNE', 'N', 8, 4),  # puste
                ('ODDZIAL', 'C', 7, 0),  # cyfry z 'litery'
                ('PODODDZ', 'C', 3, 0),  # litery z 'litery'
                ('ZM', 'C', 1, 0),  # puste
                ('PREJ', 'N', 6, 0),  # puste
            ]

            stat_ok = 0
            stat_brak_folderu = 0
            stat_puste = 0

            for idx, xls_path in enumerate(xls_files, start=1):
                self.check_stop()
                self.progress_current_file = xls_path.name

                # Nazwa obrębu = nazwa pliku bez przyrostka "_Rozliczone" (i wszystkiego po nim)
                v_name = re.sub(r'(?i)_?rozliczone.*$', '', xls_path.stem).strip()
                if not v_name:
                    v_name = xls_path.stem
                v_norm = re.sub(r'[\s_\-]', '', v_name.lower())

                # Szukamy pasującego folderu obrębu wśród Mietków (ściśle, bez fałszywych "LIS"/"LISIE POLE")
                target_mietek = None
                for folder in mietki_dir.iterdir():
                    if folder.is_dir():
                        f_norm = re.sub(r'[\s_\-]', '', folder.name.lower())
                        if f_norm and f_norm == v_norm:
                            target_mietek = folder
                            break
                if not target_mietek:
                    self.log(f"  ⚠️ Pominięto {xls_path.name} — nie znaleziono folderu obrębu '{v_name}' w Mietkach.")
                    stat_brak_folderu += 1
                    self.set_progress(idx / total, current_file=xls_path.name, current=idx)
                    continue

                try:
                    df = pd.read_excel(str(xls_path))
                    if df.shape[1] < 6:
                        self.log(f"  ❌ {xls_path.name}: plik ma mniej niż 6 kolumn — nie mogę odczytać kolumny F.")
                        self.set_progress(idx / total, current_file=xls_path.name, current=idx)
                        continue

                    # Powierzchnia ZAWSZE z kolumny F (tam użytkownik redaguje krzyżówki)
                    col_pow_name = df.columns[5]
                    df_pow = pd.to_numeric(df.iloc[:, 5], errors='coerce')
                    df_work = df.copy()
                    df_work['__POW'] = df_pow
                    df_filt = df_work[df_work['__POW'].notna()]

                    if df_filt.empty:
                        self.log(
                            f"  ℹ️ {xls_path.name}: kolumna F ('{col_pow_name}') pusta — brak krzyżówek do wpisania.")
                        stat_puste += 1
                        self.set_progress(idx / total, current_file=xls_path.name, current=idx)
                        continue

                    records = []
                    for _, row in df_filt.iterrows():
                        try:
                            nrrej_val = int(float(row.get('J. rej.', 0)))
                        except Exception:
                            nrrej_val = 0

                        # NOWA LOGIKA: Jeśli zaznaczono usuwanie i brakuje nr rejestru (0), pomiń wiersz!
                        if usun_puste_jrej and nrrej_val == 0:
                            continue

                        nr_dz = str(row.get('nr_dz', '')).strip()
                        litery = str(row.get('litery', ''))
                        # "X" w litery → ODDZIAL (kolumna H), nie PODODDZ (kolumna I)
                        if litery.strip().upper() == 'X':
                            oddzial = 'X'
                            pododdz = ''
                        else:
                            oddzial = "".join(ch for ch in litery if ch.isdigit())[:7]  # cyfry  -> ODDZIAL
                            pododdz = "".join(ch for ch in litery if ch.isalpha())[:3]  # litery -> PODODDZ
                        pow_val = row['__POW']
                        records.append({
                            'NRREJ': nrrej_val,  # <-- KONIECZNIE, jako pierwsze
                            'NR_DZIAL': nr_dz[:9],
                            'POW': f"{float(pow_val):.4f}",
                            'POW_L_ZAL': f"{float(pow_val):.4f}",
                            # POW_L_NZAL / POW_N_ZAL / POW_INNE / ZM / PREJ celowo POMIJAMY
                            # -> write_dbf zapisze je jako PUSTE (zgodnie ze strukturą MIETEK.EXE)
                            'ODDZIAL': oddzial,
                            'PODODDZ': pododdz,
                        })

                    # Szukamy istniejącego D*.DBF rekurencyjnie (nazwa podkatalogu bywa różna)
                    d_dbfs = []
                    seen = set()
                    for p in (list(target_mietek.rglob("D*.DBF")) + list(target_mietek.rglob("D*.dbf")) +
                              list(target_mietek.rglob("d*.DBF")) + list(target_mietek.rglob("d*.dbf"))):
                        key = str(p).upper()
                        if key not in seen:
                            seen.add(key)
                            d_dbfs.append(p)
                    if d_dbfs:
                        target_dbf = d_dbfs[0]
                    else:
                        # brak DBF -> zapisz do istniejącego podkatalogu *.001, a gdy go nie ma: WOL.001
                        sub = self._find_001_dir(target_mietek)
                        if sub is None:
                            sub = target_mietek / "WOL.001"
                        sub.mkdir(parents=True, exist_ok=True)
                        target_dbf = sub / "D0011019.DBF"
                    self.write_dbf(str(target_dbf), dbf_fields, records)
                    self.log(
                        f"  ✅ {xls_path.name} → {target_mietek.name}/{target_dbf.name} "
                        f"({len(records)} rekordów, kolumna F='{col_pow_name}')"
                    )
                    stat_ok += 1
                except Exception as e:
                    self.log(f"  ❌ Błąd przetwarzania {xls_path.name}: {e}")

                self.set_progress(idx / total, current_file=xls_path.name, current=idx)

            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            self.log(
                f"\n✅ KRZYŻÓWKI: zapisano {stat_ok}, puste {stat_puste}, "
                f"brak folderu {stat_brak_folderu} (z {total})."
            )
            self.after(
                0, lambda: messagebox.showinfo("Sukces", f"Wstrzyknięto krzyżówki do {stat_ok} obrębów.")
            )
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
    # ZAKŁADKA: HALIZNY (HALIZNY.TXT -> D*.DBF)
    # ==========================================
