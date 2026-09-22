"""
Forestly — Mixin: TabTworzenieMietkowMixin
Połączona zakładka: Tworzenie i wpisywanie mietków.
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import re
import shutil
import warnings
warnings.filterwarnings("ignore", message=".*OLE2 inconsistency.*")
warnings.filterwarnings("ignore", message=".*file size.*not.*sector size.*")
warnings.filterwarnings("ignore", message=".*SSCS size.*")
import traceback
import pandas as pd
import numpy as np

from app.core.word_worker import (
    get_resource_path,
)

from app.core.excel_tasks import (
    wczytaj_i_przetworz_wlascicieli, WSIE_FIELDS,
)

class TabTworzenieMietkowMixin:
    """Mixin dla ModernApp — łączy tworzenie mietków i wpisywanie krzyżówek."""
    pass

    def setup_tworzenie_mietkow_tab(self, parent):
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

        ctk.CTkLabel(card, text="1. Folder Główny XLS (Ewidencja):", font=font_label, text_color="#E0E0E0").grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.mietki_bazowy_entry = ctk.CTkEntry(card, placeholder_text="Stąd program pobierze nazwy wsi i właścicieli...", height=36)
        _saved = self.get_setting("folder_mietki_bazowy_entry")
        if _saved: self.mietki_bazowy_entry.insert(0, _saved)
        self.mietki_bazowy_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(card, text="Przeglądaj", image=self.icon_folder, command=lambda: self.select_dir(self.mietki_bazowy_entry), width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444").grid(row=0, column=2, padx=15, pady=(15, 8))

        ctk.CTkLabel(card, text="2. Folder XLSX (Rozliczone):", font=font_label, text_color="#E0E0E0").grid(row=1, column=0, padx=15, pady=8, sticky="w")
        self.mietki_rozlicz_entry = ctk.CTkEntry(card, placeholder_text="Stąd program pobierze numery J.rej i krzyżówki...", height=36)
        _saved = self.get_setting("folder_mietki_rozlicz_entry")
        if _saved: self.mietki_rozlicz_entry.insert(0, _saved)
        self.mietki_rozlicz_entry.grid(row=1, column=1, padx=5, pady=8, sticky="ew")
        ctk.CTkButton(card, text="Przeglądaj", image=self.icon_folder, command=lambda: self.select_dir(self.mietki_rozlicz_entry), width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444").grid(row=1, column=2, padx=15, pady=8)

        ctk.CTkLabel(card, text="3. Folder docelowy zapisu:", font=font_label, text_color="#E0E0E0").grid(row=2, column=0, padx=15, pady=(8, 15), sticky="w")
        self.mietki_out_entry = ctk.CTkEntry(card, placeholder_text="Gdzie zapisać gotowe struktury MS-DOS z bazą DBF?", height=36)
        _saved = self.get_setting("folder_mietki_out_entry")
        if _saved: self.mietki_out_entry.insert(0, _saved)
        self.mietki_out_entry.grid(row=2, column=1, padx=5, pady=(8, 15), sticky="ew")
        ctk.CTkButton(card, text="Przeglądaj", image=self.icon_folder, command=lambda: self.select_dir(self.mietki_out_entry), width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444").grid(row=2, column=2, padx=15, pady=(8, 15))

        # --- POLA NAGŁÓWKA WSIE.DBF ---
        wsie_frame = ctk.CTkFrame(card, fg_color="#1E1E1E", border_width=1, border_color="#333333")
        wsie_frame.grid(row=3, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="ew")
        wsie_frame.grid_columnconfigure(1, weight=1)
        wsie_frame.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(wsie_frame, text="Dane nagłówka WSIE.DBF (stałe dla całego uruchomienia):",
                     font=font_label, text_color="#A0A0A0").grid(row=0, column=0, columnspan=4, padx=10, pady=(8, 6), sticky="w")

        def _wsie_row(r, c_label, c_entry, label, default, placeholder):
            ctk.CTkLabel(wsie_frame, text=label, font=font_btn, text_color="#E0E0E0").grid(row=r, column=c_label, padx=(10, 6), pady=4, sticky="e")
            e = ctk.CTkEntry(wsie_frame, height=30, placeholder_text=placeholder)
            if default:
                e.insert(0, default)
            e.grid(row=r, column=c_entry, padx=(0, 12), pady=4, sticky="ew")
            return e

        self.wsie_wojew_entry  = _wsie_row(1, 0, 1, "Województwo (kod):", "10",          "np. 10")
        self.wsie_powiat_entry = _wsie_row(1, 2, 3, "Powiat:",            "",            "np. WYSZKOWSKI")
        self.wsie_stan_entry   = _wsie_row(2, 0, 1, "Stan na:",           "01.01.2023",  "DD.MM.RRRR")
        self.wsie_obod_entry   = _wsie_row(2, 2, 3, "Obowiązuje od:",     "01.01.2023",  "DD.MM.RRRR")
        self.wsie_obdo_entry   = _wsie_row(3, 0, 1, "Obowiązuje do:",     "31.12.2032",  "DD.MM.RRRR")
        self.wsie_nrws_entry   = _wsie_row(3, 2, 3, "Nr wsi:",            "1",           "np. 1")
        self.wsie_rokz_entry   = _wsie_row(4, 0, 1, "Rok zal.:",          "19",          "np. 19")

        # --- Przywróć zapisane wartości pól WSIE.DBF ---
        _wsie_defaults = {
            "wsie_wojew": "10", "wsie_powiat": "", "wsie_stan": "01.01.2023",
            "wsie_obod": "01.01.2023", "wsie_obdo": "31.12.2032",
            "wsie_nrws": "1", "wsie_rokz": "19",
        }
        for _attr, _default in _wsie_defaults.items():
            _entry = getattr(self, f"{_attr}_entry", None)
            if _entry is not None:
                _saved = self.get_setting(f"wsie_{_attr}", _default)
                if _saved:
                    _entry.delete(0, "end")
                    _entry.insert(0, _saved)
        ctk.CTkLabel(wsie_frame, text="(NAZWA i GMINA = nazwa obrębu, wpisywane automatycznie)",
                     font=ctk.CTkFont(size=11), text_color="#777777").grid(row=4, column=2, columnspan=2, padx=(0, 12), pady=4, sticky="w")

        # --- CHECKBOX: USUWANIE WYDZIELEŃ BEZ WŁAŚCICIELI ---
        self.krzyz_usun_puste_jrej_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            card, text="Usuń wydzielenia bez właścicieli (pomiń wiersze bez J. rej. przy wpisywaniu krzyżówek)",
            variable=self.krzyz_usun_puste_jrej_var, font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#8B0000", hover_color="#A52A2A"
        ).grid(row=5, column=0, columnspan=3, padx=15, pady=(0, 10), sticky="w")

        # --- DWA PRZYCISKI ---
        btn_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        self.mietki_start_btn = ctk.CTkButton(
            btn_frame, text="Generuj same mietki", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#0067C0", hover_color="#005A9E", height=44, corner_radius=6,
            command=self.start_tworzenie_mietkow_pipeline
        )
        self.mietki_start_btn.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        self.mietki_krzyz_start_btn = ctk.CTkButton(
            btn_frame, text="Generuj mietki i wpisz krzyżówki", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#27ae60", hover_color="#219653", height=44, corner_radius=6,
            command=self.start_mietki_i_krzyzowki_pipeline
        )
        self.mietki_krzyz_start_btn.grid(row=0, column=1, padx=(8, 0), sticky="ew")

    # ==========================================
    # PIPELINE 1: SAME MIETKI
    # ==========================================
    def start_tworzenie_mietkow_pipeline(self, with_krzyzowki=False):
        baz_dir = self.mietki_bazowy_entry.get().strip() if hasattr(self, 'mietki_bazowy_entry') and self.mietki_bazowy_entry else ""
        rozl_dir = self.mietki_rozlicz_entry.get().strip() if hasattr(self, 'mietki_rozlicz_entry') and self.mietki_rozlicz_entry else ""
        out_dir = self.mietki_out_entry.get().strip() if self.mietki_out_entry else ""

        if not baz_dir or not Path(baz_dir).exists():
            messagebox.showwarning("Błąd", "Wybierz główny folder z plikami XLS (Ewidencja).")
            return
        if not rozl_dir or not Path(rozl_dir).exists():
            messagebox.showwarning("Błąd", "Wybierz folder z plikami rozliczonymi (XLSX).")
            return
        if not out_dir:
            messagebox.showwarning("Błąd", "Wybierz folder docelowy dla nowych obrębów.")
            return

        # --- WALIDACJA PRZED STARTEM ---
        baz_path = Path(baz_dir)
        rozl_path = Path(rozl_dir)
        baz_xls = sorted([p for p in baz_path.iterdir() if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"} and not p.name.startswith("~$")]) if baz_path.exists() else []
        rozl_xls = sorted([p for p in rozl_path.iterdir() if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"} and not p.name.startswith("~$")]) if rozl_path.exists() else []

        _problems = []
        if not baz_xls:
            _problems.append("• Folder bazowy nie zawiera żadnych plików .xls/.xlsx.")
        if not rozl_xls:
            _problems.append("• Folder rozliczeń nie zawiera żadnych plików .xls/.xlsx.")

        if baz_xls and rozl_xls:
            import re as _re
            # Buduj klucze z nazw bazowych (np. "BIAŁCZ" → "białcz")
            baz_map = {}
            for p in baz_xls:
                baz_map[_re.sub(r"[\s_]", "", p.stem.lower())] = p.stem
            # Z rozliczeń usuń sufiks _Rozliczone przed porównaniem
            rozl_map = {}
            for p in rozl_xls:
                _stem = p.stem
                # Usuń sufiks _Rozliczone (case-insensitive)
                _stem_lower = _stem.lower()
                if _stem_lower.endswith("_rozliczone"):
                    _stem = _stem[:-len("_Rozliczone")]
                rozl_map[_re.sub(r"[\s_]", "", _stem.lower())] = p.stem

            brak_w_rozl = sorted(set(baz_map.keys()) - set(rozl_map.keys()))
            brak_w_baz = sorted(set(rozl_map.keys()) - set(baz_map.keys()))

            if brak_w_rozl:
                nazwy = [baz_map[k] for k in brak_w_rozl]
                _problems.append(f"• {len(nazwy)} obrębów w folderze bazowym nie ma w folderze rozliczeń:")
                for nazwa in nazwy[:10]:
                    _problems.append(f"    — {nazwa}.xls → brak w rozliczeniach")
                if len(nazwy) > 10:
                    _problems.append(f"    ... i {len(nazwy) - 10} więcej")

            if brak_w_baz:
                nazwy = [rozl_map[k] for k in brak_w_baz]
                _problems.append(f"• {len(nazwy)} plików w rozliczeniach nie ma w folderze bazowym:")
                for nazwa in nazwy[:10]:
                    _problems.append(f"    — {nazwa}.xls → brak w bazowym")
                if len(nazwy) > 10:
                    _problems.append(f"    ... i {len(nazwy) - 10} więcej")

        out_path = Path(out_dir)
        if out_path.exists():
            existing = [p for p in out_path.iterdir() if p.is_dir()]
            if existing:
                _problems.append(f"• Folder wyjściowy zawiera już {len(existing)} podfolderów:")
                for p in existing[:10]:
                    _problems.append(f"    — {p.name}")
                if len(existing) > 10:
                    _problems.append(f"    ... i {len(existing) - 10} więcej")
                _problems.append("  Zostaną one nadpisane!")

        if _problems:
            msg = "Wykryto problemy przed startem:\n\n" + "\n".join(_problems)
            msg += "\n\nCzy chcesz kontynuować mimo to?"
            if not messagebox.askyesno("Walidacja — znaleziono problemy", msg):
                return

        baz_path = Path(baz_dir)
        xls_files = [
            f.stem for f in baz_path.iterdir()
            if f.is_file() and f.suffix.lower() in {'.xls', '.xlsx'} and not f.name.startswith("~$")
        ]

        if not xls_files:
            messagebox.showwarning("Błąd", "We wskazanym folderze XLS Ewidencji nie znaleziono żadnych plików, z których można by pobrać nazwy obrębów.")
            return

        names_list = sorted(list(set(xls_files)))

        base_dir = get_resource_path("pusty")
        if not Path(base_dir).exists() or not Path(base_dir).is_dir():
            messagebox.showerror("Błąd", f"Nie znaleziono wbudowanego folderu 'pusty' w plikach programu!\nŚcieżka: {base_dir}")
            return

        if self.running: return
        self.last_output_dir = Path(out_dir)
        self._disable_ui_for_process()
        self.set_progress(0)

        # Zapisz wartości pól WSIE.DBF w ustawieniach
        self.set_setting("wsie_wsie_wojew", self.wsie_wojew_entry.get().strip())
        self.set_setting("wsie_wsie_powiat", self.wsie_powiat_entry.get().strip())
        self.set_setting("wsie_wsie_stan", self.wsie_stan_entry.get().strip())
        self.set_setting("wsie_wsie_obod", self.wsie_obod_entry.get().strip())
        self.set_setting("wsie_wsie_obdo", self.wsie_obdo_entry.get().strip())
        self.set_setting("wsie_wsie_nrws", self.wsie_nrws_entry.get().strip())
        self.set_setting("wsie_wsie_rokz", self.wsie_rokz_entry.get().strip())
        self.set_setting("folder_mietki_bazowy_entry", baz_dir)
        self.set_setting("folder_mietki_rozlicz_entry", rozl_dir)
        self.set_setting("folder_mietki_out_entry", out_dir)

        wsie_meta = {
            'WOJEW': self.wsie_wojew_entry.get().strip(),
            'POWIAT': self.wsie_powiat_entry.get().strip(),
            'STAN_NA': self.wsie_stan_entry.get().strip(),
            'OBOW_OD': self.wsie_obod_entry.get().strip(),
            'OBOW_DO': self.wsie_obdo_entry.get().strip(),
            'NR_WSI': self.wsie_nrws_entry.get().strip() or "1",
            'ROK_ZAL': self.wsie_rokz_entry.get().strip(),
        }
        if not wsie_meta['POWIAT']:
            self.log("[UWAGA] Pole 'Powiat' w danych WSIE.DBF jest puste — uzupełnij je, jeśli MIETEK go wymaga.")
        threading.Thread(
            target=self.run_tworzenie_mietkow_thread,
            args=(base_dir, out_dir, names_list, baz_dir, rozl_dir, wsie_meta, with_krzyzowki),
            daemon=True
        ).start()

    # ==========================================
    # PIPELINE 2: MIETKI + KRZYŻÓWKI
    # ==========================================
    def start_mietki_i_krzyzowki_pipeline(self):
        self.start_tworzenie_mietkow_pipeline(with_krzyzowki=True)

    def read_dbf(self, filename):
        """Odczytuje plik dBase III zwracając (fields, records).
        fields  = lista (nazwa, typ, dlugosc, decimals)
        records = lista słowników {nazwa_pola: wartosc_strip}
        Round-trip z write_dbf jest bezstratny dla pól C i N."""
        import struct
        with open(filename, 'rb') as f:
            header = f.read(32)
            if len(header) < 32:
                raise Exception(f"Za krótki nagłówek DBF: {filename}")
            num_records = struct.unpack('<I', header[4:8])[0]
            header_length = struct.unpack('<H', header[8:10])[0]
            record_length = struct.unpack('<H', header[10:12])[0]
            fields = []
            while True:
                fld = f.read(32)
                if len(fld) < 32 or fld[0] == 0x0D:
                    break
                name = fld[0:11].split(b'\x00', 1)[0].decode('ascii', 'replace')
                typ = chr(fld[11])
                length = fld[16]
                decimals = fld[17]
                fields.append((name, typ, length, decimals))
            f.seek(header_length)
            records = []
            for i in range(num_records):
                rec_raw = f.read(record_length)
                if len(rec_raw) < record_length:
                    break
                rec = {}
                off = 1  # pomijamy bajt flagi usunięcia
                for (name, typ, length, decimals) in fields:
                    raw = rec_raw[off:off + length]
                    off += length
                    if typ in ('C', 'M', 'G'):
                        rec[name] = raw.decode('cp852', 'replace').strip()
                    elif typ in ('N', 'F'):
                        clean_val = raw.decode('ascii', 'replace').replace('\x00', '').strip()
                        rec[name] = clean_val
                    elif typ == 'D':
                        rec[name] = raw.decode('ascii', 'replace').strip()
                    elif typ == 'L':
                        rec[name] = chr(raw[0]) if raw else ''
                    else:
                        rec[name] = raw.decode('cp852', 'replace').strip()
                records.append(rec)
        return fields, records

    def write_dbf(self, filename, fields, records):
        import struct
        import datetime
        num_records = len(records)
        header_length = 32 + (len(fields) * 32) + 1
        record_length = 1 + sum(f[2] for f in fields)
        with open(filename, 'wb') as f:
            f.write(struct.pack('<B', 0x03))
            now = datetime.datetime.now()
            f.write(struct.pack('<3B', now.year - 1900, now.month, now.day))
            f.write(struct.pack('<I', num_records))
            f.write(struct.pack('<H', header_length))
            f.write(struct.pack('<H', record_length))
            f.write(b'\x00' * 20)
            for field in fields:
                name, typ, length, decimals = field
                name_bytes = name.encode('ascii')[:10].ljust(11, b'\x00')
                f.write(name_bytes)
                f.write(typ.encode('ascii'))
                f.write(b'\x00' * 4)
                f.write(struct.pack('<B', length))
                f.write(struct.pack('<B', decimals))
                f.write(b'\x00' * 14)
            f.write(struct.pack('<B', 0x0D))
            for rec in records:
                f.write(b' ')
                for field in fields:
                    name, typ, length, decimals = field
                    val = rec.get(name, "0") if typ == 'N' else rec.get(name, "")
                    if typ == 'C':
                        val_bytes = str(val).encode('cp852', errors='replace')[:length].ljust(length, b' ')
                        f.write(val_bytes)
                    elif typ == 'N':
                        val_str = str(val)[:length]
                        val_bytes = val_str.encode('ascii', errors='ignore').rjust(length, b' ')
                        f.write(val_bytes)
                    elif typ == 'D':
                        val_bytes = str(val).encode('ascii', errors='ignore')[:length].ljust(length, b' ')
                        f.write(val_bytes)
            f.write(struct.pack('<B', 0x1A))

    def parse_wlasciciel(self, text, j_rej):
        if pd.isna(text): return []
        text = str(text).strip()
        blocks = re.split(r'(?m)^(\d+/\d+)\s+\[.*?\]\s*', text)
        if len(blocks) == 1:
            text = "1/1 [własność] " + text
            blocks = re.split(r'(?m)^(\d+/\d+)\s+\[.*?\]\s*', text)

        results = []
        for i in range(1, len(blocks), 2):
            share = blocks[i].strip()
            if share == '1/1': share = ""
            rest = blocks[i + 1]
            lines = [line.strip() for line in rest.split('\n') if line.strip()]

            # --- ROZBIJANIE LINII ZE ŚREDNIKAMI ---
            # Jeśli linia ma średnik: część przed pierwszym średnikiem to nazwisko,
            # a cała reszta połączona przecinkami to PEŁNY adres tej osoby
            # (ulica + numer + kod pocztowy + miejscowość).
            # entries: [nazwisko, pełny_adres] — adres z jednej linii (po średniku
            # po nazwisku) to JEDEN pełny adres tej osoby: ulica + miejscowość + kod
            entries = []
            addr_pool = []   # adresy z linii czysto adresowych (bez nazwiska)
            for line in lines:
                if ';' in line:
                    parts = [p.strip().rstrip(';').strip() for p in line.split(';')]
                    parts = [p for p in parts if p]
                    if parts:
                        first = parts[0]
                        has_marker = bool(re.search(r'\[(OF|OP|PG)\]', first))
                        # Jeśli pierwsza część nie ma markera i wygląda na adres
                        # (ma cyfry, pasuje do wzorca ulicy) → cała linia to adres
                        looks_like_address = bool(
                            not has_marker and (
                                re.search(r'\d{2}-\d{3}', first) or
                                re.search(r'\d+\s*m\.?\s*\d+', first, re.IGNORECASE) or
                                re.search(r'\bm\.\s*\d+', first, re.IGNORECASE) or
                                re.search(r'^\D+\s+\d+[A-Za-z]?(?:\s+m\.?\s*\d+)?\s*$', first) is not None or
                                'ul.' in first.lower() or
                                'ulica' in first.lower()
                            )
                        )
                        # Dodatkowo: jeśli nie ma markera, nie wygląda na adres,
                        # ale składa się z jednego słowa lub kończy się myślnikiem
                        # i mamy już nazwiska → to nazwa miejscowości, nie nazwisko
                        looks_like_place = bool(
                            not has_marker and entries and (
                                len(first.split()) <= 1 or
                                bool(re.search(r'^\S+\s*-\s*$', first))
                            )
                        )
                        if looks_like_address or looks_like_place:
                            addr_pool.append(", ".join(parts))
                        else:
                            if has_marker:
                                clean = re.sub(r'\s*\[(OF|OP|PG)\]', '', first).strip()
                            else:
                                clean = first
                            entries.append([clean, ", ".join(parts[1:])])
                else:
                    # Bez średnika — klasyfikuj heurystycznie
                    has_marker = bool(re.search(r'\[(OF|OP|PG)\]', line))
                    is_address = bool(
                        re.search(r'\d{2}-\d{3}', line) or
                        'ul.' in line.lower() or
                        'miejsc.' in line.lower() or
                        re.search(r'\d+\s*m\.\s*\d+', line, re.IGNORECASE) or
                        re.search(r'\d+\s*m\s*\d+', line, re.IGNORECASE) or
                        re.search(r'\bm\.\s*\d+', line, re.IGNORECASE) or
                        re.search(r'^\D+\s+\d+[A-Za-z]?(?:\s+m\.?\s*\d+)?\s*$', line) is not None
                    )
                    if has_marker:
                        clean_name = re.sub(r'\s*\[(OF|OP|PG)\]', '', line).strip()
                        entries.append([clean_name, ""])
                    elif is_address:
                        addr_pool.append(line)
                    elif entries:
                        # Linia bez średnika, bez markera, bez cech adresu,
                        # ale mamy już nazwiska → to nazwa miejscowości (kontynuacja adresu)
                        addr_pool.append(line)
                    else:
                        entries.append([line, ""])

            # Wypełnij brakujące adresy:
            # 1) nazwisko bez własnego adresu dziedziczy ostatni widziany adres
            last = ""
            for e in entries:
                if e[1]:
                    last = e[1]
                else:
                    e[1] = last
            # 2) nazwiska z początku bloku (przed pierwszym adresem) — weź
            #    pierwszy dostępny adres (od osób albo z puli adresów)
            first_avail = next((e[1] for e in entries if e[1]), "") or (addr_pool[0] if addr_pool else "")
            for e in entries:
                if not e[1]:
                    e[1] = first_avail

            for j, (name, addr) in enumerate(entries):
                addr = self.napraw_powtorzenia_adresu(addr)

                try:
                    nrrej_val = int(float(j_rej))
                except:
                    nrrej_val = 0

                results.append({
                    'NRREJ': nrrej_val,
                    'NAZWISKO': str(name)[:30].strip(),
                    'IMIE': str(share)[:30].strip() if j == 0 else '',
                    'RODZICE': '',
                    'ADRES': str(addr)[:60].strip()
                })
        return results

    def _parse_dbf_date(self, s):
        """Zamienia datę z pola GUI (np. '1.01.2023', '01.01.2023', '2023-01-01')
        na format dBase 'YYYYMMDD'. Zwraca '' gdy pusto/niepoprawnie."""
        if not s:
            return ""
        nums = re.findall(r'\d+', str(s))
        if len(nums) == 3:
            d, m, y = nums[0], nums[1], nums[2]
            return f"{int(y):04d}{int(m):02d}{int(d):02d}"
        if len(nums) == 1 and len(nums[0]) == 8:
            return nums[0]
        return ""

    def build_wsie_record(self, name, meta):
        """Buduje jeden rekord WSIE.DBF dla obrębu 'name'."""
        return {
            'NAZWA':   str(name)[:40],
            'WOJEW':   str(meta.get('WOJEW', ''))[:30],
            'GMINA':   str(name)[:30],
            'STAN_NA': self._parse_dbf_date(meta.get('STAN_NA', '')),
            'OBOW_OD': self._parse_dbf_date(meta.get('OBOW_OD', '')),
            'OBOW_DO': self._parse_dbf_date(meta.get('OBOW_DO', '')),
            'NR_WSI':  str(meta.get('NR_WSI', '1')),
            'ROK_ZAL': str(meta.get('ROK_ZAL', ''))[:2],
            'POWIAT':  str(meta.get('POWIAT', ''))[:30],
        }

    def napraw_powtorzenia_adresu(self, addr):
        """Usuwa powtórzoną nazwę miejscowości w adresie."""
        if not addr:
            return addr
        s = addr.strip()
        m = re.match(r'^(\d{2}-\d{3})\s+(.+?)\s+\2(?:\s+(\d.*))?\s*$', s)
        if m:
            kod, miejsc, numer = m.group(1), m.group(2), m.group(3)
            return f"{kod} {miejsc} {numer}".strip() if numer else f"{kod} {miejsc}"
        return s

    def _find_001_dir(self, obr):
        """Zwraca istniejący podkatalog pasujący do *.001 (WOL.001 / KAM.001 / ...)
        w obrębie folderu obrębu, albo None jeśli takiego nie ma."""
        for d in obr.rglob("*.001"):
            if d.is_dir():
                return d
        return None

    def process_mietek_dbf(self, path_bazowy, path_rozl=None):
        try:
            self.log(f"  [DBF] Odczyt ewidencji XLS: {Path(path_bazowy).name}")

            tabela_xls, df_full = wczytaj_i_przetworz_wlascicieli(str(path_bazowy))

            if tabela_xls.empty:
                self.log("  [DBF] Ostrzeżenie: Plik XLS nie zawiera działek 'Ls'. Baza DBF będzie pusta.")
                return []

            self.log(f"  [DBF] Pomyślnie zlokalizowano {len(tabela_xls)} wydzieleń z lasem (Ls).")

            def clean_jrej(val):
                v = str(val).strip()
                if 'G' in v: v = v.split('G')[-1]
                v = re.sub(r'\.0$', '', v)
                try:
                    return str(int(float(v)))
                except:
                    return v

            tabela_xls['J. rej. clean'] = tabela_xls['J. rej.'].apply(clean_jrej)

            if path_rozl:
                df_rozl = pd.read_excel(str(path_rozl))
                rozl_jrej_col = next((c for c in df_rozl.columns if 'j. rej' in str(c).lower()), None)

                if rozl_jrej_col:
                    unique_j_rej = df_rozl[rozl_jrej_col].dropna().apply(clean_jrej).unique()
                    matched_rows = tabela_xls[tabela_xls['J. rej. clean'].isin(unique_j_rej)]
                    self.log(f"  [DBF] Po przefiltrowaniu rozliczonych zostało: {len(matched_rows)} wierszy.")
                else:
                    matched_rows = tabela_xls
            else:
                matched_rows = tabela_xls

            matched_rows = matched_rows.drop_duplicates(subset=['J. rej. clean'])
            self.log(f"  [DBF] Unikalnych rejestrów (właścicieli) gotowych do DBF: {len(matched_rows)}")

            dbf_records = []
            for _, row in matched_rows.iterrows():
                j_rej_val = row['J. rej. clean']
                wlasciciel_text = str(row.get('Właściciel', 'Brak danych')).replace('nan', 'Brak danych')

                recs = self.parse_wlasciciel(wlasciciel_text, j_rej_val)
                dbf_records.extend(recs)

            self.log(f"  [DBF] Sukces! Wygenerowano {len(dbf_records)} gotowych rekordów do MS-DOS.")
            return dbf_records

        except Exception as e:
            import traceback
            self.log(f"  [Błąd DBF] Wyjątek: {e}")
            self.log(traceback.format_exc())
            return []

    def run_tworzenie_mietkow_thread(self, base_dir_str, out_dir_str, names_list, baz_dir_str, rozl_dir_str=None,
                                     wsie_meta=None, with_krzyzowki=False):
        try:
            self.update_status("Generowanie struktury MIETEK...", "#0078D7")
            base_dir = Path(base_dir_str)
            out_dir = Path(out_dir_str)
            out_dir.mkdir(parents=True, exist_ok=True)

            baz_dir = Path(baz_dir_str) if baz_dir_str else None
            rozl_dir = Path(rozl_dir_str) if rozl_dir_str else None

            total = len(names_list)
            self.start_progress_tracking(total, "Kopiowanie folderów bazowych")

            stat_sukces = 0
            stat_bledy = []

            dbf_fields = [
                ('NRREJ', 'N', 5, 0), ('NAZWISKO', 'C', 30, 0),
                ('IMIE', 'C', 30, 0), ('RODZICE', 'C', 30, 0), ('ADRES', 'C', 60, 0),
                ('KOLEJNY', 'N', 3, 0), ('PREJ', 'N', 6, 0)
            ]

            for idx, name in enumerate(names_list, start=1):
                self.check_stop()
                self.progress_current_file = name
                self.log(f"[MIETKI] Tworzenie folderu dla: {name}")

                target_dir = out_dir / name
                try:
                    # Jeśli folder istnieje — pomiń, nie nadpisuj
                    if target_dir.exists():
                        self.log(f"  -> Folder '{name}' już istnieje — pominięto (nie nadpisano).")
                        self.set_progress(idx / total)
                        continue
                    shutil.copytree(base_dir, target_dir)

                    # WSTRZYKIWANIE BAZY WŁAŚCICIELI
                    if baz_dir:
                        path_baz = self.find_matching_file(baz_dir, name)
                        path_rozl = self.find_matching_file(rozl_dir, name) if rozl_dir else None

                        if path_baz:
                            dbf_records = self.process_mietek_dbf(path_baz, path_rozl)
                            if dbf_records:
                                w_dbfs = []
                                seen = set()
                                for p in (list(target_dir.rglob("W*.DBF")) + list(target_dir.rglob("W*.dbf")) +
                                          list(target_dir.rglob("w*.DBF")) + list(target_dir.rglob("w*.dbf"))):

                                    if p.stem.upper() == "WSIE":
                                        continue

                                    key = str(p).upper()
                                    if key not in seen:
                                        seen.add(key)
                                        w_dbfs.append(p)

                                if w_dbfs:
                                    target_dbf = w_dbfs[0]
                                else:
                                    sub = self._find_001_dir(target_dir)
                                    if sub is None:
                                        sub = target_dir / "WOL.001"
                                    sub.mkdir(parents=True, exist_ok=True)
                                    target_dbf = sub / "W0011019.DBF"

                                self.write_dbf(str(target_dbf), dbf_fields, dbf_records)
                                self.log(f"  -> Zapisano {len(dbf_records)} właścicieli do {target_dbf.name}")
                            else:
                                self.log(f"  -> Brak danych właścicieli do wpisania dla '{name}'.")
                        else:
                            self.log(f"  -> Ominięto wpisywanie właścicieli. Brak pliku XLS Ewidencji dla '{name}'.")

                    # ZAPIS WSIE.DBF (metadane obrębu, 1 rekord)
                    try:
                        wol_dir_wsie = self._find_001_dir(target_dir)
                        if wol_dir_wsie is None:
                            wol_dir_wsie = target_dir / "WOL.001"
                        wol_dir_wsie.mkdir(parents=True, exist_ok=True)

                        wsie_dbf = wol_dir_wsie / "WSIE.DBF"
                        wsie_record = self.build_wsie_record(name, wsie_meta or {})
                        self.write_dbf(str(wsie_dbf), WSIE_FIELDS, [wsie_record])
                        self.log(
                            f"  -> Zapisano WSIE.DBF (NAZWA={name}, GMINA={name}, POWIAT={(wsie_meta or {}).get('POWIAT', '')})")

                        # --- WALIDACJA DBF PO ZAPISIE ---
                        try:
                            _fields, _records = self.read_dbf(str(wsie_dbf))
                            if not _records:
                                self.log(f"  ⚠️ [WALIDACJA] WSIE.DBF dla '{name}' jest pusty — brak rekordów!")
                            else:
                                _rec = _records[0]
                                _dbf_issues = []
                                if not _rec.get('NAZWA', '').strip():
                                    _dbf_issues.append("NAZWA jest pusta")
                                if not _rec.get('GMINA', '').strip():
                                    _dbf_issues.append("GMINA jest pusta")
                                if not _rec.get('WOJEW', '').strip():
                                    _dbf_issues.append("WOJEW jest puste")
                                if not _rec.get('POWIAT', '').strip():
                                    _dbf_issues.append("POWIAT jest puste")
                                try:
                                    _nr_wsi = int(_rec.get('NR_WSI', '0'))
                                    if _nr_wsi <= 0:
                                        _dbf_issues.append("NR_WSI wynosi 0")
                                except (ValueError, TypeError):
                                    _dbf_issues.append("NR_WSI nie jest liczbą")
                                if _dbf_issues:
                                    self.log(f"  ⚠️ [WALIDACJA] WSIE.DBF dla '{name}': {', '.join(_dbf_issues)}")
                                else:
                                    self.log(f"  ✅ [WALIDACJA] WSIE.DBF dla '{name}' — OK ({len(_records)} rekord, pola wypełnione)")
                        except Exception as _ve:
                            self.log(f"  ⚠️ [WALIDACJA] Nie można odczytać WSIE.DBF dla '{name}': {_ve}")
                    except Exception as e:
                        self.log(f"  -> [Ostrzeżenie] Błąd zapisu WSIE.DBF dla '{name}': {e}")

                    stat_sukces += 1
                except Exception as e:
                    self.log(f"  ❌ Błąd kopiowania dla '{name}': {e}")
                    stat_bledy.append(name)

                self.set_progress(idx / total)

            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            self.log(f"\n✅ Zakończono generowanie folderów. Utworzono: {stat_sukces}/{total}")

            # --- JEŚLI ZAZNACZONO KRZYŻÓWKI, URUCHOM AUTOMATYCZNIE ---
            if with_krzyzowki and rozl_dir and out_dir:
                self.log("\n" + "="*50)
                self.log("[KRZYŻÓWKI] Automatyczne wpisywanie krzyżówek...")
                self.run_krzyzowki_thread(str(rozl_dir), str(out_dir), self.krzyz_usun_puste_jrej_var.get())
            else:
                self.after(0,
                           lambda: messagebox.showinfo("Sukces", f"Wygenerowano pomyślnie {stat_sukces} folderów MIETEK."))

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
    # WPISYWANIE KRZYŻÓWEK (D*.DBF)
    # ==========================================
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

            dbf_fields = [
                ('NRREJ', 'N', 5, 0),
                ('NR_DZIAL', 'C', 9, 0),
                ('POW', 'N', 9, 4),
                ('POW_L_ZAL', 'N', 9, 4),
                ('POW_L_NZAL', 'N', 8, 4),
                ('POW_N_ZAL', 'N', 9, 4),
                ('POW_INNE', 'N', 8, 4),
                ('ODDZIAL', 'C', 7, 0),
                ('PODODDZ', 'C', 3, 0),
                ('ZM', 'C', 1, 0),
                ('PREJ', 'N', 6, 0),
            ]

            stat_ok = 0
            stat_brak_folderu = 0
            stat_puste = 0

            for idx, xls_path in enumerate(xls_files, start=1):
                self.check_stop()
                self.progress_current_file = xls_path.name

                v_name = re.sub(r'(?i)_?rozliczone.*$', '', xls_path.stem).strip()
                if not v_name:
                    v_name = xls_path.stem
                v_norm = re.sub(r'[\s_\-]', '', v_name.lower())

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

                        if usun_puste_jrej and nrrej_val == 0:
                            continue

                        nr_dz = str(row.get('nr_dz', '')).strip()
                        litery = str(row.get('litery', ''))
                        # "X" w litery → ODDZIAL (kolumna H), nie PODODDZ (kolumna I)
                        if litery.strip().upper() == 'X':
                            oddzial = 'X'
                            pododdz = ''
                        else:
                            oddzial = "".join(ch for ch in litery if ch.isdigit())[:7]
                            pododdz = "".join(ch for ch in litery if ch.isalpha())[:3]
                        pow_val = row['__POW']
                        records.append({
                            'NRREJ': nrrej_val,
                            'NR_DZIAL': nr_dz[:9],
                            'POW': f"{float(pow_val):.4f}",
                            'POW_L_ZAL': f"{float(pow_val):.4f}",
                            'ODDZIAL': oddzial,
                            'PODODDZ': pododdz,
                        })

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
