"""
Forestly — Mixin: TabMietekPlus10Mixin
Zakładka: Mietki +10 lat — przesuwanie wieku w opisach taksacyjnych.

Zastępuje makro VBA „Dodaj10DoZakresowKolE": każde /65-75/80 (lub /65-75/80l)
w polu OP_TAX plików O*.DBF dostaje +N lat do wszystkich trzech liczb
(np. /65-75/80l  ->  /75-85/90l). Domyślnie przetwarzane jest też pole OP_TAX1,
jeśli w nim również występują zakresy wieku.

Funkcje read_dbf/write_dbf to kopie z tab_tworzenie_mietkow — zakładka jest
samodzielna (nie zależy od innych mixinów) i w pełni testowalna bez GUI.
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import re
import struct
import datetime
import shutil

# Wzorzec wieku z makra VBA: /65-75/80 lub /65-75/80l (l/L opcjonalne)
WZORZEC_WIEKU = re.compile(r'/(\d+)-(\d+)/(\d+)([lL]?)')


def przesun_wiek_w_tekscie(tekst, lata):
    """Dodaje `lata` do każdego zapisu wieku /x-y/z (i /x-y/zl) w tekście."""
    if not tekst:
        return tekst

    def _zamien(m):
        n1 = int(m.group(1)) + lata
        n2 = int(m.group(2)) + lata
        n3 = int(m.group(3)) + lata
        return "/%d-%d/%d%s" % (n1, n2, n3, m.group(4))

    return WZORZEC_WIEKU.sub(_zamien, tekst)


def read_dbf(filename):
    """Odczytuje plik dBase III zwracając (fields, records).
    (kopia z tab_tworzenie_mietkow — round-trip z write_dbf bezstratny dla C i N)"""
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


def write_dbf(filename, fields, records):
    """Zapisuje listę rekordów do dBase III. (kopia z tab_tworzenie_mietkow)"""
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


def przesun_wiek_w_pliku(path, lata, zapisz=False, backup=True, pola=("OP_TAX", "OP_TAX1")):
    """Przesuwa wiek w jednym pliku O*.DBF.

    Zwraca słownik: {rekordow, zmienionych, przyklady: [(przed, po), ...]}.
    Przy zapisz=True robi kopię .BAK (jeśli backup=True) i zapisuje plik.
    """
    path = Path(path)
    fields, records = read_dbf(path)
    nazwy_pol = {f[0] for f in fields}
    cele = [p for p in pola if p in nazwy_pol]
    dlugosci = {f[0]: f[2] for f in fields}

    zmienione = 0
    przyklady = []
    for rec in records:
        zmieniono_ten = False
        for p in cele:
            stara = rec.get(p, "")
            if not stara:
                continue
            nowa = przesun_wiek_w_tekscie(stara, lata)
            if nowa != stara:
                if len(nowa) > dlugosci[p]:
                    raise Exception(
                        f"Po przesunięciu wiek nie mieści się w polu {p} "
                        f"({len(nowa)} > {dlugosci[p]} znaków) w pliku {path.name}. "
                        f"Nie zapisano."
                    )
                rec[p] = nowa
                zmieniono_ten = True
                if len(przyklady) < 3:
                    # pokaż fragment wokół pierwszego wzorca
                    m = WZORZEC_WIEKU.search(stara)
                    if m:
                        start = max(0, m.start() - 12)
                        przyklady.append((stara[start:start + 40], nowa[start:start + 40]))
        if zmieniono_ten:
            zmienione += 1

    if zapisz and zmienione:
        if backup:
            shutil.copy2(path, path.parent / (path.name + ".bak"))
        write_dbf(path, fields, records)

    return {"rekordow": len(records), "zmienionych": zmienione, "przyklady": przyklady}


class TabMietekPlus10Mixin:
    """Mixin dla ModernApp — zakładka przesuwania wieku w mietkach (O*.DBF)."""

    def setup_mietek_plus10_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)

        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)
        font_small = ctk.CTkFont(family="Segoe UI", size=12)

        card = ctk.CTkFrame(scroll_frame, fg_color="#252526", corner_radius=8,
                            border_width=1, border_color="#333333")
        card.grid(row=0, column=0, padx=20, pady=(15, 15), sticky="new")
        card.grid_columnconfigure(1, weight=1)

        # 1. Folder z mietkami
        ctk.CTkLabel(card, text="1. Folder z mietkami (pliki O*.DBF):",
                     font=font_label, text_color="#E0E0E0").grid(
            row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.plus10_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Folder, w którym leżą O0011019.DBF itd...", height=36)
        _saved = self.get_setting("folder_plus10_entry")
        if _saved:
            self.plus10_folder_entry.insert(0, _saved)
        self.plus10_folder_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(card, text="Przeglądaj", image=self.icon_folder,
                      command=lambda: self.select_dir(self.plus10_folder_entry),
                      width=110, height=36, font=font_btn,
                      fg_color="#333333", hover_color="#444444").grid(
            row=0, column=2, padx=15, pady=(15, 8))

        # 2. Przesunięcie wieku
        ctk.CTkLabel(card, text="2. Przesunięcie wieku (lata):",
                     font=font_label, text_color="#E0E0E0").grid(
            row=1, column=0, padx=15, pady=8, sticky="w")
        self.plus10_lata_entry = ctk.CTkEntry(card, width=90, height=36, justify="center")
        self.plus10_lata_entry.insert(0, "10")
        self.plus10_lata_entry.grid(row=1, column=1, padx=5, pady=8, sticky="w")

        # 3. Opcje
        self.plus10_backup_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(card, text="Kopia zapasowa przed zapisem (plik .BAK)",
                        font=font_small, variable=self.plus10_backup_var,
                        fg_color="#0067C0", hover_color="#005A9E").grid(
            row=2, column=0, padx=15, pady=(2, 2), sticky="w")
        self.plus10_tax1_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(card, text="Przetwarzaj również pole OP_TAX1",
                        font=font_small, variable=self.plus10_tax1_var,
                        fg_color="#0067C0", hover_color="#005A9E").grid(
            row=3, column=0, padx=15, pady=(2, 8), sticky="w")

        # 4. Przyciski
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.grid(row=4, column=0, columnspan=3, padx=15, pady=(4, 15), sticky="w")
        ctk.CTkButton(btn_frame, text="🔍 Podgląd zmian",
                      command=lambda: self._plus10_start(zapisz=False),
                      height=38, font=font_btn,
                      fg_color="#333333", hover_color="#444444").pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="✅ Zastosuj przesunięcie wieku",
                      command=lambda: self._plus10_start(zapisz=True),
                      height=38, font=font_btn,
                      fg_color="#0067C0", hover_color="#005A9E").pack(side="left")

        # Opis
        info = ctk.CTkLabel(scroll_frame,
                            text="Zamiana dawnego makra VBA: każde /65-75/80 (lub /65-75/80l)\n"
                                 "w opisach taksacyjnych (OP_TAX) dostaje +N lat do wszystkich trzech liczb,\n"
                                 "np. /65-75/80l → /75-85/90l. Przetwarzane są wszystkie pliki O*.DBF w folderze.",
                            font=font_small, text_color="#9E9E9E", justify="left")
        info.grid(row=1, column=0, padx=25, pady=(0, 20), sticky="w")

    # ---------------------------------------------------------------

    def _plus10_zbierz_dane(self):
        """Waliduje wejście i zwraca (folder_path, lata, pliki, pola) albo None."""
        folder = self.plus10_folder_entry.get().strip()
        if not folder or not Path(folder).exists():
            messagebox.showwarning("Błąd", "Wybierz folder z mietkami (pliki O*.DBF).")
            return None
        try:
            lata = int(self.plus10_lata_entry.get().strip())
            if lata < 1 or lata > 100:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Błąd", "Przesunięcie wieku musi być liczbą od 1 do 100.")
            return None

        folder_path = Path(folder)
        pliki = sorted([p for p in folder_path.iterdir()
                        if p.is_file() and p.name.upper().startswith("O")
                        and p.suffix.upper() == ".DBF"])
        if not pliki:
            messagebox.showwarning("Błąd",
                                   "W tym folderze nie ma żadnych plików O*.DBF.\n"
                                   "(szukam plików zaczynających się od litery O, np. O0011019.DBF)")
            return None

        pola = ["OP_TAX"] + (["OP_TAX1"] if self.plus10_tax1_var.get() else [])
        return folder_path, lata, pliki, pola

    def _plus10_start(self, zapisz):
        dane = self._plus10_zbierz_dane()
        if not dane:
            return
        folder_path, lata, pliki, pola = dane

        # Zapamiętaj folder i opcje
        self.set_setting("folder_plus10_entry", str(folder_path))
        self.set_setting("plus10_lata", str(lata))

        # --- Przebieg próbny (odczyt + podgląd) — pliki są małe, robimy w wątku GUI ---
        wyniki = {}
        total_zmienionych = 0
        blad = None
        try:
            for p in pliki:
                w = przesun_wiek_w_pliku(p, lata, zapisz=False, pola=pola)
                wyniki[p] = w
                total_zmienionych += w["zmienionych"]
        except Exception as e:
            blad = e

        if blad:
            messagebox.showerror("Błąd", str(blad))
            return

        if total_zmienionych == 0:
            messagebox.showinfo("Brak zmian",
                                "W żadnym z opisów nie znaleziono zapisów wieku /x-y/z.\n"
                                "Nic nie zmieniono.")
            return

        # Log podglądu
        tryba = "PODGLĄD" if not zapisz else "ZAPIS"
        self.log(f"[MIETKI +{lata} LAT] {tryba} — folder: {folder_path.name}")
        for p, w in wyniki.items():
            self.log(f"  {p.name}: {w['zmienionych']}/{w['rekordow']} rekordów ze zmianą")
            for przed, po in w["przyklady"]:
                self.log(f"     {przed.strip()}  →  {po.strip()}")

        if not zapisz:
            self.log(f"[MIETKI +{lata} LAT] To był tylko podgląd — pliki niezmienione.")
            return

        # --- Potwierdzenie ---
        if not messagebox.askyesno(
                "Potwierdź zmiany",
                f"Przesunięcie wieku o +{lata} lat zmieni opisy w {total_zmienionych} rekordach "
                f"w {len(pliki)} plikach(ach).\n\n"
                f"Kopia zapasowa .BAK: {'TAK' if self.plus10_backup_var.get() else 'NIE'}.\n\n"
                f"Kontynuować?"):
            self.log(f"[MIETKI +{lata} LAT] Anulowano przez użytkownika.")
            return

        # --- Zapis w tle ---
        def _zapisz():
            try:
                zrobione = 0
                for p in pliki:
                    w = przesun_wiek_w_pliku(
                        p, lata, zapisz=True,
                        backup=self.plus10_backup_var.get(), pola=pola)
                    if w["zmienionych"]:
                        zrobione += 1
                        self.log(f"  ✔ {p.name}: przesunięto wiek w {w['zmienionych']} rekordach"
                                 + (f" (kopia: {p.name}.bak)" if self.plus10_backup_var.get() else ""))
                self.log(f"[MIETKI +{lata} LAT] Gotowe — zmieniono {zrobione} plików.")
            except Exception as e:
                self.log(f"[MIETKI +{lata} LAT] BŁĄD: {e}")

        threading.Thread(target=_zapisz, daemon=True).start()
