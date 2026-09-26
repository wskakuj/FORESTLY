"""
Forestly — Mixin: TabMietekPlus10Mixin
Zakładka: Mietki +10 lat — przesuwanie wieku w opisach taksacyjnych.

O*.DBF: każde /65-75/80 (lub /65-75/80l) w polu OP_TAX dostaje +N lat do
wszystkich trzech liczb (np. /65-75/80l -> /75-85/90l). Zastępuje makro VBA.

R*.DBF: pole WIEK dostaje +N lat, a klasa wieku (KL_WIEK) jest przesuwana
o pół klasy (a->b, b->następna klasa; powyżej V bez podziału). Rekordy z
WIEK=0 (halizny/zrąb) są pomijane.

Wskazujemy FOLDER MIETKA (np. z podfolderem WOL.001) — program sam znajduje
pliki O*.DBF i R*.DBF w całym drzewie (w podfolderach *.001).

Funkcje read_dbf/write_dbf to kopie z tab_tworzenie_mietkow — zakładka jest
samodzielna i w pełni testowalna bez GUI.
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import os
import re
import struct
import datetime
import shutil

# Wzorzec wieku z makra VBA: /65-75/80 lub /65-75/80l (l/L opcjonalne)
WZORZEC_WIEKU = re.compile(r'/(\d+)-(\d+)/(\d+)([lL]?)')

_RZYMSKIE = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]


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


def nastepna_klasa_wieku(kl):
    """Przesuwa klasę wieku o pół klasy (10 lat): IIa->IIb, IIb->IIIa, Vb->VI.

    Klasy powyżej V nie mają połówek (VI = 101-120) — zostają bez zmian.
    Nieznane/puste wartości zostają bez zmian.
    """
    if not kl:
        return kl
    m = re.fullmatch(r'(I{1,3}|IV|VI{0,3}|IX|X)(a|b)?', kl.strip())
    if not m:
        return kl
    rzymska, pol = m.group(1), m.group(2)
    if pol == "a":
        return rzymska + "b"
    if pol == "b":
        try:
            nxt = _RZYMSKIE[_RZYMSKIE.index(rzymska) + 1]
        except (ValueError, IndexError):
            return kl
        # klasy VI i wyższe nie mają połówek
        return nxt if nxt in ("VI", "VII", "VIII", "IX", "X") else nxt + "a"
    return kl  # np. samo "VI" — bez zmian


def znajdz_pliki_dbf(folder_mietka, litera):
    """Szuka plików <litera>*.DBF w drzewie folderu mietka (np. WOL.001/O0011019.DBF).

    Zwraca posortowaną listę Path — działa też, gdy wskazany zostanie
    bezpośrednio folder *.001 albo folder z plikami DBF na wierzchu.
    Pomija kopie zapasowe (.BAK).
    """
    root = Path(folder_mietka)
    trafienia = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        for fn in sorted(filenames):
            if fn[:1].upper() == litera.upper() and fn.upper().endswith(".DBF"):
                p = Path(dirpath) / fn
                if p.suffix.upper() == ".BAK":
                    continue
                trafienia.append(p)
    return sorted(trafienia)


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
                elif typ == 'L':
                    # pole logiczne (np. PRZES w R*.DBF): T/F/Y/N/?/spacja — 1 bajt
                    v = (str(val)[:1] if str(val) else " ") or " "
                    f.write(v.encode('ascii', errors='replace'))
                else:
                    # dowolny inny typ — tekst wypelniony spacjami (bezpieczny fallback)
                    val_bytes = str(val).encode('cp852', errors='replace')[:length].ljust(length, b' ')
                    f.write(val_bytes)
        f.write(struct.pack('<B', 0x1A))


def _zapisz_z_backupem(path, fields, records, backup):
    if backup:
        shutil.copy2(path, path.parent / (path.name + ".bak"))
    write_dbf(path, fields, records)


def przesun_wiek_w_pliku(path, lata, zapisz=False, backup=True, pola=("OP_TAX", "OP_TAX1")):
    """Przesuwa wiek w pliku O*.DBF (wzorce /x-y/z w OP_TAX).

    Zwraca słownik: {rekordow, zmienionych, przyklady: [(przed, po), ...]}.
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
                    m = WZORZEC_WIEKU.search(stara)
                    if m:
                        start = max(0, m.start() - 12)
                        przyklady.append((stara[start:start + 40], nowa[start:start + 40]))
        if zmieniono_ten:
            zmienione += 1

    if zapisz and zmienione:
        _zapisz_z_backupem(path, fields, records, backup)

    return {"rekordow": len(records), "zmienionych": zmienione, "przyklady": przyklady}


def _dodaj_do_pola_liczbowego(rec, pole, dodaj, dlugosc, nazwa_pliku):
    """Dodaje `dodaj` do wartości pola N (tekstowo). Zwraca (stara, nowa) lub None.

    Pomija puste i zerowe wartości (np. WYS=0 przy haliznach).
    """
    stara = rec.get(pole, "").strip()
    if not stara or not stara.isdigit() or int(stara) == 0 or dodaj == 0:
        return None
    nowa = str(int(stara) + dodaj)
    if len(nowa) > dlugosc:
        raise Exception(
            f"Nowa wartość ({nowa}) nie mieści się w polu {pole} "
            f"({dlugosc} znaków) w pliku {nazwa_pliku}. Nie zapisano."
        )
    rec[pole] = nowa
    return (stara, nowa)


def przesun_wiek_w_pliku_r(path, lata, zapisz=False, backup=True, wys=0, piers=0):
    """Przesuwa wiek w pliku R*.DBF: WIEK + klasa KL_WIEK, oraz WYS i PIERS.

    Rekordy z WIEK=0 lub pustym są pomijane (halizny/zrąb).
    Zwraca słownik jak przesun_wiek_w_pliku.
    """
    path = Path(path)
    fields, records = read_dbf(path)
    nazwy_pol = {f[0] for f in fields}
    dlugosci = {f[0]: f[2] for f in fields}
    if "WIEK" not in nazwy_pol:
        raise Exception(f"Plik {path.name} nie ma pola WIEK — to nie wygląda na R*.DBF.")

    zmienione = 0
    przyklady = []
    for rec in records:
        stary_wiek = rec.get("WIEK", "").strip()
        if not stary_wiek or not stary_wiek.isdigit() or int(stary_wiek) == 0:
            continue
        nowy = str(int(stary_wiek) + lata)
        if len(nowy) > dlugosci["WIEK"]:
            raise Exception(
                f"Nowy wiek ({nowy}) nie mieści się w polu WIEK "
                f"({dlugosci['WIEK']} znaków) w pliku {path.name}. Nie zapisano."
            )
        rec["WIEK"] = nowy
        stara_kl = rec.get("KL_WIEK", "").strip()
        nowa_kl = nastepna_klasa_wieku(stara_kl) if stara_kl else stara_kl
        if nowa_kl and len(nowa_kl) > dlugosci.get("KL_WIEK", 4):
            nowa_kl = stara_kl  # nie zmieści się — zostaw starą
        rec["KL_WIEK"] = nowa_kl
        opis_przed = f"{rec.get('ODDZIAL', '?')}: w{stary_wiek} {stara_kl}"
        opis_po = f"{rec.get('ODDZIAL', '?')}: w{nowy} {nowa_kl}"
        if "WYS" in nazwy_pol:
            z = _dodaj_do_pola_liczbowego(rec, "WYS", wys, dlugosci["WYS"], path.name)
            if z:
                opis_przed += f", h{z[0]}"
                opis_po += f", h{z[1]}"
        if "PIERS" in nazwy_pol:
            z = _dodaj_do_pola_liczbowego(rec, "PIERS", piers, dlugosci["PIERS"], path.name)
            if z:
                opis_przed += f", d{z[0]}"
                opis_po += f", d{z[1]}"
        zmienione += 1
        if len(przyklady) < 3:
            przyklady.append((opis_przed, opis_po))

    if zapisz and zmienione:
        _zapisz_z_backupem(path, fields, records, backup)

    return {"rekordow": len(records), "zmienionych": zmienione, "przyklady": przyklady}


class TabMietekPlus10Mixin:
    """Mixin dla WebBackend — zakładka przesuwania wieku w mietkach (O + R)."""

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

        # 1. Folder z mietkiem
        ctk.CTkLabel(card, text="1. Folder z mietkiem (np. z podfolderem WOL.001):",
                     font=font_label, text_color="#E0E0E0").grid(
            row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.plus10_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Program sam znajdzie pliki O*.DBF i R*.DBF w podfolderach .001...",
            height=36)
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
        _saved_l = self.get_setting("plus10_lata")
        self.plus10_lata_entry.insert(0, _saved_l if _saved_l else "10")
        self.plus10_lata_entry.grid(row=1, column=1, padx=5, pady=8, sticky="w")

        # 3. Wysokość i pierśnica (R*.DBF)
        ctk.CTkLabel(card, text="3. Wysokość (WYS, +):",
                     font=font_label, text_color="#E0E0E0").grid(
            row=2, column=0, padx=15, pady=8, sticky="w")
        self.plus10_wys_entry = ctk.CTkEntry(card, width=90, height=36, justify="center")
        self.plus10_wys_entry.insert(0, self.get_setting("plus10_wys") or "1")
        self.plus10_wys_entry.grid(row=2, column=1, padx=5, pady=8, sticky="w")
        ctk.CTkLabel(card, text="4. Pierśnica (PIERS, +):",
                     font=font_label, text_color="#E0E0E0").grid(
            row=3, column=0, padx=15, pady=8, sticky="w")
        self.plus10_piers_entry = ctk.CTkEntry(card, width=90, height=36, justify="center")
        self.plus10_piers_entry.insert(0, self.get_setting("plus10_piers") or "2")
        self.plus10_piers_entry.grid(row=3, column=1, padx=5, pady=8, sticky="w")

        # (kopia .BAK, pola OP_TAX/OP_TAX1 i przetwarzanie R*.DBF — zawsze włączone)

        # 5. Przyciski
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
                            text="O*.DBF: każde /65-75/80 (lub /65-75/80l) w opisach taksacyjnych (OP_TAX)\n"
                                 "dostaje +N lat do wszystkich trzech liczb, np. /65-75/80l → /75-85/90l.\n"
                                 "R*.DBF: pole WIEK +N lat, klasa wieku (KL_WIEK) przesuwana o pół klasy\n"
                                 "(IIa→IIb, IIb→IIIa, Vb→VI). Rekordy z wiekiem 0 (halizny) pomijane.\n"
                                 "Zawsze włączone: kopia .BAK przed zapisem, pola OP_TAX i OP_TAX1.\n"
                                 "Wskaż folder mietka — program sam znajdzie pliki w podfolderach (np. WOL.001).",
                            font=font_small, text_color="#9E9E9E", justify="left")
        info.grid(row=1, column=0, padx=25, pady=(0, 20), sticky="w")

    # ---------------------------------------------------------------

    def _plus10_zbierz_dane(self):
        """Waliduje wejście i zwraca (folder_path, lata, pliki_o, pliki_r, pola) albo None."""
        folder = self.plus10_folder_entry.get().strip()
        if not folder or not Path(folder).exists():
            messagebox.showwarning("Błąd", "Wybierz folder z mietkiem (np. folder z WOL.001 w środku).")
            return None
        try:
            lata = int(self.plus10_lata_entry.get().strip())
            if lata < 1 or lata > 100:
                raise ValueError
            wys = int((self.plus10_wys_entry.get().strip() or "0"))
            piers = int((self.plus10_piers_entry.get().strip() or "0"))
            if wys < 0 or wys > 99 or piers < 0 or piers > 99:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Błąd",
                "Przesunięcie wieku musi być liczbą od 1 do 100,\n"
                "a wysokość i pierśnica od 0 do 99 (0 = bez zmian).")
            return None

        folder_path = Path(folder)
        pliki_o = znajdz_pliki_dbf(folder_path, "O")
        pliki_r = znajdz_pliki_dbf(folder_path, "R")
        if not pliki_o and not pliki_r:
            messagebox.showwarning(
                "Błąd",
                "W wybranym folderze (i jego podfolderach, np. WOL.001)\n"
                "nie znaleziono plików O*.DBF ani R*.DBF.")
            return None

        pola = ["OP_TAX", "OP_TAX1"]
        return folder_path, lata, pliki_o, pliki_r, pola, wys, piers

    def _plus10_koniec(self, tekst, kolor="#108C4C", folder_wynikow=None):
        """Standardowe domknięcie zadania: status, folder wyników, odblokowanie przycisków.

        Dzięki temu po zadaniu pojawia się dźwięk, powiadomienie „Zakończono”
        i przycisk „Otwórz folder wyników” — tak jak w pozostałych zakładkach.
        """
        try:
            self.update_status(tekst, kolor)
        except Exception:
            pass
        if folder_wynikow is not None and hasattr(self, "_zapamietaj_folder_wynikow"):
            self._zapamietaj_folder_wynikow(folder_wynikow)
            if hasattr(self, "last_output_dir"):
                self.last_output_dir = Path(folder_wynikow)
        if hasattr(self, "restore_all_buttons"):
            self.restore_all_buttons()

    def _plus10_start(self, zapisz):
        # sygnalizacja startu (running=True) — dzięki temu frontend widzi
        # przejście na koniec i pokazuje dźwięk + „Zakończono” + „Otwórz folder”
        if hasattr(self, "disable_all_buttons"):
            self.disable_all_buttons()
        dane = self._plus10_zbierz_dane()
        if not dane:
            self._plus10_koniec("Przerwano — brak danych.", "#D83B01")
            return
        folder_path, lata, pliki_o, pliki_r, pola, wys, piers = dane

        # Zapamiętaj folder i opcje
        self.set_setting("folder_plus10_entry", str(folder_path))
        self.set_setting("plus10_lata", str(lata))
        self.set_setting("plus10_wys", str(wys))
        self.set_setting("plus10_piers", str(piers))

        # --- Przebieg próbny (odczyt + podgląd) ---
        wyniki = {}
        total_zmienionych = 0
        blad = None
        try:
            for p in pliki_o:
                w = przesun_wiek_w_pliku(p, lata, zapisz=False, pola=pola)
                wyniki[p] = ("O", w)
                total_zmienionych += w["zmienionych"]
            for p in pliki_r:
                w = przesun_wiek_w_pliku_r(p, lata, zapisz=False, wys=wys, piers=piers)
                wyniki[p] = ("R", w)
                total_zmienionych += w["zmienionych"]
        except Exception as e:
            blad = e

        if blad:
            messagebox.showerror("Błąd", str(blad))
            self._plus10_koniec(f"Błąd: {blad}", "#D83B01")
            return

        if total_zmienionych == 0:
            messagebox.showinfo("Brak zmian",
                                "W żadnym z plików nie znaleziono nic do przesunięcia.\n"
                                "Nic nie zmieniono.")
            self._plus10_koniec("Gotowe — brak zmian do przesunięcia.")
            return

        # Log podglądu
        tryba = "PODGLĄD" if not zapisz else "ZAPIS"
        self.log(f"[MIETKI +{lata} LAT] {tryba} — folder: {folder_path.name} "
                 f"({len(pliki_o)}× O*.DBF, {len(pliki_r)}× R*.DBF)")
        for p, (rodzaj, w) in wyniki.items():
            wzgledna = p.relative_to(folder_path)
            self.log(f"  [{rodzaj}] {wzgledna}: {w['zmienionych']}/{w['rekordow']} rekordów ze zmianą")
            for przed, po in w["przyklady"][:2]:
                self.log(f"     {przed}  →  {po}")

        if not zapisz:
            self.log(f"[MIETKI +{lata} LAT] To był tylko podgląd — pliki niezmienione.")
            self._plus10_koniec(f"Gotowe — podgląd: {total_zmienionych} rekordów do przesunięcia.")
            return

        # --- Potwierdzenie ---
        if not messagebox.askyesno(
                "Potwierdź zmiany",
                f"Przesunięcie wieku o +{lata} lat zmieni {total_zmienionych} rekordów "
                f"w {len(pliki_o) + len(pliki_r)} plikach(ach).\n\n"
                f"(O*.DBF: opisy — /x-y/z; R*.DBF: WIEK, klasa wieku"
                + (f", WYS +{wys}" if wys else "")
                + (f", PIERS +{piers}" if piers else "") + ")\n\n"
                f"Przed zapisem każdorazowo powstaje kopia zapasowa .BAK.\n\n"
                f"Kontynuować?"):
            self.log(f"[MIETKI +{lata} LAT] Anulowano przez użytkownika.")
            self._plus10_koniec("Anulowano.", "#D83B01")
            return

        # --- Zapis ---
        try:
            zrobione = 0
            for p, (rodzaj, w) in wyniki.items():
                if rodzaj == "O":
                    w = przesun_wiek_w_pliku(p, lata, zapisz=True, backup=True, pola=pola)
                else:
                    w = przesun_wiek_w_pliku_r(p, lata, zapisz=True, backup=True, wys=wys, piers=piers)
                if w["zmienionych"]:
                    zrobione += 1
                    wzgledna = p.relative_to(folder_path)
                    self.log(f"  ✔ [{rodzaj}] {wzgledna}: przesunięto wiek w {w['zmienionych']} rekordach"
                             f" (kopia: {p.name}.bak)")
            self.log(f"[MIETKI +{lata} LAT] Gotowe — zmieniono {zrobione} plików.")
            self._plus10_koniec(
                f"Gotowe — wiek przesunięty w {zrobione} plikach ({total_zmienionych} rekordów).",
                folder_wynikow=folder_path)
        except Exception as e:
            self.log(f"[MIETKI +{lata} LAT] BŁĄD: {e}")
            self._plus10_koniec(f"Błąd zapisu: {e}", "#D83B01")
