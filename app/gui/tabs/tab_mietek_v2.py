# -*- coding: utf-8 -*-
"""Mietek v2.0 — edytor danych mietka (MIETEK | Mietek v2.0 — edytor danych).

Przeglądanie i ręczna edycja plików DBF mietka (W/R/O/D/Z/WSIE) w siatce
z polskimi nagłówkami, zapis bezpośrednio do DBF z automatycznym backupem
.BAK oraz generowanie dokumentów nowymi szablonami FORESTLY prosto
z danych (DBF -> TXT z silnika wydruków -> HTML -> PDF, bez Worda;
STR_TYT — z ustawień kreatora, konwersja przez Worda jak w 1-Click).

Metody mietek_v2_* są wywoływane z JavaScriptu przez pywebview.api."""

import os
import shutil
from pathlib import Path

from app.core.leniwe_importy import leniwy

from app.gui.tabs.tab_mietek_plus10 import read_dbf, write_dbf

Document = leniwy("docx", "Document")

# ---------------------------------------------------------------------------
# polskie etykiety pól DBF (reszta pól pokazuje się surową nazwą)
LABELY = {
    "NRREJ": "Nr rej.", "NR_DZIAL": "Nr działki", "POW": "Powierzchnia",
    "POW_L_ZAL": "Pow. zalesiona", "POW_L_NZAL": "Pow. niezales.",
    "POW_N_ZAL": "Pow. nie zales.", "POW_INNE": "Pow. inne",
    "POW_WYDZ": "Pow. wydzielenia", "POW_ZM": "Pow. zmian",
    "ODDZIAL": "Oddział", "PODODDZ": "Pododdz.", "NR_WYDZ": "Nr wydzielenia",
    "NAZWISKO": "Nazwisko", "IMIE": "Imię", "RODZICE": "Rodzice",
    "ADRES": "Adres", "ZM": "Zmiany", "PREJ": "PREJ",
    "GATUNEK": "Gatunek", "WIEK": "Wiek", "KL_WIEK": "Klasa wieku",
    "WYS": "Wysokość", "PIERS": "Pierśnica", "BONIT": "Bonitacja",
    "ZADRZEW": "Zadrzewienie", "JAKOSC": "Jakość", "ZASOB": "Zasobność",
    "PRZES": "Przesunięty",
    "KAT_OCH": "Kat. ochrony", "RODZ_POW": "Rodzaj pow.",
    "CECH_STR": "Cecha struktury", "TYP_SIED": "Typ siedliska",
    "GTD1": "GTD 1", "GTD2": "GTD 2", "GTD3": "GTD 3",
    "OP_TAX": "Opis taksacyjny", "OP_TAX1": "Opis taksacyjny 2",
    "NAZWA": "Nazwa wsi", "WOJEW": "Województwo", "GMINA": "Gmina",
    "STAN_NA": "Stan na", "OBOW_OD": "Obowiązuje od", "OBOW_DO": "Obowiązuje do",
    "NR_WSI": "Nr wsi", "ROK_ZAL": "Rok założenia", "POWIAT": "Powiat",
    "PRZYB": "Przybyło", "UBYL": "Ubyło", "PRZYCZYNY": "Przyczyny",
}
for _i in range(1, 7):
    LABELY[f"WSK{_i}"] = f"Wskazanie {_i}"
    LABELY[f"POW_WSK{_i}"] = f"Pow. wsk. {_i}"
    LABELY[f"MIAZ{_i}"] = f"Miąższość {_i}"

# dokumenty do generowania: klucz JS -> (nazwa pliku TXT, opis)
DOKUMENTY = [
    ("OPTAX", "Opisy taksacyjne lasu (OPTAX)"),
    ("TAB_KLW3", "Tabela klas wieku (TAB_KLW3)"),
    ("ZEST1", "Zestawienie powierzchni (ZEST1)"),
    ("REJESTR1", "Rejestr działek (REJESTR1)"),
    ("WSKAZ1", "Wskazania gospodarcze (WSKAZ1)"),
    ("WYK_NEG", "Wykaz d-stanów negatywnych (WYK_NEG)"),
    ("WSK_ZB", "Czynności na 10-lecie (WSK_ZB)"),
    ("HALIZNY", "Halizny (HALIZNY)"),
]
_OPIS_PLIKU = dict(DOKUMENTY)


class TabMietekV2Mixin:
    """Edytor danych mietka — most między siatką w web UI a DBF-ami."""

    # ------------------------------------------------------------- stan
    def _mv2(self):
        if not hasattr(self, "_mv2_stan"):
            self._mv2_stan = {"obreby": [], "pliki": {}, "seq": 0}
        return self._mv2_stan

    # -------------------------------------------------------- wczytanie
    def mietek_v2_load(self, folder):
        """Skanuje folder mietka i zwraca obręby (foldery z DBF-ami)."""
        try:
            folder = str(folder or "").strip()
            if not folder or not Path(folder).exists():
                return {"ok": False, "blad": "Wskaż folder mietka "
                        "(np. folder z WOL.001 albo z plikami DBF w środku)."}
            root = Path(folder)
            kandydaci = sorted(root.rglob("*.001"))
            if any(k.is_dir() for k in kandydaci):
                kandydaci = [k for k in kandydaci if k.is_dir()]
            else:
                kandydaci = []
            if root != next((k for k in kandydaci), None):
                kandydaci = [root] + kandydaci
            stan = self._mv2()
            stan["obreby"], stan["pliki"] = [], {}
            obreby = []
            for kand in kandydaci:
                pliki = []
                for p in sorted(kand.glob("*")):
                    if p.suffix.upper() != ".DBF" or not p.is_file():
                        continue
                    try:
                        pola, rek = read_dbf(p)
                    except Exception:
                        continue
                    stan["seq"] += 1
                    pid = f"mv2p{stan['seq']}"
                    stan["pliki"][pid] = {"path": str(p), "pola": pola, "rek": rek}
                    typ = p.stem.upper()
                    typ = "WSIE" if typ.startswith("WSIE") else typ[:1]
                    pliki.append({
                        "id": pid, "nazwa": p.name, "typ": typ,
                        "rekordow": len(rek),
                        "kolumny": [
                            {"name": f[0],
                             "label": LABELY.get(f[0], f[0]),
                             "typ": f[1], "len": f[2]}
                            for f in pola],
                    })
                if pliki:
                    obreby.append({
                        "id": f"mv2o{len(obreby)}", "nazwa": kand.name,
                        "folder": str(kand), "pliki": pliki})
            stan["obreby"] = obreby
            if not obreby:
                return {"ok": False, "blad": "Nie znaleziono plików DBF "
                        "w wybranym folderze (ani w podfolderach *.001)."}
            return {"ok": True, "obreby": obreby}
        except Exception as e:
            return {"ok": False, "blad": f"Nie udało się wczytać mietka: {e}"}

    # ----------------------------------------------------------- dane
    def mietek_v2_dane(self, plik_id, szukaj=""):
        """Rekordy jednego pliku (z filtrem tekstowym po wszystkich polach)."""
        info = self._mv2()["pliki"].get(plik_id)
        if info is None:
            return {"ok": False, "blad": "Plik nie jest wczytany — odśwież mietek."}
        kolumny = info["pola"]
        fraza = str(szukaj or "").strip().lower()
        wiersze, idx = [], []
        for i, rec in enumerate(info["rek"]):
            wart = [str(rec.get(f[0], "") or "") for f in kolumny]
            if fraza and not any(fraza in w.lower() for w in wart):
                continue
            wiersze.append(wart)
            idx.append(i)
        return {"ok": True, "kolumny": [
                    {"name": f[0], "label": LABELY.get(f[0], f[0]),
                     "typ": f[1], "len": f[2]} for f in kolumny],
                "wiersze": wiersze, "idx": idx}

    # ---------------------------------------------------------- zapis
    def mietek_v2_zapisz(self, plik_id, zmiany):
        """Zmiany komórek -> DBF z backupem .BAK.

        zmiany: lista {"i": nr rekordu, "pole": nazwa, "v": nowa wartość}."""
        try:
            info = self._mv2()["pliki"].get(plik_id)
            if info is None:
                return {"ok": False, "blad": "Plik nie jest wczytany."}
            pola = {f[0]: f for f in info["pola"]}
            rek = info["rek"]
            odrzucone = []
            n = 0
            for z in (zmiany or []):
                i, pole, v = int(z.get("i", -1)), str(z.get("pole", "")), \
                    str(z.get("v", ""))
                f = pola.get(pole)
                if f is None or i < 0 or i >= len(rek):
                    odrzucone.append(f"wiersz {i + 1}/{pole}")
                    continue
                _, typ, dl, _dec = f
                if typ == "N":
                    v = v.strip().replace(",", ".")
                    if v and not v.replace("-", "").replace(".", "").isdigit():
                        odrzucone.append(f"wiersz {i + 1}/{pole}: '{v}' "
                                         "nie jest liczbą")
                        continue
                elif typ == "L":
                    v = (v[:1] or " ").upper()
                    if v not in ("T", "F", "Y", "N", "?", " "):
                        odrzucone.append(f"wiersz {i + 1}/{pole}: T/F")
                        continue
                else:   # C / D — obcięcie do długości pola
                    v = v[:dl]
                if str(rek[i].get(pole, "") or "") != v:
                    rek[i][pole] = v
                    n += 1
            path = Path(info["path"])
            if n:
                bak = path.with_suffix(path.suffix + ".BAK")
                if not bak.exists():
                    shutil.copy2(path, bak)
                write_dbf(path, info["pola"], rek)
                # odśwież stan z dysku (write_dbf normalizuje wartości)
                info["pola"], info["rek"] = read_dbf(path)
            return {"ok": True, "zapisano": n, "plik": path.name,
                    "bak": str(bak) if n else "", "odrzucone": odrzucone}
        except Exception as e:
            return {"ok": False, "blad": f"Zapis nie udał się: {e}"}

    # ---------------------------------------------------- generowanie
    def _mv2_obreb(self, obr_id):
        for o in self._mv2()["obreby"]:
            if o["id"] == obr_id:
                return o
        return None

    def _mv2_generuj_txt(self, folder, typ):
        """Generuje jeden raport TXT z DBF-ów (silnik wydruków). Zwraca path."""
        from app.core import wydruki
        agencja = wydruki.czytaj_agencje(folder)
        typ = typ.upper()
        if typ == "HALIZNY":
            p, _n = wydruki.generuj_halizny_txt(folder, agencja=agencja)
            return p
        out = wydruki.generuj_wszystkie_po_przeniesieniu(
            folder, agencja=agencja, tylko={typ + ".TXT"})
        return out.get(typ + ".TXT")

    def mietek_v2_generuj(self, obr_id, typ, out_dir, podglad=False):
        """Jeden dokument: DBF -> TXT -> nowy szablon (PDF)."""
        try:
            obr = self._mv2_obreb(obr_id)
            if obr is None:
                return {"ok": False, "blad": "Obręb nie jest wczytany."}
            typ = str(typ or "").upper()
            if typ not in _OPIS_PLIKU:
                return {"ok": False, "blad": f"Nieznany dokument: {typ}"}
            out_dir = Path(out_dir or (Path(obr["folder"]).parent / "wydruki_v2"))
            out_dir.mkdir(parents=True, exist_ok=True)
            txt = self._mv2_generuj_txt(obr["folder"], typ)
            if not txt:
                return {"ok": False, "blad": f"Nie udało się wygenerować danych "
                        f"dla {typ} (brak odpowiednich DBF-ów?)."}
            from app.core import szablony
            pdf = out_dir / f"{typ}.pdf"
            szablony.generuj_raport_pdf(typ, Path(txt), pdf)
            if podglad and os.name == "nt":
                try:
                    os.startfile(str(pdf))
                except Exception:
                    pass
            return {"ok": True, "pdf": str(pdf), "nazwa": _OPIS_PLIKU[typ]}
        except Exception as e:
            return {"ok": False, "blad": f"Generowanie nie wyszło: {e}"}

    def mietek_v2_generuj_wszystko(self, obr_id, out_dir):
        """Komplet dokumentów nowymi szablonami + STR_TYT z kreatora."""
        try:
            obr = self._mv2_obreb(obr_id)
            if obr is None:
                return {"ok": False, "blad": "Obręb nie jest wczytany."}
            out_dir = Path(out_dir or (Path(obr["folder"]).parent / "wydruki_v2"))
            out_dir.mkdir(parents=True, exist_ok=True)
            folder = Path(obr["folder"])
            zrobione, bledy = [], []
            for typ, _opis in DOKUMENTY:
                try:
                    r = self.mietek_v2_generuj(obr_id, typ, str(out_dir))
                    if r.get("ok"):
                        zrobione.append(typ)
                    else:
                        bledy.append(f"{typ}: {r.get('blad', '?')}")
                except Exception as e:
                    bledy.append(f"{typ}: {e}")
            # STR_TYT — z ustawień kreatora (Wersja 1/2/3) + OPTAX
            strtyt_pdf = ""
            try:
                from app.core import szablony
                optax_txt_s = folder / "OPTAX.TXT"
                if optax_txt_s.exists():
                    tpl = self._zbuduj_szablon_str_tyt_dla_all()
                    if tpl:
                        obiekt, _, _ = szablony.meta_z_pliku(optax_txt_s)
                        razem = szablony.razem_z_optax(optax_txt_s)
                        doc = Document(tpl)
                        self.replace_text_robust(doc, "NAZWA WSI",
                                                 obiekt or "NIEZNANA_WIES")
                        self.replace_text_robust(doc, "POWIERZCHNIA_WSI", razem)
                        dpath = out_dir / "STR_TYT.docx"
                        doc.save(str(dpath))
                        try:
                            p = self._docx_na_pdf_wordem(str(dpath))
                        except Exception:
                            p = None
                        if p:
                            strtyt_pdf = p
                            zrobione.append("STR_TYT")
                            try:
                                Path(dpath).unlink()
                            except OSError:
                                pass
                        else:
                            zrobione.append("STR_TYT (tylko DOCX)")
                    try:
                        Path(tpl).unlink()
                    except (OSError, TypeError):
                        pass
            except Exception as e:
                bledy.append(f"STR_TYT: {e}")
            return {"ok": not bledy, "zrobione": zrobione, "bledy": bledy,
                    "folder": str(out_dir), "strtyt_pdf": strtyt_pdf}
        except Exception as e:
            return {"ok": False, "blad": f"Generowanie nie wyszło: {e}"}

    def mietek_v2_otworz(self, path):
        """Otwiera plik/folder w systemie (PDF po podglądzie, folder wyników)."""
        try:
            p = Path(str(path))
            if not p.exists():
                return {"ok": False, "error": "Nie ma takiego pliku."}
            if os.name == "nt":
                os.startfile(str(p))
                return {"ok": True}
            return {"ok": False, "error": "Otwieranie działa na Windows."}
        except Exception as e:
            return {"ok": False, "error": str(e)}
