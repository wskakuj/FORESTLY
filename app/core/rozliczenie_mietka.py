# -*- coding: utf-8 -*-
"""
Forestly — Czyszczenie rejestru (usuwanie pozycji nierozliczonych)
=================================================================

Zakładka MIETEK | Czyszczenie rejestru. Dla każdego obrębu (folderu z plikami
DBF mietka, także w podfolderach) sprawdza i usuwa pozycje NIEROZLICZONE:

  * DZIAŁKI (D*.DBF) — rekordy, którym NIE PRZYPISANO PODODDZIAŁU (litery),
    np. wydzielenie „1a” ma literę „a”, a „1” jej nie ma → działka
    nierozliczona,
  * WŁAŚCICIELE / POZYCJE REJESTROWE (W*.DBF) — rekordy, których numer rejestru
    (NRREJ) nie występuje przy żadnej rozliczonej działce (czyli w W jest
    więcej pozycji rejestrowych, niż D ma rozliczonych) → sierocy właściciele
    usuwani razem ze swoją pozycją rejestrową.

Dodatkowo można ODŚWIEŻYĆ już wygenerowane raporty (nowe szablony, HTML + PDF),
żeby nie pokazywały usuniętych pozycji — funkcja odswiez_raporty().

Bezpieczeństwo:
  * przed każdym zapisem powstaje kopia <plik>.DBF.BAK (pierwsza kopia zostaje
    nietknięta — jak w edytorze Mietek v2.0),
  * tryb „tylko raport” niczego nie zapisuje (podgląd),
  * pliki O*.DBF (opis taksacyjny) i R*.DBF nie są zmieniane.

Wykaz usuniętych pozycji trafia do raportu TXT (funkcja zbuduj_raport).
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from app.core.kontrola_pow import _read_dbf, _write_dbf, _bak


# ---------------------------------------------------------------------------
# pomocnicze
# ---------------------------------------------------------------------------

def _znajdz_dbf(obreb, litera):
    """Pierwszy plik <litera>NNN....DBF w obrębie (np. D0011019.DBF).

    Świadomie NIE używa zwykłego prefiksu, żeby „W” nie złapało WSIE.DBF.
    """
    obreb = Path(obreb)
    wzor = re.compile(rf"^{re.escape(litera)}\d", re.I)
    kand = [p for p in obreb.rglob("*.DBF") if wzor.match(p.stem)]
    if not kand:
        kand = [p for p in obreb.rglob("*.DBF")
                if p.stem.upper() == litera.upper()]
    return sorted(kand)[0] if kand else None


def _ma_litere(rec):
    """Czy działka ma przypisany pododdział (literę wydzielenia)."""
    return bool(str(rec.get("PODODDZ", "") or "").strip())


def _oddzp(rec):
    oddz = str(rec.get("ODDZIAL", "") or "").strip()
    pod = str(rec.get("PODODDZ", "") or "").strip()
    return f"{oddz}{pod}".strip() or "(brak wydzielenia)"


def _liczba(v):
    try:
        return float(str(v).replace(',', '.').strip() or 0)
    except (TypeError, ValueError):
        return 0.0


def _nrrej(v):
    try:
        return int(float(str(v).strip() or 0))
    except (TypeError, ValueError):
        return 0


def _nazwa(rec):
    if not rec:
        return ""
    czesci = [str(rec.get("NAZWISKO", "") or "").strip(),
              str(rec.get("IMIE", "") or "").strip()]
    return " ".join(x for x in czesci if x).strip()


def znajdz_obreby(root):
    """Foldery-obręby w drzewie (te, które zawierają pliki DBF mietka)."""
    root = Path(root)
    obreby = {p.parent for p in root.rglob("*.DBF")
              if re.match(r"^[ODW]\d", p.stem, re.I)}
    if not obreby:
        obreby = {p.parent for p in root.rglob("*.DBF")}
    return sorted(obreby)


# ---------------------------------------------------------------------------
# czyszczenie jednego obrębu
# ---------------------------------------------------------------------------

def przeczysc_obreb(obreb_dir, usun=True):
    """Sprawdza (i opcjonalnie usuwa) nierozliczone pozycje jednego obrębu.

    Zwraca słownik z wynikiem:
      obreb, o_path/d_path/w_path, d_rekordow, w_rekordow,
      d_usuniete (lista), w_usuniete (lista), bak (lista kopii), blad.
    """
    obreb = Path(obreb_dir)
    wynik = {
        "obreb": obreb.name, "sciezka": str(obreb),
        "o_path": None, "d_path": None, "w_path": None,
        "d_rekordow": 0, "w_rekordow": 0,
        "d_usuniete": [], "w_usuniete": [], "bak": [], "blad": None,
    }
    o_path = _znajdz_dbf(obreb, "O")
    d_path = _znajdz_dbf(obreb, "D")
    w_path = _znajdz_dbf(obreb, "W")
    wynik["o_path"], wynik["d_path"], wynik["w_path"] = o_path, d_path, w_path

    if d_path is None:
        wynik["blad"] = "brak pliku D*.DBF (działki)"
        return wynik

    # 1) właściciele (do dopisania nazwiska/adresu w raporcie)
    w_fields, w_recs = (None, [])
    if w_path is not None:
        w_fields, w_recs = _read_dbf(str(w_path))
        wynik["w_rekordow"] = len(w_recs)
    w_by_nr = {}
    for r in w_recs:
        w_by_nr.setdefault(_nrrej(r.get("NRREJ")), r)

    # 2) działki: rozliczone = mają przypisany pododdział (literę)
    d_fields, d_recs = _read_dbf(str(d_path))
    wynik["d_rekordow"] = len(d_recs)
    d_ok, d_nierozl = [], []
    for r in d_recs:
        (d_ok if _ma_litere(r) else d_nierozl).append(r)
    nrrej_ok = {_nrrej(r.get("NRREJ")) for r in d_ok}

    for r in d_nierozl:
        nr = _nrrej(r.get("NRREJ"))
        wl = w_by_nr.get(nr)
        wynik["d_usuniete"].append({
            "nr_dzial": str(r.get("NR_DZIAL", "") or "").strip(),
            "oddzp": _oddzp(r),
            "pow": _liczba(r.get("POW")),
            "nrrej": nr,
            "wlasciciel": _nazwa(wl),
            "adres": str((wl or {}).get("ADRES", "") or "").strip(),
        })

    # 3) właściciele: zostają tylko ci z rozliczoną działką
    w_ok = []
    if w_path is not None:
        for r in w_recs:
            if _nrrej(r.get("NRREJ")) in nrrej_ok:
                w_ok.append(r)
            else:
                wynik["w_usuniete"].append({
                    "nrrej": _nrrej(r.get("NRREJ")),
                    "wlasciciel": _nazwa(r),
                    "adres": str(r.get("ADRES", "") or "").strip(),
                })

    # 4) zapis (z kopią .BAK)
    if usun and (d_nierozl or wynik["w_usuniete"]):
        if d_nierozl:
            b = _bak(d_path)
            if b:
                wynik["bak"].append(b)
            _write_dbf(str(d_path), d_fields, d_ok)
        if w_path is not None and wynik["w_usuniete"]:
            b = _bak(w_path)
            if b:
                wynik["bak"].append(b)
            _write_dbf(str(w_path), w_fields, w_ok)

    return wynik


# ---------------------------------------------------------------------------
# odświeżenie wygenerowanych raportów (nowe szablony: HTML + PDF)
# ---------------------------------------------------------------------------

def odswiez_raporty(obreb_dir, raporty_dir, typy=("REJESTR1", "ZEST1"),
                    margins=None, czcionki=None):
    """Generuje na nowo raporty (HTML + PDF) z już oczyszczonego mietka.

    Raporty pokazujące działki (REJESTR1, ZEST1) budowane są od nowa
    z plików DBF obrębu — dzięki temu nie zawierają usuniętych pozycji.
    Zapis: <raporty_dir>/<TYP>.html oraz <raporty_dir>/pdf/<TYP>.pdf.
    Zwraca listę odświeżonych typów.
    """
    from app.core.wydruki import generuj_wszystkie_po_przeniesieniu
    from app.core import szablony

    raporty_dir = Path(raporty_dir)
    (raporty_dir / "pdf").mkdir(parents=True, exist_ok=True)
    tylko = {f"{t}.TXT" for t in typy}
    txts = generuj_wszystkie_po_przeniesieniu(obreb_dir, tylko=tylko)
    zrobione = []
    try:
        for typ in typy:
            txt = txts.get(f"{typ}.TXT")
            if not txt:
                continue
            szablony.generuj_raport_pdf(
                typ, txt, raporty_dir / "pdf" / f"{typ}.pdf",
                margins=margins, czcionki=czcionki,
                html_out=raporty_dir / f"{typ}.html")
            zrobione.append(typ)
    finally:
        # TXT to pliki pośrednie — nie zostawiamy ich w folderze mietka
        for p in txts.values():
            try:
                Path(p).unlink()
            except OSError:
                pass
    return zrobione


# ---------------------------------------------------------------------------
# raport
# ---------------------------------------------------------------------------

def zbuduj_raport(wyniki, dry_run=False, kiedy=None):
    """Buduje tekstowy raport z listy wyników przeczysc_obreb()."""
    kiedy = kiedy or datetime.now().strftime("%Y-%m-%d %H:%M")
    linie = []
    linie.append("CZYSZCZENIE REJESTRU — wykaz pozycji nierozliczonych")
    linie.append("(działki bez przypisanego pododdziału oraz właściciele "
                 "bez rozliczonej działki)")
    linie.append(f"Data: {kiedy}")
    linie.append("Tryb: " + ("PODGLĄD — nic nie usunięto" if dry_run else
                             "USUWANIE — przed każdym zapisem powstaje kopia .BAK"))
    linie.append("=" * 78)

    suma_d = suma_w = 0
    obreby_z_zmianami = 0
    for w in wyniki:
        d_u, w_u = w["d_usuniete"], w["w_usuniete"]
        if not d_u and not w_u and not w["blad"]:
            continue
        obreby_z_zmianami += 1
        linie.append("")
        linie.append(f"Obręb: {w['obreb']}")
        if w["blad"]:
            linie.append(f"  UWAGA: {w['blad']} — obręb pominięty.")
            continue
        linie.append(f"  Rekordów w D*.DBF: {w['d_rekordow']}, "
                     f"w W*.DBF: {w['w_rekordow']}")
        linie.append(f"  Usunięte działki (bez przypisanej litery): {len(d_u)}")
        if d_u:
            linie.append("    nr działki | oddz/pod | pow. [ha] | nr rej | "
                         "właściciel | adres")
            for p in d_u:
                linie.append("    " + " | ".join([
                    p["nr_dzial"] or "-", p["oddzp"], f"{p['pow']:.4f}",
                    str(p["nrrej"] or "-"), p["wlasciciel"] or "-",
                    p["adres"] or "-"]))
        linie.append(f"  Usunięci właściciele / pozycje rejestrowe "
                     f"(bez rozliczonej działki): {len(w_u)}")
        if w_u:
            linie.append("    nr rej | właściciel | adres")
            for p in w_u:
                linie.append("    " + " | ".join([
                    str(p["nrrej"] or "-"), p["wlasciciel"] or "-",
                    p["adres"] or "-"]))
        if w.get("raporty"):
            linie.append("  Odświeżone raporty: " + ", ".join(w["raporty"]))
        if w["bak"]:
            linie.append("  Kopie zapasowe (.BAK): "
                         + ", ".join(Path(b).name for b in w["bak"]))
        suma_d += len(d_u)
        suma_w += len(w_u)

    linie.append("")
    linie.append("=" * 78)
    linie.append(f"PODSUMOWANIE: obrębów z pozycjami do usunięcia: "
                 f"{obreby_z_zmianami}; działek: {suma_d}; "
                 f"właścicieli/pozycji: {suma_w}.")
    if dry_run and (suma_d or suma_w):
        linie.append("To był PODGLĄD — uruchom ponownie bez „Tylko raport”, "
                     "aby usunąć te pozycje.")
    if not suma_d and not suma_w:
        linie.append("Nic nie znaleziono — rejestr jest rozliczony.")
    return "\n".join(linie)
