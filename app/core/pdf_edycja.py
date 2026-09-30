"""Edycja PDF — dopisywanie tekstu, zakrywanie/zamiana fragmentów, strony.

Zakładka „Edycja PDF": użytkownik wskazuje dowolny plik PDF z dysku,
widzi podgląd stron i nakłada operacje (adnotacje tekstowe, białe
zakrycia — też jako zamiennik błędnego tekstu, usuwanie / obracanie /
przesuwanie stron), a na końcu zapisuje wynik jako NOWY plik PDF
(obok źródłowego, z sufiksem). Oryginał pozostaje nietknięty.

Współrzędne: frontend pracuje na wyrenderowanym obrazie strony w skali
`zoom` (zwracanej przez otworz()); punkty PDF = piksele / zoom.
"""

from pathlib import Path

import base64


def _fitz():
    """PyMuPDF — nowsze API pod nazwą `pymupdf`, starsze jako `fitz`."""
    try:
        import pymupdf as f  # PyMuPDF (nowe API)
        return f
    except Exception:
        import fitz  # starsze wersje
        return f


def _font_dok():
    """DejaVuSans.ttf z folderu programu (polskie znaki); fallback: wbudowany."""
    kand = Path(__file__).resolve().parent.parent.parent / "DejaVuSans.ttf"
    if kand.exists():
        return str(kand)
    return None


# --------------------------------------------------------------------------
# otwarcie pliku: podgląd stron
# --------------------------------------------------------------------------
def _kol_hex(liczba):
    """Kolor spanu PyMuPDF (int sRGB) -> '#RRGGBB'."""
    try:
        c = int(liczba)
        return f"#{c >> 16 & 255:02x}{c >> 8 & 255:02x}{c & 255:02x}"
    except Exception:
        return "#000000"


def _linie_strony(page, limit=1200):
    """Linie tekstu strony (do klikalnej edycji).

    Scalam spany jednej linii w jeden wpis: pozycja bazowej, bbox,
    tekst, rozmiar (najczęstszy span) i kolor. Linie puste pomijam.
    """
    f = _fitz()
    out = []
    try:
        d = page.get_text("dict")
    except Exception:
        return out
    # get_text daje współrzędne w układzie NIEOBRÓCONYM strony, a pixmap
    # (podgląd w edytorze) renderuje układ po obrocie — dla stron z /Rotate
    # (np. mapy w scalonych UPUL) nakładki lądowały w zupełnie innym
    # miejscu niż tekst. Przeliczamy bbox/origin na układ WIZUALNY.
    M = page.rotation_matrix if page.rotation else None
    for blok in d.get("blocks") or []:
        for ln in blok.get("lines") or []:
            spany = [s for s in (ln.get("spans") or [])
                     if str(s.get("text") or "").strip()]
            if not spany:
                continue
            tekst = "".join(str(s.get("text") or "") for s in ln.get("spans") or [])
            if not tekst.strip():
                continue
            bbox = ln.get("bbox") or spany[0]["bbox"]
            rozmiary = {}
            for s in spany:
                rozmiary[round(float(s.get("size") or 11), 1)] = \
                    rozmiary.get(round(float(s.get("size") or 11), 1), 0) + 1
            rozmiar = max(rozmiary, key=rozmiary.get)
            org = spany[0].get("origin") or (bbox[0], bbox[3])
            if M is not None:
                rr = f.Rect(bbox[0], bbox[1], bbox[2], bbox[3]) * M
                oo = f.Point(org[0], org[1]) * M
                bbox = (rr.x0, rr.y0, rr.x1, rr.y1)
                org = (oo.x, oo.y)
            out.append({"x": round(float(org[0]), 1),
                        "y": round(float(org[1]), 1),
                        "bbox": [round(float(bbox[0]), 1),
                                 round(float(bbox[1]), 1),
                                 round(float(bbox[2]), 1),
                                 round(float(bbox[3]), 1)],
                        "tekst": tekst,
                        "rozmiar": round(float(rozmiar), 1),
                        "kolor": _kol_hex(spany[0].get("color", 0)),
                        "font": str(spany[0].get("font") or "")})
            if len(out) >= limit:
                return out
    return out


def otworz(path, zoom=1.6, max_stron=300):
    """Otwiera PDF i renderuje podglądy stron (PNG, base64).

    Zwraca dict {ok, plik, zoom, ile, strony:[{nr, w, h, png, linie}],
    blad?}. `png` to data-URL gotowy do <img src=...>; `linie` to
    linie tekstu (do klikalnej edycji istniejącego tekstu).
    """
    try:
        f = _fitz()
        doc = f.open(str(path))
        strony = []
        for i, page in enumerate(doc):
            if i >= max_stron:
                break
            pix = page.get_pixmap(matrix=f.Matrix(zoom, zoom))
            b64 = base64.b64encode(pix.tobytes("png")).decode("ascii")
            strony.append({"nr": i + 1,
                           "w": round(page.rect.width, 1),
                           "h": round(page.rect.height, 1),
                           "png": "data:image/png;base64," + b64,
                           "linie": _linie_strony(page)})
        n = len(doc)
        doc.close()
        return {"ok": True, "plik": str(path), "zoom": zoom,
                "ile": n, "strony": strony}
    except Exception as e:
        return {"ok": False, "blad": f"Nie udało się otworzyć PDF: {e}"}


def otworz_b64(nazwa, dane_b64, zoom=1.6, max_stron=300):
    """Zapisuje PDF rzucony do okna (base64) w katalogu tymczasowym
    i otwiera go jak otworz(). Zwraca to samo + ścieżkę pliku."""
    import re as _re
    import tempfile
    try:
        czysta = str(dane_b64 or "")
        czysta = _re.sub(r"^data:[^,]+,", "", czysta).strip()
        dane = base64.b64decode(czysta)
        if not dane.startswith(b"%PDF"):
            return {"ok": False,
                    "blad": "To nie wygląda na plik PDF."}
        kat = Path(tempfile.gettempdir()) / "forestly_pdf"
        kat.mkdir(parents=True, exist_ok=True)
        bezpieczna = _re.sub(r"[^A-Za-z0-9_.-]", "_", str(nazwa or "plik.pdf"))
        if not bezpieczna.lower().endswith(".pdf"):
            bezpieczna += ".pdf"
        plik = kat / bezpieczna
        n = 1
        while plik.exists():
            plik = kat / f"{n}_{bezpieczna}"
            n += 1
        plik.write_bytes(dane)
        r = otworz(plik, zoom=zoom, max_stron=max_stron)
        if r.get("ok"):
            r["tymczasowy"] = True
        return r
    except Exception as e:
        return {"ok": False, "blad": f"Nie udało się wczytać pliku: {e}"}


# --------------------------------------------------------------------------
# zapis: zastosowanie operacji do nowego pliku
# --------------------------------------------------------------------------
def _kolor(hexstr, domyslny=(0, 0, 0)):
    s = str(hexstr or "").strip().lstrip("#")
    if len(s) == 6:
        try:
            return (int(s[0:2], 16) / 255.0,
                    int(s[2:4], 16) / 255.0,
                    int(s[4:6], 16) / 255.0)
        except Exception:
            pass
    return domyslny


def _p_page(page, x, y):
    """Punkt z układu WIZUALNEGO (edytor) -> układ strony (PDF)."""
    if page.rotation:
        f = _fitz()
        p = f.Point(float(x), float(y)) * page.derotation_matrix
        return p.x, p.y
    return float(x), float(y)


def _r_page(page, x, y, w, h):
    """Prostokąt z układu wizualnego -> układ strony."""
    r = _fitz().Rect(float(x), float(y),
                     float(x) + float(w), float(y) + float(h))
    if page.rotation:
        r = r * page.derotation_matrix
    return r


def _dopisz_tekst(page, f, x, y, tekst, rozmiar, kolor, fontref=None,
                  fontname="F0"):
    """Wstawia tekst. Jeśli podano fontref (ścieżka do pełnej czcionki
    systemowej LUB bufor wyciągnięty z dokumentu), używa go — dzięki
    czemu poprawiony tekst nie różni się krojem od reszty raportu.
    Bufor z dokumentu bywa subsetem: sprawdzamy pokrycie znaków;
    gdy jakiegoś glifu brak, wracamy do DejaVu (polskie znaki zawsze)."""
    if fontref:
        try:
            if isinstance(fontref, bytes):
                fz = f.Font(fontbuffer=fontref)
            else:
                fz = f.Font(fontfile=fontref)
            braki = [ch for ch in set(tekst)
                     if ch not in (" ", "\t") and not fz.has_glyph(ord(ch))]
            if not braki:
                page.insert_text((x, y), tekst, fontsize=rozmiar,
                                 fontname=fontname, fontfile=fontref,
                                 color=kolor)
                return
        except Exception:
            pass                      # awaryjnie: DejaVu poniżej
    fontplik = _font_dok()
    if fontplik:
        page.insert_text((x, y), tekst, fontsize=rozmiar,
                         fontname="FD", fontfile=fontplik, color=kolor)
    else:
        page.insert_text((x, y), tekst, fontsize=rozmiar,
                         fontname="helv", color=kolor)


# v2.0.129/v2.0.131: mapowanie nazw fontów z PDF na pliki Windows.
# Chromium (generator raportów FORESTLY) osadza w PDF SUBSETY - np.
# AAAAAA+SegoeUI zawiera tylko część glifów, więc bufor wyciągnięty
# z dokumentu nie nadaje się do pisania nowym tekstem. Zamiast tego
# używamy PEŁNYCH czcionek systemowych.
# v2.0.131: PRAWDZIWE nazwy plików Windows (Segoe UI Semibold to
# seguisb.ttf, NIE segoeuisb.ttf — zła nazwa powodowała wpadanie
# w zwykły Bold = grubszy tekst) + LiberationSans (metryczny klon
# Arial — raporty z Linuxa/serwera; bez mapowania trafiał w DejaVu,
# który ma większą wysokość liter = "większy" tekst).
_FONT_RODZINY = {
    "segoeui": "segoeui", "segoeuilight": "segoeuil",
    "calibri": "calibri", "arial": "arial",
    "liberationsans": "arial",        # metrycznie zgodny z Arial
    "liberationserif": "times",
    "liberationmono": "cour",
    "tahoma": "tahoma", "verdana": "verdana", "georgia": "georgia",
    "timesnewroman": "times", "timesnewromanpsmt": "times",
    "timesnewromanps": "times",
    "couriernew": "cour", "trebuchetms": "trebuc",
    "dejavusans": "DejaVuSans", "dejavuserif": "DejaVuSerif",
}
# pliki Segoe UI mają "poszarpane" nazwy — pełna tabela
_FONT_SEGUE = {
    "": "segoeui", "regular": "segoeui",
    "bold": "segoeuib", "italic": "segoeuii", "bolditalic": "segoeuiz",
    "semibold": "seguisb", "semibolditalic": "seguisbi",
    "light": "segoeuil", "semilight": "segoeuisl",
    "black": "seguibl", "blackitalic": "seguibli",
}
# degradacja grubości, gdy pliku nie ma na danym Windowsie
_FONT_DEGRADACJA = {
    "semibolditalic": "bolditalic", "blackitalic": "bolditalic",
    "semibold": "bold", "semilight": "light", "light": "",
    "black": "bold", "bolditalic": "bold", "bold": "", "italic": "",
    "bolditali": "bold", "semibolditali": "semibolditalic",
}
_FONT_ODMIANY = {          # sufiksy plików wg odmiany (rodziny poza Segoe)
    "bold": ("b", "bd"),
    "bolditalic": ("bi", "z", "bz"),
    "italic": ("i",),
    "semibold": ("sb", "b"),
    "semibolditalic": ("sbi", "bi", "z"),
    "light": ("l", "sl"),
    "semilight": ("sl", "l"),
    "black": ("blk", "b"),
}


def _odmiana_znazwy(o):
    """Normalizuje odmianę z nazwy fontu (toleruje skrócone nazwy
    z PDF, np. 'LiberationSans-BoldItali')."""
    o = str(o or "").lower().strip()
    for pelna, prefiksy in (
            ("bolditalic", ("bolditalic", "boldital")),
            ("semibolditalic", ("semibolditalic", "semiboldital")),
            ("blackitalic", ("blackitalic", "blackital")),
            ("semibold", ("semibold",)),
            ("semilight", ("semilight",)),
            ("semibolditalic", ("semibolditali",)),
            ("bold", ("bold",)),
            ("italic", ("italic", "oblique")),
            ("light", ("light",)),
            ("black", ("black",)),
            ("regular", ("regular", "book", "roman"))):
        if o.startswith(prefiksy):
            return pelna
    return ""


def _font_systemowy(nazwa):
    """Ścieżka do pliku czcionki (Windows / katalog aplikacji) dla
    nazwy fontu z get_text, np. 'SegoeUI-Semibold' -> seguisb.ttf.
    Zwraca None, gdy nie można dopasować."""
    import os
    n = str(nazwa or "").strip()
    if not n:
        return None
    n = n.split("+")[-1]               # zetnij prefiks subsetu
    czesci = n.split("-")
    rodzina = czesci[0]
    odmiana = "-".join(czesci[1:]) if len(czesci) > 1 else ""
    if klucz_dejavu(rodzina):           # 'DejaVu Sans Book' itp.
        f = _font_dok()
        return f if f else None
    odm = _odmiana_znazwy(odmiana)
    klucz = rodzina.lower().replace(" ", "")
    if not odm and klucz.endswith("mt"):    # 'ArialMT', 'Arial-BoldMT'
        klucz = klucz[:-2]
    # rodziny typu 'SegoeUIlight' zapisane jednym wyrazem
    if not odm and klucz.startswith("segoeui") and klucz != "segoeui" \
            and klucz[7:] in ("light", "semibold", "semilight", "black",
                              "bold", "italic"):
        odm = klucz[7:]
        klucz = "segoeui"
    if klucz.startswith("dejavu"):
        f = _font_dok()
        return f if f else None
    rdzen = _FONT_RODZINY.get(klucz)
    if not rdzen:
        # 'ArialBold' / 'Arial-ItalicMT' sklejone w jedną nazwę
        for odmiana_sklej in ("bolditalic", "semibolditalic", "semibold",
                              "semilight", "bold", "italic", "light",
                              "black"):
            if klucz.endswith(odmiana_sklej):
                klucz = klucz[:-len(odmiana_sklej)]
                odm = odm or odmiana_sklej
                rdzen = _FONT_RODZINY.get(klucz)
                break
    if not rdzen:
        return None
    kandydaci = []
    if klucz == "segoeui":
        o = odm
        odwiedzone = set()
        while o is not None and o not in odwiedzone:
            odwiedzone.add(o)
            if o in _FONT_SEGUE:
                kandydaci.append(_FONT_SEGUE[o] + ".ttf")
            o = _FONT_DEGRADACJA.get(o)
    else:
        for suf in _FONT_ODMIANY.get(odm, ()):
            kandydaci.append(rdzen + suf + ".ttf")
        kandydaci.append(rdzen + ".ttf")   # bez odmiany na końcu
    katalki = [os.path.join(os.environ.get("WINDIR", r"C:\Windows"),
                            "Fonts")]
    try:
        katalki.append(os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        pass
    for kat in katalki:
        for f in kandydaci:
            sc = os.path.join(kat, f)
            if os.path.exists(sc):
                return sc
    return None


def klucz_dejavu(rodzina):
    r = str(rodzina or "").lower().replace(" ", "")
    return r.startswith("dejavu")


def _font_z_dokumentu(doc, page, nazwa):
    """Bufor fontu osadzonego w dokumencie (po nazwie z get_text).

    Raporty FORESTLY mają osadzone SegoeUI/Semibold/Italic itd. —
    wyciągnięty bufor pozwala wstawić poprawkę DOKŁADNIE tym samym
    krojem. Nazwa może mieć prefiks subsetu (ABCDEF+SegoeUI)."""
    if not nazwa:
        return None
    try:
        nr_strony = page.number
        for xref, _ext, _typ, basefont, _ref, _enc in \
                doc.get_page_fonts(nr_strony):
            bf = str(basefont or "")
            if (nazwa == bf or nazwa in bf
                    or bf.split("+")[-1] == nazwa):
                _b, _e, _t, buf = doc.extract_font(xref)
                if buf:
                    return buf
    except Exception:
        pass
    return None


def zapisz(path, operacje, cel=None):
    """Stosuje operacje i zapisuje PDF.

    Gdy cel wskazuje na plik oryginalny — zapis W MIEJSCU: przed pierwszym
    nadpisaniem powstaje kopia .BAK (pierwsza kopia zostaje, jak przy DBF).

    operacje = {
      "edycje":  [{"nr", "x", "y", "tekst", "rozmiar", "kolor",
                    "bbox": [x0, y0, x1, y1]}],
                  # REALNA edycja istniejącego tekstu: stara linia jest
                  # usuwana z PDF (redakcja), a nowa wstawiana w jej
                  # miejsce (ten sam początek, rozmiar i kolor)
      "teksty":  [{"nr", "x", "y", "tekst", "rozmiar", "kolor"}],
      "zakrycia": [{"nr", "x", "y", "w", "h", "nowy_tekst"?, "rozmiar"?,
                    "kolor"?}],       # białe pole (+ opcjonalnie nowy tekst)
      "usun":    [nr...],
      "obroc":   [{"nr", "kat"}],     # kat: 90 / 180 / 270
      "kolejnosc":[nr...]             # nowa kolejność wszystkich stron
    }
    Współrzędne w punktach PDF. Zwraca {ok, plik, ile, blad?}.
    """
    operacje = dict(operacje or {})
    try:
        import os as _os
        import shutil as _shutil
        f = _fitz()
        src = Path(path)
        if cel:
            cel = Path(cel)
            if cel.suffix.lower() != ".pdf":
                cel = cel.with_suffix(".pdf")
        else:
            cel = src.with_name(src.stem + "_POPRAWIONY.pdf")
        # zapis W MIEJSCU (cel == oryginał): kopia .BAK (pierwsza zostaje),
        # zapis przez plik tymczasowy — inaczej nie nadpisujemy istniejących
        w_miejscu = False
        try:
            w_miejscu = cel.resolve() == src.resolve()
        except Exception:
            w_miejscu = str(cel) == str(src)
        if w_miejscu:
            bak = src.with_suffix(src.suffix + ".BAK")
            if not bak.exists():
                _shutil.copy2(str(src), str(bak))
        else:
            bazowy = cel
            n = 1
            while cel.exists():
                n += 1
                cel = bazowy.with_name(f"{bazowy.stem}_{n}{bazowy.suffix}")

        doc = f.open(str(src))

        # --- 0. EDYCJE istniejącego tekstu: najpierw usuwamy stare linie
        #     (redakcja), dopiero potem wstawiamy nowe — inaczej nowy
        #     tekst też zostałby zredagowany
        edycje = operacje.get("edycje") or []
        strony_z_edycjami = {}
        for e in edycje:
            try:
                nr = int(e.get("nr", 1))
                strony_z_edycjami.setdefault(nr, []).append(e)
            except Exception:
                continue
        for nr, lista in strony_z_edycjami.items():
            try:
                page = doc[nr - 1]
                for e in lista:
                    b = e.get("bbox")
                    if not b or len(b) != 4:
                        continue
                    kat = _r_page(page, float(b[0]) - 1, float(b[1]) - 1,
                                 (float(b[2]) - float(b[0])) + 2,
                                 (float(b[3]) - float(b[1])) + 2)
                    if kat.is_empty or kat.is_infinite:
                        continue
                    page.add_redact_annot(kat)
                # usuwamy tylko tekst: obrazki i grafika zostają
                try:
                    page.apply_redactions(images=f.PDF_REDACT_IMAGE_NONE)
                except TypeError:
                    page.apply_redactions()  # starsze PyMuPDF
            except Exception:
                continue
        for e in edycje:
            try:
                tekst = str(e.get("tekst") or "")
                if not tekst:
                    continue
                page = doc[int(e.get("nr", 1)) - 1]
                ex, ey = _p_page(page, float(e.get("x") or 50),
                                 float(e.get("y") or 50))
                # v2.0.129: ta sama czcionka co edytowany wiersz;
                # fontname unikalny per KROJ (na stronie może być
                # kilka różnych czcionek edytowanych wierszy).
                # Najpierw PEŁNA czcionka systemowa Windows (subsety
                # w PDF mają okrojone glify), potem awaryjnie bufor
                # z dokumentu, na końcu DejaVu.
                fref = _font_systemowy(e.get("font"))
                if not fref:
                    fref = _font_z_dokumentu(doc, page, e.get("font"))
                fn = "FE" + "".join(c for c in str(e.get("font") or "X")
                                    if c.isalnum())[:20]
                _dopisz_tekst(page, f, ex, ey, tekst,
                              float(e.get("rozmiar") or 11),
                              _kolor(e.get("kolor")), fontref=fref,
                              fontname=fn)
            except Exception:
                continue

        # --- 1. teksty i zakrycia (na stronach źródłowych) ---
        for t in (operacje.get("teksty") or []):
            try:
                tekst = str(t.get("tekst") or "")
                if not tekst:
                    continue
                page = doc[int(t.get("nr", 1)) - 1]
                tx, ty = _p_page(page, float(t.get("x") or 50),
                                 float(t.get("y") or 50))
                _dopisz_tekst(page, f, tx, ty, tekst,
                              float(t.get("rozmiar") or 11),
                              _kolor(t.get("kolor")))
            except Exception:
                continue

        for z in (operacje.get("zakrycia") or []):
            try:
                page = doc[int(z.get("nr", 1)) - 1]
                x, y = float(z.get("x") or 0), float(z.get("y") or 0)
                w, h = float(z.get("w") or 10), float(z.get("h") or 10)
                # białe pole zakrywające (współrzędne wizualne edytora)
                rp = _r_page(page, x, y, w, h)
                page.draw_rect(rp, color=None, fill=(1, 1, 1))
                # ...i opcjonalny nowy tekst na nim (zamiana)
                nowy = str(z.get("nowy_tekst") or "")
                if nowy:
                    roz = float(z.get("rozmiar") or 11)
                    zx, zy = _p_page(page, x + 2, y + roz + 2)
                    _dopisz_tekst(page, f, zx, zy, nowy, roz,
                                  _kolor(z.get("kolor")))
            except Exception:
                continue

        # --- 2. obracanie ---
        for o in (operacje.get("obroc") or []):
            try:
                page = doc[int(o.get("nr", 1)) - 1]
                page.set_rotation((page.rotation
                                   + int(o.get("kat") or 90)) % 360)
            except Exception:
                continue

        # --- 3. usunięcie stron / nowa kolejność ---
        def _num(li):
            out = []
            for k in li or []:
                try:
                    out.append(int(k))
                except Exception:
                    pass
            return out

        usun = set(_num(operacje.get("usun")))
        kol = _num(operacje.get("kolejnosc"))
        zmiana = None
        if kol and sorted(kol) == list(range(1, len(doc) + 1)):
            zmiana = kol                        # pełna permutacja
        elif usun:
            zmiana = [i + 1 for i in range(len(doc)) if (i + 1) not in usun]
        if zmiana:
            doc2 = f.open()
            try:
                for nr in zmiana:
                    if 1 <= nr <= len(doc):
                        doc2.insert_pdf(doc, from_page=nr - 1, to_page=nr - 1)
                doc.close()
                doc = doc2
            except Exception:
                doc2.close()

        if w_miejscu:
            tmp = src.with_name(src.stem + "._tmp_zapis.pdf")
            doc.save(str(tmp), garbage=3, deflate=True)
            n = len(doc)
            doc.close()
            _os.replace(str(tmp), str(src))
            return {"ok": True, "plik": str(src), "ile": n,
                    "bak": str(bak)}
        doc.save(str(cel), garbage=3, deflate=True)
        n = len(doc)
        doc.close()
        return {"ok": True, "plik": str(cel), "ile": n}
    except Exception as e:
        return {"ok": False, "blad": f"Nie udało się zapisać PDF: {e}"}
