"""Automatyczne układanie opisów na mapie GEO-MAP.

Problem: opisy (np. „SO35|1.00") mają punkt bazowy w środku wydzielenia i przy
zagęszczeniu nachodzą na siebie oraz na linie poligonów. GEO-MAP trzyma
przesunięcie każdego podpisu w linii ``L``:

    L 2 <dx> <dy> <kąt> <skala> <n>

Domyślnie ``dx = dy = 0``. Ręczne układanie to zmiana tych dwóch liczb.

Ten moduł liczy nowe ``dx/dy`` dla całej mapy tak, aby prostokąty opisów:
  * nie nachodziły na siebie,
  * w miarę możliwości mieściły się w swoim wydzieleniu,
  * a gdy się nie mieszczą — wychodziły na zewnątrz, ale nadal się nie nakładały.

Rozmiar opisu liczymy z wysokości pisma (mm) i skali mapy:

    wysokość opisu [m] = wysokość_pisma_mm * skala / 1000
    szerokość znaku   ≈ 0.62 * wysokość

Tekst dzielimy po „|" na wiersze (tak robi GEO-MAP w podpisach A2).

Zależności: tylko biblioteka standardowa.
"""

import math
import re

# --- domyślne parametry (można nadpisać) ---
WYSOKOSC_MM = 2.5      # wysokość pisma warstwy 5310 wg biblioteki .lay
SKALA = 5000           # skala mapy (1:5000)
SZER_ZNAKU = 0.62      # przybliżony stosunek szerokości znaku do wysokości
ODSTEP = 1.15          # mnożnik odstępu między opisami (>1 = trochę luzu)
KROK = 1.0             # krok siatki wyszukiwania (w wysokościach opisu)


# ------------------------------------------------------------- odczyt mapy

RE_L = re.compile(r"^L\s+(\d+)\s")


def wszystkie_poligony(mapa):
    """Wszystkie obiekty z geometrią (≥3 punkty) — do sprawdzania linii.

    Zwraca listę słowników jak ``poligony_z_mapy``, ale bez filtra na opis.
    """
    lines = mapa["lines"]
    out = []
    i = 0
    n = len(lines)
    while i < n:
        if not lines[i].startswith("*"):
            i += 1
            continue
        j = i + 1
        while j < n and not lines[j].startswith("*"):
            j += 1
        blok = lines[i + 1:j]
        pts = []
        for l in blok:
            if l.startswith("P "):
                q = l.split()
                if len(q) >= 4:
                    try:
                        pts.append((float(q[2]), float(q[3])))
                    except ValueError:
                        pass
        if len(pts) >= 3:
            tekst = litera = ""
            for l in blok:
                if l.startswith(":A2[") and not tekst:
                    tekst = l[4:-1]
                elif l.startswith(":A1[") and not litera:
                    litera = l[4:-1]
            linie_l = []
            for k in range(i + 1, j):
                m = RE_L.match(lines[k])
                if m:
                    linie_l.append((int(m.group(1)), k))
            linie_l.sort()
            # typ linii jest pewny: L 2 = litera (A1), L 3 = opis (A2)
            _op = next((k for t, k in linie_l if t == 3), None)
            _li = next((k for t, k in linie_l if t == 2), None)
            out.append({"start": i, "linie": blok, "punkty": pts,
                        "tekst": tekst, "litera": litera, "linie_L": linie_l,
                        "linia_opisu": _op, "linia_litery": _li})
        i = j
    return out


def poligony_z_mapy(mapa):
    """Zwraca poligony z podpisem A2: geometrię, tekst i indeks linii ``L 2``.

    Każdy element: {"start": i, "linie": [...], "punkty": [(x, y), ...],
    "tekst": "SO35|1.00", "linia_L2": j albo None}
    """
    lines = mapa["lines"]
    out = []
    i = 0
    n = len(lines)
    while i < n:
        if not lines[i].startswith("*"):
            i += 1
            continue
        j = i + 1
        while j < n and not lines[j].startswith("*"):
            j += 1
        blok = lines[i + 1:j]
        pts = []
        for l in blok:
            if l.startswith("P "):
                q = l.split()
                if len(q) >= 4:
                    try:
                        pts.append((float(q[2]), float(q[3])))
                    except ValueError:
                        pass
        if len(pts) >= 3:
            tekst = ""
            for l in blok:
                if l.startswith(":A2["):
                    tekst = l[4:-1]
                    break
            litera = ""
            for l in blok:
                if l.startswith(":A1["):
                    litera = l[4:-1]
                    break
            linie_l = []
            for k in range(i + 1, j):
                m = RE_L.match(lines[k])
                if m:
                    linie_l.append((int(m.group(1)), k))
            # UWAGA (potwierdzone testem u użytkownika): w GEO-MAP podpis
            # oznaczenia (A2, np. „8SO25|0.93") to OSTATNIA linia L obiektu,
            # a wcześniejsza (L 2) to litera wydzielenia (A1, np. „o").
            linie_l.sort()
            linia_opisu = next((k for t, k in linie_l if t == 3), None)
            linia_litery = next((k for t, k in linie_l if t == 2), None)
            if tekst.strip():
                out.append({"start": i, "linie": blok, "punkty": pts,
                            "tekst": tekst, "litera": litera,
                            "linie_L": linie_l,
                            "linia_opisu": linia_opisu,
                            "linia_litery": linia_litery})
        i = j
    return out


# ------------------------------------------------------------- geometria

def _srodek(pts):
    """Środek ciężkości wielokąta (dla niepoprawnego — średnia wierzchołków)."""
    n = len(pts)
    if n < 3:
        return None
    A = cx = cy = 0.0
    for i in range(n):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % n]
        cr = x0 * y1 - x1 * y0
        A += cr
        cx += (x0 + x1) * cr
        cy += (y0 + y1) * cr
    if abs(A) < 1e-9:
        return (sum(p[0] for p in pts) / n, sum(p[1] for p in pts) / n)
    A *= 0.5
    return (cx / (6 * A), cy / (6 * A))


def srodek_bazowy(pts):
    """Punkt bazowy opisu taki, jakiego używa GEO-MAP.

    Ustalone na mapach użytkownika: GEO-MAP NIE bierze środka ciężkości —
    dla nieregularnych wydzieleni różnił się on o 20-60 m. Zgodny jest
    środek prostokąta opisującego (bbox).
    """
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return ((min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0)


def _rozszerz(prost, m):
    """Prostokąt powiększony o margines ``m`` z każdej strony."""
    return (prost[0] - m, prost[1] - m, prost[2] + m, prost[3] + m)


def _box_w_srodku(prost, pts):
    """Czy CAŁY prostokąt (nie tylko środek) leży w wielokącie.

    Sprawdzanie samego środka przepuszczało opisy i litery, które wystawały
    za granicę wydzielenia. Sprawdzamy 4 rogi i środki boków.
    """
    x0, y0, x1, y1 = prost
    for p in ((x0, y0), (x1, y0), (x1, y1), (x0, y1),
              ((x0 + x1) / 2, y0), ((x0 + x1) / 2, y1),
              (x0, (y0 + y1) / 2), (x1, (y0 + y1) / 2)):
        if not _w_srodku(p, pts):
            return False
    return True


def _w_srodku(p, pts):
    """Czy punkt leży w wielokącie (ray casting)."""
    x, y = p
    n = len(pts)
    ins = False
    for i in range(n):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % n]
        if ((y0 > y) != (y1 > y)) and (x < (x1 - x0) * (y - y0) / (y1 - y0) + x0):
            ins = not ins
    return ins


def _przecina_krawedzie(prost, pts):
    """Czy prostokąt podpisu przecina którąś krawędź poligonu.

    Liczy, ile z czterech boków prostokąta ma część wspólną z odcinkami granicy.
    Prosta, zachowawcza heurystyka — wystarcza do „nie nachodzenia na linie".
    """
    x0, y0, x1, y1 = prost
    # szybkie odrzucenie: środek prostokąta daleko od poligonu
    if not _w_srodku(((x0 + x1) / 2, (y0 + y1) / 2), pts) and \
       not _w_srodku((x0, y0), pts) and not _w_srodku((x1, y1), pts) and \
       not _w_srodku((x0, y1), pts) and not _w_srodku((x1, y0), pts):
        return False          # cały na zewnątrz — nie liczymy jako przecięcie
    # czy któryś wierzchołek poligonu leży w prostokącie, albo bok prostokąta
    # przecina bok poligonu — wtedy uznajemy, że podpis wchodzi na linię
    for p in pts:
        if x0 <= p[0] <= x1 and y0 <= p[1] <= y1:
            return True
    boki = [((x0, y0), (x1, y0)), ((x1, y0), (x1, y1)),
            ((x1, y1), (x0, y1)), ((x0, y1), (x0, y0))]
    n = len(pts)
    for a, b in boki:
        for i in range(n):
            if _odcinki(a, b, pts[i], pts[(i + 1) % n]):
                return True
    return False


def _odcinki(a, b, c, d):
    """Czy odcinki ab i cd się przecinają (orientacja + przypadki brzegowe)."""
    def orient(p, q, r):
        v = (q[0] - p[0]) * (r[1] - p[1]) - (q[1] - p[1]) * (r[0] - p[0])
        return (v > 1e-9) - (v < -1e-9)

    def na_odc(p, q, r):
        return (min(p[0], r[0]) - 1e-9 <= q[0] <= max(p[0], r[0]) + 1e-9 and
                min(p[1], r[1]) - 1e-9 <= q[1] <= max(p[1], r[1]) + 1e-9)

    o1, o2 = orient(a, b, c), orient(a, b, d)
    o3, o4 = orient(c, d, a), orient(c, d, b)
    if o1 != o2 and o3 != o4:
        return True
    if o1 == 0 and na_odc(a, c, b):
        return True
    if o2 == 0 and na_odc(a, d, b):
        return True
    if o3 == 0 and na_odc(c, a, d):
        return True
    if o4 == 0 and na_odc(c, b, d):
        return True
    return False


# ------------------------------------------------------------- opisy

def _tekst_a2(o):
    for l in o["linie"]:
        if l.startswith(":A2["):
            return l[4:-1]
    return ""


def _wiersze(tekst):
    """Dzieli oznaczenie na wiersze (GEO-MAP łamie podpis po „|")."""
    t = (tekst or "").strip()
    if not t:
        return []
    return [w.strip() for w in t.split("|") if w.strip()] or [t]


def rozmiar_opisu(tekst, wysokosc_mm=WYSOKOSC_MM, skala=SKALA, obrot=0.0):
    """Zwraca (szerokość, wysokość) prostokąta opisu w jednostkach mapy [m].

    ``obrot`` to kąt w radianach (odpowiednik P3 w GEO-MAP) — wtedy liczymy
    prostokąt opisujący obrócony napis.
    """
    w = _wiersze(tekst)
    if not w:
        return (0.0, 0.0)
    h_linii = wysokosc_mm * skala / 1000.0
    szer = max(len(x) for x in w) * h_linii * SZER_ZNAKU
    wys = len(w) * h_linii * ODSTEP
    if abs(obrot) > 1e-9:
        c, si = abs(math.cos(obrot)), abs(math.sin(obrot))
        szer, wys = szer * c + wys * si, szer * si + wys * c
    return (szer, wys)


def _prost(srodek, rozmiar):
    """Prostokąt (x0, y0, x1, y1) wokół środka."""
    cx, cy = srodek
    w, h = rozmiar
    return (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)


def _nakladka(a, b):
    """Pole części wspólnej dwóch prostokątów."""
    x = min(a[2], b[2]) - max(a[0], b[0])
    y = min(a[3], b[3]) - max(a[1], b[1])
    return x * y if (x > 0 and y > 0) else 0.0


def _odc_przecina_prost(a, b, prost):
    """Czy odcinek a-b przecina prostokąt (albo leży w nim)."""
    x0, y0, x1, y1 = prost
    for p in (a, b):
        if x0 <= p[0] <= x1 and y0 <= p[1] <= y1:
            return True
    boki = [((x0, y0), (x1, y0)), ((x1, y0), (x1, y1)),
            ((x1, y1), (x0, y1)), ((x0, y1), (x0, y0))]
    for c, d in boki:
        if _odcinki(a, b, c, d):
            return True
    return False


def _siatka_krawedzi(poligony, kom):
    """Indeks krawędzi wszystkich poligonów: komórka -> lista odcinków."""
    siatka = {}
    for o in poligony:
        pts = o["punkty"]
        n = len(pts)
        for i in range(n):
            a, b = pts[i], pts[(i + 1) % n]
            cx0 = int(min(a[0], b[0]) // kom)
            cx1 = int(max(a[0], b[0]) // kom)
            cy0 = int(min(a[1], b[1]) // kom)
            cy1 = int(max(a[1], b[1]) // kom)
            for cx in range(cx0, cx1 + 1):
                for cy in range(cy0, cy1 + 1):
                    siatka.setdefault((cx, cy), []).append((a, b))
    return siatka


def _krawedzie_w(prost, siatka, kom):
    """Krawędzie z komórek, które obejmuje prostokąt."""
    x0, y0, x1, y1 = prost
    widz = []
    for cx in range(int(x0 // kom), int(x1 // kom) + 1):
        for cy in range(int(y0 // kom), int(y1 // kom) + 1):
            widz.extend(siatka.get((cx, cy), ()))
    return widz


def _punkt_na_granicy(a, b, pts):
    """Pierwszy punkt przecięcia odcinka a->b z granicą poligonu (albo b)."""
    n = len(pts)
    best = None
    for i in range(n):
        c, d = pts[i], pts[(i + 1) % n]
        t = _przeciecie(a, b, c, d)
        if t is not None and 0.0 <= t <= 1.0:
            if best is None or t < best:
                best = t
    if best is None:
        return b
    return (a[0] + (b[0] - a[0]) * best, a[1] + (b[1] - a[1]) * best)


def _przeciecie(a, b, c, d):
    """Parametr t (na odcinku a->b) przecięcia z odcinkiem c->d, albo None."""
    r = (b[0] - a[0], b[1] - a[1])
    s_ = (d[0] - c[0], d[1] - c[1])
    mian = r[0] * s_[1] - r[1] * s_[0]
    if abs(mian) < 1e-12:
        return None
    ac = (c[0] - a[0], c[1] - a[1])
    t = (ac[0] * s_[1] - ac[1] * s_[0]) / mian
    u = (ac[0] * r[1] - ac[1] * r[0]) / mian
    if 0.0 <= t <= 1.0 and 0.0 <= u <= 1.0:
        return t
    return None


# ------------------------------------------------------------- układanie

def uloz(mapa, wysokosc_mm=WYSOKOSC_MM, skala=SKALA, iteracje=12,
         kara_zewnatrz=20.0, kara_linii=6.0, kara_odleglosci=0.02,
         kara_litery=19.0, kara_opisu=2.0, kara_przeszkody=2.0,
         margines_litery=1.5,
         poligony=None, obrot=0.0, tylko_opisy=True,
         prog_dalekiego=500.0, tylko_srodek=False):
    """Liczy nowe przesunięcia (dx, dy) dla podpisów OPISÓW jednej mapy.

    Ruchome są tylko opisy wydzieleń (tekst z „|"). Wszystkie pozostałe
    etykiety (numery, litery, tytuły) są nieruchomymi przeszkodami — opisy
    mają je omijać. Etykiety świadomie ustawione daleko (> ``prog_dalekiego``)
    zostawiamy w spokoju.

    Zwraca listę elementów ruchomych (opisów) z policzonym ``offset``.
    """
    wszystkie = poligony if poligony is not None else poligony_z_mapy(mapa)
    wszystkie_geo = wszystkie_poligony(mapa)
    h_linii = wysokosc_mm * skala / 1000.0
    elementy = []          # ruchome opisy
    przeszkody = []        # nieruchome etykiety (tylko prostokąty)
    litery = []            # prostokąty liter (nie wolno ich zasłaniać)

    def _off_linii(i):
        if i is None:
            return (0.0, 0.0)
        q = mapa["lines"][i].split()
        if len(q) >= 4:
            try:
                return (float(q[2]), float(q[3]))
            except ValueError:
                pass
        return (0.0, 0.0)

    for o in wszystkie:
        tekst = o["tekst"]
        pts = o["punkty"]
        srodek = srodek_bazowy(pts)
        if srodek is None:
            continue
        off = _off_linii(o.get("linia_opisu"))
        baz = (srodek[0] + off[0], srodek[1] + off[1])   # gdzie opis jest teraz

        # litera (A1) na swoim miejscu — nieruchoma przeszkoda
        lit = (o.get("litera") or "").strip()
        lit_info = None
        if lit and o.get("linia_litery") is not None:
            ol = _off_linii(o["linia_litery"])
            roz_l = rozmiar_opisu(lit, wysokosc_mm, skala, obrot)
            if roz_l[0] > 0:
                pb = _prost((srodek[0] + ol[0], srodek[1] + ol[1]), roz_l)
                przeszkody.append(pb)
                litery.append(pb)
                lit_info = (len(litery) - 1, ol, roz_l)

        roz = rozmiar_opisu(tekst, wysokosc_mm, skala, obrot)
        if roz[0] <= 0:
            continue
        daleko = (off[0] ** 2 + off[1] ** 2) ** 0.5 > prog_dalekiego
        if tylko_opisy and "|" not in tekst:
            przeszkody.append(_prost(baz, roz))
            continue
        if daleko:
            przeszkody.append(_prost(baz, roz))
            continue
        elementy.append({
            "obiekt": o, "tekst": tekst, "pts": pts, "srodek": srodek,
            "rozmiar": roz, "offset": (0.0, 0.0),
            "prost": _prost(srodek, roz), "wewnatrz": True, "wysiegnik": False,
            "lit_info": lit_info, "offset_litery": None,
            "lit_roz": (lit_info[2] if lit_info else (0.0, 0.0)),
            "bb": (max(p[0] for p in pts) - min(p[0] for p in pts),
                   max(p[1] for p in pts) - min(p[1] for p in pts)),
        })

    if not elementy:
        return []

    if tylko_srodek:
        # tryb „wyśrodkowane i obrócone": każdy opis w środku swojego
        # wydzielenia (dx=dy=0), bez rozsuwania i bez wysięgników
        for e in elementy:
            e["offset"] = (0.0, 0.0)
            e["prost"] = _prost(e["srodek"], e["rozmiar"])
            e["wewnatrz"] = True
            e["wysiegnik"] = False
        return elementy

    krok = max(h_linii * KROK, 1e-6)
    # drobne przesunięcia WŁASNEJ litery (gdy opis nie mieści się obok niej)
    krok_lit = max(h_linii * 0.25, 1e-6)
    # pozycje litery: drobno (1 m) i daleko (2 m do 60 m) — litera musi
    # mieć gdzie uciec od opisu, a to bywa kilkadziesiąt metrów
    kand_drob = [(0.0, 0.0)]
    for r in range(1, 17):
        for k in range(16):
            a = k * math.pi / 8
            kand_drob.append((math.cos(a) * r, math.sin(a) * r))
    for r in range(9, 31):
        for k in range(16):
            a = k * math.pi / 8
            kand_drob.append((math.cos(a) * 2.0 * r, math.sin(a) * 2.0 * r))
    kand_lit = [(0.0, 0.0)]
    for r in range(1, 7):
        for k in range(16):
            a = k * math.pi / 8
            kand_lit.append((math.cos(a) * krok_lit * r, math.sin(a) * krok_lit * r))
    # gęsta siatka: co pół kroku i co 22,5° — inaczej drobne, dobre
    # pozycje (jak ręczne ustawienie użytkownika) wypadają między punktami
    kand = [(0.0, 0.0)]
    for r in range(1, 13):
        for k in range(16):
            a = k * math.pi / 8
            kand.append((math.cos(a) * krok * 0.5 * r, math.sin(a) * krok * 0.5 * r))
    kand_lit16 = sorted(kand, key=lambda t: t[0] * t[0] + t[1] * t[1])[:16]
    # pozycje z dala (na zewnątrz wydzielenia) — tu wchodzi wysięgnik
    kand_dal = []
    for r in (7, 9, 11, 13):
        for k in range(16):
            a = k * math.pi / 8
            kand_dal.append((math.cos(a) * krok * r, math.sin(a) * krok * r))

    kom = max(h_linii * 4.0, 1e-6)
    zasieg = kom * 6
    # linie WSZYSTKICH wydzieleń — opis nie może ich przecinać
    kraw = _siatka_krawedzi(wszystkie_geo, kom)

    prob_litery = 0
    prob_drobne = 0
    for _ in range(max(1, iteracje)):
        # siatka obejmuje i opisy, i przeszkody
        siatka = {}
        for idx, el in enumerate(elementy):
            p0 = el["srodek"]
            for cx in range(int((p0[0] - zasieg) // kom), int((p0[0] + zasieg) // kom) + 1):
                for cy in range(int((p0[1] - zasieg) // kom), int((p0[1] + zasieg) // kom) + 1):
                    siatka.setdefault((cx, cy), []).append(("e", idx))
        for idx, pr in enumerate(przeszkody):
            cx0, cy0 = (pr[0] + pr[2]) / 2, (pr[1] + pr[3]) / 2
            for cx in range(int(cx0 // kom) - 2, int(cx0 // kom) + 3):
                for cy in range(int(cy0 // kom) - 2, int(cy0 // kom) + 3):
                    siatka.setdefault((cx, cy), []).append(("p", idx))
        for idx, lb in enumerate(litery):
            cx0, cy0 = (lb[0] + lb[2]) / 2, (lb[1] + lb[3]) / 2
            for cx in range(int(cx0 // kom) - 2, int(cx0 // kom) + 3):
                for cy in range(int(cy0 // kom) - 2, int(cy0 // kom) + 3):
                    siatka.setdefault((cx, cy), []).append(("l", idx))

        przesunieto = 0
        for i, el in enumerate(elementy):
            base = el["srodek"]
            pole = el["rozmiar"][0] * el["rozmiar"][1] + 1e-9
            sasiedzi = siatka.get((int(base[0] // kom), int(base[1] // kom)), [])
            najlepsza = None
            najlepsza_czysta = None
            for dx, dy in (kand + kand_dal):
                cx, cy = base[0] + dx, base[1] + dy
                prost = _prost((cx, cy), el["rozmiar"])
                kara = 0.0
                for typ, jj in sasiedzi:
                    if typ == "e":
                        if jj == i:
                            continue
                        ov = _nakladka(prost, elementy[jj]["prost"])
                    else:
                        ov = _nakladka(prost, przeszkody[jj])
                    if ov:
                        if typ == "e":
                            kara += kara_opisu      # nie nachodzimy na inny opis
                        else:
                            kara += kara_przeszkody  # ani na obcą etykietę
                        if kara >= 1.0:
                            break
                # nie zasłaniamy ŻADNEJ litery wydzielenia (tylko z sąsiedztwa)
                for typ2, jj2 in sasiedzi:
                    if typ2 != "l":
                        continue
                    if _nakladka(prost, litery[jj2]) > 0:
                        kara += kara_litery
                        break
                wewn = _box_w_srodku(prost, el["pts"])
                if not wewn:
                    kara += kara_zewnatrz          # opłaca się zostać w środku
                # zakaz przecinania linii (swojego i obcego wydzielenia)
                for (ka, kb) in _krawedzie_w(prost, kraw, kom):
                    if _odc_przecina_prost(ka, kb, prost):
                        kara += kara_linii
                        break
                kara += kara_odleglosci * math.hypot(dx, dy) / krok
                if najlepsza is None or kara < najlepsza[0] - 1e-12:
                    najlepsza = (kara, dx, dy, prost, wewn)
                if kara < 1.0 and (najlepsza_czysta is None or kara < najlepsza_czysta[0] - 1e-12):
                    najlepsza_czysta = (kara, dx, dy, prost, wewn)
            wybor = najlepsza_czysta or najlepsza
            if wybor is None:
                continue
            _, dx, dy, prost, wewn = wybor
            # Reguła: opis i litera MAJĄ być W ŚRODKU wydzielenia, jeśli się
            # mieszczą. Gdy wybrany opis nachodzi na WŁASNĄ literę (albo litera
            # wypada poza wydzielenie) — przesuwamy literę tak, żeby wyszła
            # z drogi opisowi, ale została w środku.
            if el.get("lit_info"):
                li, ol, roz_l = el["lit_info"]
                akt = el.get("offset_litery") or ol
                lpoz = (base[0] + akt[0], base[1] + akt[1])
                lb0 = _prost(lpoz, roz_l)
                if (_nakladka(prost, lb0) > 0
                        or not _box_w_srodku(lb0, el["pts"])):
                    stara_lb = litery[li]
                    znaleziono = None
                    for ldx, ldy in kand_drob:
                        lpoz2 = (lpoz[0] + ldx, lpoz[1] + ldy)
                        lb = _prost(lpoz2, roz_l)
                        if not _box_w_srodku(lb, el["pts"]):
                            continue
                        zle_l = False
                        for (ka, kb) in _krawedzie_w(lb, kraw, kom):
                            if _odc_przecina_prost(ka, kb, lb):
                                zle_l = True
                                break
                        if zle_l or _nakladka(_rozszerz(prost, margines_litery), lb) > 0:
                            continue
                        # litera nie może też wpaść na INNY opis ani literę
                        for typ2, jj2 in sasiedzi:
                            if typ2 == "e" and jj2 != i and _nakladka(lb, elementy[jj2]["prost"]) > 0:
                                zle_l = True
                                break
                            if typ2 == "l" and jj2 != li and _nakladka(lb, litery[jj2]) > 0:
                                zle_l = True
                                break
                        if zle_l:
                            continue
                        znaleziono = (ldx, ldy, lb)
                        break
                    if znaleziono:
                        ldx, ldy, lb = znaleziono
                        litery[li] = lb
                        el["offset_litery"] = (akt[0] + ldx, akt[1] + ldy)
                    else:
                        litery[li] = stara_lb
            if abs(dx - el["offset"][0]) > 1e-9 or abs(dy - el["offset"][1]) > 1e-9:
                przesunieto += 1
            el["offset"] = (dx, dy)
            el["prost"] = prost
            el["wewnatrz"] = wewn
            el["wysiegnik"] = not wewn
        if przesunieto == 0:
            break
    # ---- końcowy przebieg naprawczy ------------------------------------
    # Żadna litera nie może leżeć pod JAKIMKOLWIEK opisem (także sąsiada).
    # Dla każdej takiej litery szukamy wolnego miejsca w środku wydzielenia.
    for _ in range(4):
        poprawki = 0
        for i, el in enumerate(elementy):
            if not el.get("lit_info"):
                continue
            li, ol, roz_l = el["lit_info"]
            akt = el.get("offset_litery") or ol
            lpoz = (el["srodek"][0] + akt[0], el["srodek"][1] + akt[1])
            lb = _prost(lpoz, roz_l)
            pod = False
            for e2 in elementy:
                if _nakladka(_rozszerz(e2["prost"], margines_litery), lb) > 0:
                    pod = True
                    break
            if not pod and _box_w_srodku(_rozszerz(lb, margines_litery), el["pts"]):
                continue
            stara = litery[li]
            litery[li] = lb
            zn = None
            for ldx, ldy in kand_drob:
                lp2 = (lpoz[0] + ldx, lpoz[1] + ldy)
                lb2 = _prost(lp2, roz_l)
                if not _box_w_srodku(lb2, el["pts"]):
                    continue
                zle = False
                for (ka, kb) in _krawedzie_w(lb2, kraw, kom):
                    if _odc_przecina_prost(ka, kb, lb2):
                        zle = True
                        break
                if zle:
                    continue
                for e2 in elementy:
                    if _nakladka(lb2, e2["prost"]) > 0:
                        zle = True
                        break
                if zle:
                    continue
                for jj2 in range(len(litery)):
                    if jj2 != li and _nakladka(lb2, litery[jj2]) > 0:
                        zle = True
                        break
                if zle:
                    continue
                zn = (ldx, ldy, lb2)
                break
            if zn:
                ldx, ldy, lb2 = zn
                litery[li] = lb2
                el["offset_litery"] = (akt[0] + ldx, akt[1] + ldy)
                poprawki += 1
            else:
                litery[li] = stara
        if poprawki == 0:
            break
    return elementy


# ------------------------------------------------------------- zapis

def _off_linii_litery(mapa, o):
    i = o.get("linia_litery")
    if i is None:
        return (0.0, 0.0)
    q = mapa["lines"][i].split()
    if len(q) >= 4:
        try:
            return (float(q[2]), float(q[3]))
        except ValueError:
            pass
    return (0.0, 0.0)


def ustaw_offsety(mapa, elementy, obrot_rad=0.0):
    """Wstawia policzone dx/dy ORAZ kąt do linii opisu w tekście mapy.

    ``obrot_rad`` to kąt w radianach, jaki ma mieć każdy opis (w GEO-MAP
    ustawia się to jako P3 w gradach; przeliczenie: rad = P3 * pi / 200).
    Kąt zapisywany jest w 5. polu linii ``L``. Dzięki temu opisy są obrócone
    „na równo", a nie każde inaczej.

    Zwraca (nowe_linie, ile_zmian).
    """
    offs = {id(e["obiekt"]): e["offset"] for e in elementy}
    linie = list(mapa["lines"])
    zmiany = 0
    for e in elementy:
        o = e["obiekt"]
        dx, dy = e["offset"]
        i = o.get("linia_opisu")
        if i is None:
            continue
        stara = linie[i]
        p = stara.split()
        if e.get("wysiegnik"):
            # Reguła: opis poza wydzieleniem ZAWSZE ma wysięgnik do SWOJEJ
            # litery, a litera ma być w środku wydzielenia. Kończymy wysięgnik
            # na ŚCIANCE boxa litery (nie w jego środku).
            lit = (e.get("offset_litery") if e.get("offset_litery") is not None
                   else _off_linii_litery(mapa, o))
            lroz = e.get("lit_roz") or (0.0, 0.0)
            lpoz = (e["srodek"][0] + lit[0], e["srodek"][1] + lit[1])
            if lroz[0] > 0:
                cx, cy = lpoz
                hx, hy = lroz[0] / 2.0, lroz[1] / 2.0
                ox, oy = e["srodek"][0] + dx, e["srodek"][1] + dy
                ddx, ddy = ox - cx, oy - cy
                if abs(ddx) < 1e-9:
                    px, py = cx, cy + (hy if ddy >= 0 else -hy)
                elif abs(ddy) < 1e-9:
                    px, py = cx + (hx if ddx >= 0 else -hx), cy
                else:
                    tx = hx / abs(ddx)
                    ty = hy / abs(ddy)
                    if tx <= ty:
                        px, py = cx + (hx if ddx > 0 else -hx), cy + ddy * tx
                    else:
                        px, py = cx + ddx * ty, cy + (hy if ddy > 0 else -hy)
                kon = (px, py)
            else:
                kon = _punkt_na_granicy((e["srodek"][0] + dx, e["srodek"][1] + dy),
                                        lpoz, e["pts"])
            wx, wy = kon[0] - e["srodek"][0], kon[1] - e["srodek"][1]
            nowa = "%s %s %.3f %.3f %.7f 1.0000000 133 %.3f %.3f" % (
                p[0] if p else "L", p[1] if len(p) > 1 else "3",
                dx, dy, obrot_rad, wx, wy)
        elif len(p) >= 6:
            nowa = "%s %s %.3f %.3f %.7f %s" % (p[0], p[1], dx, dy,
                                                obrot_rad, " ".join(p[5:]))
        elif len(p) >= 4:
            nowa = "%s %s %.3f %.3f %.7f" % (p[0], p[1], dx, dy, obrot_rad)
        else:
            nowa = "L 3 %.3f %.3f %.7f 1.0000000 5" % (dx, dy, obrot_rad)
        if nowa != stara:
            linie[i] = nowa
            zmiany += 1
    # przesunięte litery (L 2)
    for e in elementy:
        ol = e.get("offset_litery")
        if ol is None:
            continue
        o = e["obiekt"]
        i = _linia_typu(mapa, o, 2)
        if i is None:
            continue
        stara = linie[i]
        p = stara.split()
        if len(p) >= 6:
            nowa = "%s %s %.3f %.3f %.7f %s" % (p[0], p[1], ol[0], ol[1],
                                                obrot_rad, " ".join(p[5:]))
        else:
            nowa = "L 2 %.3f %.3f %.7f 1.0000000 5" % (ol[0], ol[1], obrot_rad)
        if nowa != stara:
            linie[i] = nowa
            zmiany += 1
    return linie, zmiany


def _linia_typu(mapa, o, typ):
    """Zwraca indeks linii ``L <typ>`` należącej do obiektu (albo None)."""
    start = o.get("start")
    if start is None:
        return None
    for k, l in enumerate(o.get("linie", [])):
        if l.startswith("L "):
            q = l.split()
            if len(q) >= 2 and q[1] == str(typ):
                return start + k
    return None


def policz_kolizje(mapa, lines=None, wysokosc_mm=WYSOKOSC_MM, skala=SKALA,
                   obrot=0.0):
    """Liczy kolizje opisów w mapie (albo w podanych liniach).

    Zwraca słownik: ``litery`` (opis nachodzi na literę), ``opis_opis``,
    ``linie`` (opis przecina linię wydzielenia), ``opisow``, ``wewnatrz``,
    ``wysiegnik``. Służy do kontroli jakości po układaniu — 0/0/0 znaczy czysto.
    """
    if lines is not None:
        m = dict(mapa)
        m["lines"] = lines
        objs = wszystkie_poligony(m)
    else:
        objs = wszystkie_poligony(mapa)
    h = wysokosc_mm * skala / 1000.0

    def _roz(t):
        w = _wiersze(t)
        if not w:
            return (0.0, 0.0)
        a = max(len(x) for x in w) * h * SZER_ZNAKU
        b = len(w) * h * ODSTEP
        if abs(obrot) > 1e-9:
            c, si = abs(math.cos(obrot)), abs(math.sin(obrot))
            a, b = a * c + b * si, a * si + b * c
        return (a, b)

    opisy, litery = [], []
    wewn = wys = 0
    for o in objs:
        pts = o["punkty"]
        if len(pts) < 3:
            continue
        b = srodek_bazowy(pts)
        a1 = a2 = ""
        off = {}
        for l in o["linie"]:
            if l.startswith(":A1[") and not a1:
                a1 = l[4:-1]
            elif l.startswith(":A2[") and not a2:
                a2 = l[4:-1]
            elif l.startswith("L "):
                q = l.split()
                if len(q) >= 4:
                    try:
                        off[int(q[1])] = (float(q[2]), float(q[3]))
                    except ValueError:
                        pass
        if a1 and 2 in off:
            r = _roz(a1)
            if r[0] > 0:
                ox, oy = off[2]
                litery.append((o.get("start"), _prost((b[0] + ox, b[1] + oy), r)))
        if a2 and 3 in off:
            r = _roz(a2)
            if r[0] > 0:
                ox, oy = off[3]
                pkt = (b[0] + ox, b[1] + oy)
                _bx = _prost(pkt, r)
                opisy.append((o.get("start"), a1, _bx))
                if _box_w_srodku(_bx, pts):
                    wewn += 1
                else:
                    wys += 1
    na_lit = sum(1 for (_s, _a, r1) in opisy for (_s2, r2) in litery
                 if _nakladka(r1, r2) > 0)
    oo = sum(1 for i in range(len(opisy)) for j in range(i + 1, len(opisy))
             if opisy[i][0] != opisy[j][0]
             and _nakladka(opisy[i][2], opisy[j][2]) > 0)
    kom = max(h * 4.0, 1e-6)
    kraw = _siatka_krawedzi(objs, kom)
    prz = 0
    for (_s, _a, r) in opisy:
        for (ka, kb) in _krawedzie_w(r, kraw, kom):
            if _odc_przecina_prost(ka, kb, r):
                prz += 1
                break
    return {"litery": na_lit, "opis_opis": oo, "linie": prz,
            "opisow": len(opisy), "wewnatrz": wewn, "wysiegnik": wys}
