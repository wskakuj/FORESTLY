import re
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


_OGONKI = {"Ł": "L", "ł": "l", "Ó": "O", "ó": "o", "Ą": "A", "ą": "a",
           "Ę": "E", "ę": "e", "Ć": "C", "ć": "c", "Ś": "S", "ś": "s",
           "Ź": "Z", "ź": "z", "Ż": "Z", "ż": "z", "Ń": "N", "ń": "n"}


def _bez_ogonkow(tekst):
    """Polskie znaki na czytelne dla GEO-MAP (Ł->L, Ą->A, Ć->C ...)."""
    if not tekst:
        return tekst
    return "".join(_OGONKI.get(ch, ch) for ch in tekst)


def _punkty_srodkowe(pts, krok=2.0):
    """Punkty wewnątrz wielokąta, posortowane OD NAJBARDZIEJ ŚRODKOWEGO.

    Dla siatki o kroku ``krok`` liczymy, jak daleko każdy punkt leży od granicy
    wydzielenia, i sortujemy malejąco. Dzięki temu pierwsze sprawdzane pozycje
    to te najbardziej „w środku" — a nie przypadkowe punkty siatki biegunowej.
    """
    try:
        import numpy as np
        from matplotlib.path import Path as MPath
    except Exception:
        return []
    xs = [q[0] for q in pts]
    ys = [q[1] for q in pts]
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    nx = int((x1 - x0) / krok) + 1
    ny = int((y1 - y0) / krok) + 1
    if nx < 2 or ny < 2 or nx * ny > 400000:
        return []
    gx = x0 + np.arange(nx) * krok
    gy = y0 + np.arange(ny) * krok
    GX, GY = np.meshgrid(gx, gy)
    P = np.column_stack([GX.ravel(), GY.ravel()])
    path = MPath(np.asarray(pts, dtype=float), closed=True)
    w = path.contains_points(P).reshape(ny, nx)
    if not w.any():
        return []
    # odległość od granicy: iteracyjne „erozje" (tania transformata odległości)
    d = w.astype(np.int32)
    dist = np.zeros_like(d, dtype=np.float32)
    cur = d.copy()
    r = 0
    while cur.any() and r < 60:
        dist += cur
        e = cur.copy()
        e[1:, :] &= cur[:-1, :]
        e[:-1, :] &= cur[1:, :]
        e[:, 1:] &= cur[:, :-1]
        e[:, :-1] &= cur[:, 1:]
        cur = e
        r += 1
    dist *= krok
    idx = np.argsort(-dist.ravel())
    out = []
    for i in idx:
        if dist.ravel()[i] <= 0:
            break
        out.append((float(P[i, 0]), float(P[i, 1]), float(dist.ravel()[i])))
    return out


def _wydzielenia(objs):
    """Tylko prawdziwe WYDZIELENIA — nie drogi, granice, ramki mapy.

    Wydzielenie rozpoznajemy po literze (A1) albo po opisie taksacyjnym
    (A2 z „|"). Dzięki temu opisy nie „przecinają" linii dróg i granic,
    które nie są granicami wydzieleń.
    """
    out = []
    for o in objs:
        a1 = (o.get("litera") or "").strip()
        a2 = (o.get("tekst") or "")
        if a1 or ("|" in a2):
            out.append(o)
    return out


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

def uloz(mapa, wysokosc_mm=WYSOKOSC_MM, skala=SKALA, iteracje=12, tryb="wolne",
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
    if tryb == "wolne":
        return uloz_wolne(mapa, wysokosc_mm=wysokosc_mm, skala=skala,
                          obrot=obrot, margines_litery=margines_litery)
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
    kraw = _siatka_krawedzi(_wydzielenia(wszystkie_geo), kom)

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
    # ---- doszukiwanie dla opisów, które wyszły na zewnątrz ---------------
    # Siatka zgrubna bywa za rzadka — dla każdego opisu, który został na
    # zewnątrz, robimy GĘSTE przeszukanie (1,5 m, 16 kierunków, do 120 m),
    # i bierzemy pierwszą pozycję, gdzie cały opis leży w środku bez kolizji.
    krok_g = max(h_linii * 0.12, 1.5)
    for _ in range(2):
        poprawki = 0
        for i, el in enumerate(elementy):
            if el.get("wewnatrz"):
                continue
            base = el["srodek"]
            roz = el["rozmiar"]
            pts = el["pts"]
            lb_wlasna = None
            li = el.get("lit_info")
            if li:
                akt = el.get("offset_litery") or li[1]
                lb_wlasna = _prost((base[0] + akt[0], base[1] + akt[1]), li[2])
            najlepsza = None
            rmax = int(120.0 / krok_g) + 1
            for r in range(0, rmax):
                for k in range(16 if r else 1):
                    a = k * math.pi / 8
                    dx, dy = math.cos(a) * krok_g * r, math.sin(a) * krok_g * r
                    p2 = _prost((base[0] + dx, base[1] + dy), roz)
                    if not _box_w_srodku(p2, pts):
                        continue
                    zle = False
                    for (ka, kb) in _krawedzie_w(p2, kraw, kom):
                        if _odc_przecina_prost(ka, kb, p2):
                            zle = True
                            break
                    if zle:
                        continue
                    if lb_wlasna is not None and _nakladka(p2, lb_wlasna) > 0:
                        continue
                    for jj2 in range(len(litery)):
                        if _nakladka(p2, litery[jj2]) > 0:
                            zle = True
                            break
                    if zle:
                        continue
                    for j2, e2 in enumerate(elementy):
                        if j2 != i and _nakladka(p2, e2["prost"]) > 0:
                            zle = True
                            break
                    if zle:
                        continue
                    najlepsza = (dx, dy, p2)
                    break
                if najlepsza:
                    break
            if najlepsza:
                dx, dy, p2 = najlepsza
                el["offset"] = (dx, dy)
                el["prost"] = p2
                el["wewnatrz"] = True
                el["wysiegnik"] = False
                poprawki += 1
        if poprawki == 0:
            break

    # ---- doszukiwanie dla opisów, które wyszły na zewnątrz ---------------
    krok_g = max(h_linii * 0.12, 1.5)
    for _ in range(2):
        poprawki = 0
        for i, el in enumerate(elementy):
            if el.get("wewnatrz"):
                continue
            base = el["srodek"]; roz = el["rozmiar"]; pts = el["pts"]
            lb_w = None
            li = el.get("lit_info")
            if li:
                akt = el.get("offset_litery") or li[1]
                lb_w = _prost((base[0] + akt[0], base[1] + akt[1]), li[2])
            najlepsza = None
            for r in range(0, int(120.0 / krok_g) + 1):
                for k in range(16 if r else 1):
                    a = k * math.pi / 8
                    dx, dy = math.cos(a) * krok_g * r, math.sin(a) * krok_g * r
                    p2 = _prost((base[0] + dx, base[1] + dy), roz)
                    if not _box_w_srodku(p2, pts):
                        continue
                    zle = False
                    for (ka, kb) in _krawedzie_w(p2, kraw, kom):
                        if _odc_przecina_prost(ka, kb, p2):
                            zle = True; break
                    if zle: continue
                    if lb_w is not None and _nakladka(p2, lb_w) > 0: continue
                    for jj2 in range(len(litery)):
                        if _nakladka(p2, litery[jj2]) > 0:
                            zle = True; break
                    if zle: continue
                    for j2, e2 in enumerate(elementy):
                        if j2 != i and _nakladka(p2, e2["prost"]) > 0:
                            zle = True; break
                    if zle: continue
                    najlepsza = (dx, dy, p2); break
                if najlepsza: break
            if najlepsza:
                dx, dy, p2 = najlepsza
                el["offset"] = (dx, dy); el["prost"] = p2
                el["wewnatrz"] = True; el["wysiegnik"] = False
                poprawki += 1
        if poprawki == 0:
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


def _dlugosc_odc_w_prost(a, b, prost):
    """Długość części odcinka a-b leżącej wewnątrz prostokąta (x0,y0,x1,y1).

    Służy do wyboru końca kreski opisu, z którego wysięgnik nie przecina
    tekstu opisu."""
    x0, y0, x1, y1 = prost
    dx, dy = b[0] - a[0], b[1] - a[1]
    t0, t1 = 0.0, 1.0
    # Liang-Barsky: t musi spełniać lo <= wsp + t*d <= hi dla każdej osi
    for d, lo, hi, wsp in ((dx, x0, x1, a[0]), (dy, y0, y1, a[1])):
        if abs(d) < 1e-12:
            if wsp < lo or wsp > hi:
                return 0.0
            continue
        r0, r1 = (lo - wsp) / d, (hi - wsp) / d
        if r0 > r1:
            r0, r1 = r1, r0
        t0 = max(t0, r0)
        t1 = min(t1, r1)
        if t0 > t1:
            return 0.0
    return max(0.0, t1 - t0) * math.hypot(dx, dy)


def _koniec_wys(lit_srodek, lit_roz, opis_srodek):
    """Punkt na ŚCIANCE prostokąta litery, od strony opisu (koniec wysięgnika).

    Gdy litera nie ma rozmiaru — zwraca jej środek."""
    cx, cy = lit_srodek
    if not lit_roz or lit_roz[0] <= 0:
        return (cx, cy)
    hx, hy = lit_roz[0] / 2.0, lit_roz[1] / 2.0
    ddx, ddy = opis_srodek[0] - cx, opis_srodek[1] - cy
    if abs(ddx) < 1e-9:
        return (cx, cy + (hy if ddy >= 0 else -hy))
    if abs(ddy) < 1e-9:
        return (cx + (hx if ddx >= 0 else -hx), cy)
    tx, ty = hx / abs(ddx), hy / abs(ddy)
    if tx <= ty:
        return (cx + (hx if ddx > 0 else -hx), cy + ddy * tx)
    return (cx + ddx * ty, cy + (hy if ddy > 0 else -hy))


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
            # rozmiar litery: w układaniu „wolne" nie ma klucza „lit_roz",
            # jest w lit_info — bez tego wysięgnik kończył się na GRANICY
            # wydzielenia (a nie na literze) i przechodził przez tekst opisu
            lroz = e.get("lit_roz")
            if not lroz:
                _li = e.get("lit_info")
                lroz = _li[2] if _li else (0.0, 0.0)
            lpoz = (e["srodek"][0] + lit[0], e["srodek"][1] + lit[1])
            # UWAGA: offset litery bywa dopisany dopiero po tym miejscu, więc
            # strona liczona jest z OSTATECZNYCH pozycji opis↔litera
            opis_x = e["srodek"][0] + dx
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
            # STRONA WYSIĘGNIKA. Wysięgnik wychodzi z tego KOŃCA kreski opisu,
            # z którego odcinek do końca wysięgnika NIE przecina prostokąta
            # opisu (mniejsze przecięcie = czytelniej).
            # KODOWANIE W PLIKU GEO-MAP — potwierdzone na zrzutach ekranu
            # użytkownika: flaga 133 = koniec od strony WIĘKSZEGO X,
            # 69 = od strony MNIEJSZEGO X. (Wcześniejsze „odwrócenie" było
            # błędne i to ono pchało wysięgnik przez własny tekst.)
            _box = _prost((e["srodek"][0] + dx, e["srodek"][1] + dy),
                          e["rozmiar"])
            _cy = (e["srodek"][1] + dy)
            _l = _dlugosc_odc_w_prost((_box[0], _cy), kon, _box)
            _p = _dlugosc_odc_w_prost((_box[2], _cy), kon, _box)
            # mniejsze przecięcie decyduje; 69 = mniejszy X, 133 = większy X
            flaga = 69 if _l < _p else 133
            nowa = "%s %s %.3f %.3f %.7f 1.0000000 %d %.3f %.3f" % (
                p[0] if p else "L", p[1] if len(p) > 1 else "3",
                dx, dy, obrot_rad, flaga, wx, wy)
        elif len(p) >= 6:
            # Opis BEZ wysięgnika: piszemy linię bez ogona wysięgnika.
            # UWAGA: nie wolno zachować starego ogona z pliku wejściowego
            # (flaga 69/133 + koniec) — inaczej opis, który teraz mieści się
            # w środku wydzielenia, dalej nosi w pliku flagę wysięgnika
            # z poprzedniego układania i GEO-MAP rysuje niepotrzebny wysięgnik.
            skala = p[5] if len(p) > 5 and p[5] else "1.0000000"
            nowa = "%s %s %.3f %.3f %.7f %s 5" % (p[0], p[1], dx, dy,
                                                  obrot_rad, skala)
        elif len(p) >= 4:
            nowa = "%s %s %.3f %.3f %.7f" % (p[0], p[1], dx, dy, obrot_rad)
        else:
            nowa = "L 3 %.3f %.3f %.7f 1.0000000 5" % (dx, dy, obrot_rad)
        if nowa != stara:
            linie[i] = nowa
            zmiany += 1
    # tekst opisu (A2) bez polskich ogonków — GEO-MAP ich nie wyświetla
    for e in elementy:
        o = e["obiekt"]
        start = o.get("start")
        if start is None:
            continue
        for k, l in enumerate(o.get("linie", [])):
            if l.startswith(":A2["):
                i = start + k
                if 0 <= i < len(linie):
                    nowy = ":A2[" + _bez_ogonkow(l[4:-1]) + "]"
                    if nowy != linie[i]:
                        linie[i] = nowy
                        zmiany += 1
                break
    # tekst opisu (A2) bez polskich ogonków
    for e in elementy:
        o = e["obiekt"]; start = o.get("start")
        if start is None:
            continue
        for k, l in enumerate(o.get("linie", [])):
            if l.startswith(":A2["):
                i = start + k
                if 0 <= i < len(linie):
                    nowy = ":A2[" + _bez_ogonkow(l[4:-1]) + "]"
                    if nowy != linie[i]:
                        linie[i] = nowy; zmiany += 1
                break
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
    # NA KOŃCU: popraw STRONĘ WYSIĘGNIKA z tego, co faktycznie jest w liniach.
    # Liczymy z gotowego pliku (nie z pamięci), więc flaga zawsze zgadza się
    # ze stroną, po której stoi litera.
    for e in elementy:
        if not e.get("wysiegnik"):
            continue
        o = e["obiekt"]
        i_op = _linia_typu(mapa, o, 3)
        i_li = _linia_typu(mapa, o, 2)
        if i_op is None:
            continue
        q = linie[i_op].split()
        if len(q) < 9:
            continue
        try:
            opis_x = e["srodek"][0] + float(q[2])
            opis_y = e["srodek"][1] + float(q[3])
        except ValueError:
            continue
        if i_li is not None:
            ql = linie[i_li].split()
            try:
                lit_x = e["srodek"][0] + float(ql[2])
                lit_y = e["srodek"][1] + float(ql[3])
            except (ValueError, IndexError):
                lit_x, lit_y = opis_x, opis_y
        else:
            lit_x, lit_y = opis_x, opis_y
        # strona: koniec kreski, z którego wysięgnik nie przecina opisu
        # (69 = strona MNIEJSZEGO X, 133 = WIĘKSZEGO X — tak jest w pliku
        #  GEO-MAP, potwierdzone na zrzutach użytkownika)
        try:
            _box = _prost((opis_x, opis_y), e["rozmiar"])
            try:
                _kon = (e["srodek"][0] + float(q[7]), e["srodek"][1] + float(q[8]))
            except (ValueError, IndexError):
                _kon = (lit_x, lit_y)
            _l = _dlugosc_odc_w_prost((_box[0], opis_y), _kon, _box)
            _p = _dlugosc_odc_w_prost((_box[2], opis_y), _kon, _box)
            flaga = 69 if _l < _p else 133
        except Exception:
            flaga = 69 if lit_x <= opis_x else 133
        nowa = "%s %s %s %s %s %s %d %s %s" % (
            q[0], q[1], q[2], q[3], q[4], q[5], flaga, q[7], q[8])
        if nowa != linie[i_op]:
            linie[i_op] = nowa
            zmiany += 1

    # NA KOŃCU: usuń zdublowane linie L i atrybuty (podwójne litery w GEO-MAP).
    # Obiekty przetwarzamy od KOŃCA, żeby usunięcia nie psuły indeksów
    # wcześniejszych obiektów.
    for e in reversed(elementy):
        try:
            zmiany += _posprzataj_duplikaty(linie, e["obiekt"])
        except Exception:
            pass
    return linie, zmiany


RE_ATR = re.compile(r"^:([A-Za-z0-9]+)\[(.*)\]\s*$")


def _posprzataj_duplikaty(linie, o):
    """Usuwa ZDUBLOWANE linie ``L`` i atrybuty (A1/A2/...) w bloku obiektu.

    GEO-MAP rysuje literę dla KAŻDEJ linii ``L 2`` — jeśli w pliku są dwie,
    litera pokazuje się podwójnie, a skasowanie jednej usuwa obie. Ten sam
    problem dotyczy zdublowanego atrybutu ``:A2[...]``. Zostawiamy pierwszy
    wpis, resztę usuwamy.
    """
    start = o.get("start")
    if start is None:
        return 0
    blok = o.get("linie") or []
    # koniec bloku: pierwsza linia po nim, która nie należy do bloku
    koniec = start + len(blok)
    if koniec > len(linie):
        koniec = len(linie)
    widziane_l = set()
    widziane_a = set()
    do_usuniecia = []
    for k in range(start + 1, koniec):
        l = linie[k]
        if l.startswith("L "):
            q = l.split()
            klucz = q[1] if len(q) > 1 else "?"
            if klucz in widziane_l:
                do_usuniecia.append(k)
            else:
                widziane_l.add(klucz)
        else:
            mm = RE_ATR.match(l)
            if mm:
                if mm.group(1) in widziane_a:
                    do_usuniecia.append(k)
                else:
                    widziane_a.add(mm.group(1))
    for k in reversed(do_usuniecia):
        del linie[k]
    return len(do_usuniecia)


def _linia_typu(mapa, o, typ):
    """Zwraca indeks linii ``L <typ>`` należącej do obiektu (albo None)."""
    start = o.get("start")
    if start is None:
        return None
    # UWAGA: o["linie"] to blok BEZ wiersza '*' (wszystkie_poligony robi
    # blok = lines[i+1:j]), więc indeks w bloku to start + 1 + k.
    for k, l in enumerate(o.get("linie", [])):
        if l.startswith("L "):
            q = l.split()
            if len(q) >= 2 and q[1] == str(typ):
                return start + 1 + k
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
        if a2 and "|" in a2 and 3 in off:
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
    kraw = _siatka_krawedzi(_wydzielenia(objs), kom)
    prz = 0
    for (_s, _a, r) in opisy:
        for (ka, kb) in _krawedzie_w(r, kraw, kom):
            if _odc_przecina_prost(ka, kb, r):
                prz += 1
                break
    return {"litery": na_lit, "opis_opis": oo, "linie": prz,
            "opisow": len(opisy), "wewnatrz": wewn, "wysiegnik": wys}


# ==================================================== dlaczego na zewnątrz

POWODY = {
    "za_maly": "wydzielenie mniejsze niż opis — nie ma jak zmieścić",
    "brak_wewnatrz": "brak jakiejkolwiek pozycji w środku (kształt wydzielenia)",
    "linie": "w środku zawsze przecina linię wydzielenia",
    "litera": "w środku zawsze wchodzi na literę (swoją lub sąsiada)",
    "opis": "w środku zawsze wchodzi na inny opis",
    "inny": "inne",
}


def raport_zewnatrz(mapa, elementy, wysokosc_mm=WYSOKOSC_MM, skala=SKALA,
                    obrot=0.0, margines_litery=1.5, krok_drob=2.0, zasieg=120.0):
    """Dla każdego opisu, który wyszedł NA ZEWNĄTRZ — ustala PRZYCZYNĘ.

    Sprawdza po kolei, co blokuje umieszczenie opisu w środku:
      1) czy w ogóle istnieje pozycja, gdzie CAŁY opis leży w środku,
      2) czy któraś z nich nie przecina linii,
      3) czy któraś nie wchodzi na literę,
      4) czy któraś nie wchodzi na inny opis.
    Pierwszy warunek, który nie ma spełnienia, jest przyczyną.

    Zwraca listę słowników: {a6, litera, tekst, powod, opis, powod_tekst}.
    """
    h = wysokosc_mm * skala / 1000.0
    kom = max(h * 4.0, 1e-6)
    geo = wszystkie_poligony(mapa)
    kraw = _siatka_krawedzi(_wydzielenia(geo), kom)
    litery = []
    for e in elementy:
        li = e.get("lit_info")
        if not li:
            continue
        li_i, ol, roz_l = li
        akt = e.get("offset_litery") or ol
        lb = _prost((e["srodek"][0] + akt[0], e["srodek"][1] + akt[1]), roz_l)
        litery.append((id(e), lb))
    opisy = [(id(e), e["prost"]) for e in elementy]

    out = []
    for e in elementy:
        if not e.get("wysiegnik"):
            continue
        o = e["obiekt"]
        a6 = ""
        for l in o.get("linie", []):
            if l.startswith(":A6["):
                a6 = l[4:-1]
                break
        base = e["srodek"]
        roz = e["rozmiar"]
        pts = e["pts"]
        bb = e.get("bb") or (0.0, 0.0)
        if bb[0] < roz[0] or bb[1] < roz[1]:
            out.append({"a6": a6, "litera": o.get("litera", ""),
                        "tekst": e["tekst"], "powod": "za_maly",
                        "powod_tekst": POWODY["za_maly"]})
            continue
        ma_wewnatrz = ma_linie = ma_litere = ma_opis = False
        krok = krok_drob
        rmax = int(zasieg / krok) + 1
        for r in range(0, rmax):
            for k in range(16 if r else 1):
                a = k * math.pi / 8
                dx, dy = math.cos(a) * krok * r, math.sin(a) * krok * r
                p2 = _prost((base[0] + dx, base[1] + dy), roz)
                if not _box_w_srodku(p2, pts):
                    continue
                ma_wewnatrz = True
                if any(_odc_przecina_prost(ka, kb, p2)
                       for (ka, kb) in _krawedzie_w(p2, kraw, kom)):
                    continue
                ma_linie = True
                if any(_nakladka(_rozszerz(p2, margines_litery), lb) > 0
                       for (_i, lb) in litery if _i != id(e)):
                    continue
                ma_litere = True
                if any(_nakladka(p2, pr) > 0 for (_i, pr) in opisy if _i != id(e)):
                    continue
                ma_opis = True
                break
            if ma_opis:
                break
        if not ma_wewnatrz:
            powod = "brak_wewnatrz"
        elif not ma_linie:
            powod = "linie"
        elif not ma_litere:
            powod = "litera"
        elif not ma_opis:
            powod = "opis"
        else:
            powod = "inny"
        out.append({"a6": a6, "litera": o.get("litera", ""), "tekst": e["tekst"],
                    "powod": powod, "powod_tekst": POWODY[powod]})
    return out


# ====================================================== układanie „wolne"

def _kand_opisu(el, krok):
    """Kandydaci na pozycję OPISU — od najbardziej środkowego."""
    if "kand_off" not in el:
        b = el["srodek"]
        el["kand_off"] = [(x - b[0], y - b[1])
                          for (x, y, _d) in _punkty_srodkowe(el["pts"], krok)]
    return el["kand_off"]


def _kand_litery_prio(prost, roz_l, home, krok=2.0, zasieg=140.0):
    """Kandydaci dla LITERY w kolejności preferencji.

    PRIORYTET (potwierdzony przez użytkownika): litera ma stać Z LEWEJ STRONY
    OPISU, w tej samej linii (wyśrodkowana w pionie względem opisu). Dopiero
    gdy tam się nie zmieści — z prawej, potem nad/pod, na końcu dalej.
    """
    x0, y0, x1, y1 = prost
    cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
    hw, hh = roz_l[0] / 2.0, roz_l[1] / 2.0
    out = []
    # 1) z lewej, w linii opisu
    out.append((x0 - hw - krok, cy))
    for k in range(2, 8):
        out.append((x0 - hw - krok * k, cy))
    # 2) z prawej, w linii opisu
    for k in range(1, 8):
        out.append((x1 + hw + krok * k, cy))
    # 3) nad i pod opstem (wyśrodkowane w poziomie)
    for k in range(1, 8):
        out.append((cx, y0 - hh - krok * k))
        out.append((cx, y1 + hh + krok * k))
    # 4) ukosy
    for k in range(1, 8):
        out.append((x0 - hw - krok * k, y0 - hh - krok * k))
        out.append((x1 + hw + krok * k, y0 - hh - krok * k))
        out.append((x0 - hw - krok * k, y1 + hh + krok * k))
        out.append((x1 + hw + krok * k, y1 + hh + krok * k))
    # 5) wokół domu (gdyby powyższe nie wyszły)
    r = 1
    while r * krok <= zasieg:
        for k in range(16):
            a = k * math.pi / 8
            out.append((home[0] + math.cos(a) * krok * r,
                        home[1] + math.sin(a) * krok * r))
        r += 1
    return out


def _kand_litery(pts, base, home, krok=2.0, zasieg=140.0):
    """Kandydaci dla LITERY — od najbliższego jej miejscu domowemu."""
    out = [(home[0], home[1])]
    r = 1
    while r * krok <= zasieg:
        for k in range(16):
            a = k * math.pi / 8
            out.append((home[0] + math.cos(a) * krok * r,
                        home[1] + math.sin(a) * krok * r))
        r += 1
    return out


def uloz_wolne(mapa, wysokosc_mm=WYSOKOSC_MM, skala=SKALA, obrot=0.0,
               krok_srodkowy=3.0, margines_litery=0.5, maks_kand=600,
               krok_litery=2.0, iteracje=3, prog_dalekiego=500.0):
    """Układanie opisów „przez WOLNE OBSZARY".

    Kolejność jest odwrotna niż dotąd: najpierw opis zajmuje najbardziej
    ŚRODKOWE wolne miejsce w wydzieleniu (tak, jakby litery nie było), a
    dopiero potem LITERA ustępuje — szuka najbliższego wolnego miejsca, które
    nie wchodzi na żaden opis. Jeśli litera nie ma gdzie uciec, opis próbuje
    kolejne, coraz mniej środkowe miejsce. Dzięki temu oba elementy mieszczą
    się w środku, a nie ma nakładek.
    """
    wszystkie_geo = wszystkie_poligony(mapa)
    # TYLKO opisy taksacyjne (oznaczenie z „|") — nie każdy obiekt z A2.
    # UWAGA: bierzemy je z TEJ SAMEJ listy co litery (jedno parsowanie!),
    # inaczej id() obiektów nie pasuje i litery nie są dopasowywane.
    wszystkie = [o for o in wszystkie_geo if "|" in (o.get("tekst") or "")]
    h_linii = wysokosc_mm * skala / 1000.0

    # litery WSZYSTKICH wydzieleń (także tych bez opisu taksacyjnego)
    litery = []
    idx_lit = {}
    for o in wszystkie_geo:
        lit = (o.get("litera") or "").strip()
        if not lit:
            continue
        pts_l = o["punkty"]
        if len(pts_l) < 3:
            continue
        b_l = srodek_bazowy(pts_l)
        if b_l is None:
            continue
        rl = rozmiar_opisu(lit, wysokosc_mm, skala, obrot)
        if rl[0] <= 0:
            continue
        ol = _off_linii_litery(mapa, o)
        litery.append(_prost((b_l[0] + ol[0], b_l[1] + ol[1]), rl))
        idx_lit[id(o)] = len(litery) - 1

    elementy = []
    for o in wszystkie:
        pts = o["punkty"]
        srodek = srodek_bazowy(pts)
        if srodek is None:
            continue
        roz = rozmiar_opisu(o["tekst"], wysokosc_mm, skala, obrot)
        if roz[0] <= 0:
            continue
        ol = _off_linii_litery(mapa, o)
        lit = (o.get("litera") or "").strip()
        roz_l = rozmiar_opisu(lit, wysokosc_mm, skala, obrot) if lit else (0.0, 0.0)
        li = idx_lit.get(id(o))
        el = {"obiekt": o, "tekst": o["tekst"], "pts": pts, "srodek": srodek,
              "rozmiar": roz, "offset": (0.0, 0.0),
              "prost": _prost(srodek, roz), "wewnatrz": False, "wysiegnik": False,
              "lit_info": (li, ol, roz_l) if li is not None else None,
              "offset_litery": None,
              "bb": (max(q[0] for q in pts) - min(q[0] for q in pts),
                     max(q[1] for q in pts) - min(q[1] for q in pts))}
        elementy.append(el)

    if not elementy:
        return []

    kom = max(h_linii * 4.0, 1e-6)
    kraw = _siatka_krawedzi(_wydzielenia(wszystkie_geo), kom)

    def _przecina(prost):
        for (ka, kb) in _krawedzie_w(prost, kraw, kom):
            if _odc_przecina_prost(ka, kb, prost):
                return True
        return False

    for _ in range(max(1, iteracje)):
        zmiany = 0
        for i, el in enumerate(elementy):
            base = el["srodek"]
            roz = el["rozmiar"]
            pts = el["pts"]
            li_info = el.get("lit_info")
            wybor = None
            zn_lit = None
            # Zapas od granicy wydzielenia: najpierw z luzem (1 m),
            # a gdy w wydzieleniu nie ma tyle miejsca — mniejszy
            # (0,5 potem 0,25 m). Dzięki temu opis, który mieści się
            # „na styk", zostaje W ŚRODKU, zamiast wychodzić na
            # zewnątrz z wysięgnikiem.
            for marg_wew in (1.0, 0.5, 0.25):
                for (dx, dy) in _kand_opisu(el, krok_srodkowy)[:maks_kand]:
                    prost = _prost((base[0] + dx, base[1] + dy), roz)
                    if not _box_w_srodku(_rozszerz(prost, marg_wew), pts):
                        continue
                    if _przecina(prost):
                        continue
                    zle = False
                    for j2, e2 in enumerate(elementy):
                        if j2 != i and _nakladka(prost, e2["prost"]) > 0:
                            zle = True
                            break
                    if zle:
                        continue
                    # opis nie może wchodzić na CUDZE litery (własna ustąpi sama)
                    _wlasna = el["lit_info"][0] if el.get("lit_info") else None
                    for jj in range(len(litery)):
                        if jj != _wlasna and _nakladka(prost, litery[jj]) > 0:
                            zle = True
                            break
                    if zle:
                        continue
                    # opis OK — teraz litera musi ustąpić
                    if li_info is None:
                        wybor = (dx, dy, prost)
                        zn_lit = None
                        break
                    li, ol, roz_l = li_info
                    home = (base[0] + ol[0], base[1] + ol[1])
                    znal = None
                    for (lx, ly) in _kand_litery_prio(prost, roz_l, home, krok_litery):
                        lb = _prost((lx, ly), roz_l)
                        if not _box_w_srodku(_rozszerz(lb, margines_litery), pts):
                            continue
                        if _przecina(lb):
                            continue
                        if _nakladka(_rozszerz(prost, max(margines_litery, 2.5)), lb) > 0:
                            continue
                        zle = False
                        for j2, e2 in enumerate(elementy):
                            if j2 != i and _nakladka(e2["prost"], lb) > 0:
                                zle = True
                                break
                        if zle:
                            continue
                        for jj in range(len(litery)):
                            if jj != li and _nakladka(litery[jj], lb) > 0:
                                zle = True
                                break
                        if zle:
                            continue
                        znal = (lx, ly, lb)
                        break
                    if znal is None:
                        continue          # litera nie ma gdzie uciec — następny środek
                    if math.hypot(dx, dy) > 45.0:
                        continue          # nie oddalamy opisu daleko od środka
                    wybor = (dx, dy, prost)
                    zn_lit = (li, znal)
                    break

                if wybor is not None:
                    break
            if wybor is None:
                # Brak miejsca w środku — odsuwamy opis NA ZEWNĄTRZ, ale
                # DALeko od granicy (żeby nie leżał na linii) i tak, by nie
                # wchodził na inne opisy. Wtedy dostaje wysięgnik.
                _kand_out = []
                for r in range(1, 22):
                    for k in range(24):
                        a = k * math.pi / 12
                        dxx = math.cos(a) * (25.0 + r * 6.0)
                        dyy = math.sin(a) * (25.0 + r * 6.0)
                        p2 = _prost((base[0] + dxx, base[1] + dyy), roz)
                        if _przecina(p2):
                            continue
                        zle = False
                        for j2, e2 in enumerate(elementy):
                            if j2 != i and _nakladka(p2, e2["prost"]) > 0:
                                zle = True
                                break
                        if not zle:
                            _wl = el["lit_info"][0] if el.get("lit_info") else None
                            for jj in range(len(litery)):
                                if jj != _wl and _nakladka(p2, litery[jj]) > 0:
                                    zle = True
                                    break
                        if zle:
                            continue
                        _kand_out.append((dxx, dyy, p2))
                # Z kandydatów wybieramy taki, przy którym WYSIĘGNIK NIE
                # PRZECINA WŁASNEGO OPISU (inaczej linia idzie przez tekst —
                # tak było, gdy opis leżał nad/pod literą). Dopiero potem
                # decyduje odległość od środka wydzielenia.
                if _kand_out:
                    def _przeciecie_wys(kand):
                        dxx_, dyy_, pbox_ = kand
                        opis_c = (base[0] + dxx_, base[1] + dyy_)
                        if li_info is not None and li_info[2][0] > 0:
                            lit_c = (base[0] + li_info[1][0],
                                     base[1] + li_info[1][1])
                            kon_ = _koniec_wys(lit_c, li_info[2], opis_c)
                        else:
                            kon_ = _punkt_na_granicy(opis_c, base, pts)
                        return min(
                            _dlugosc_odc_w_prost((pbox_[0], opis_c[1]), kon_, pbox_),
                            _dlugosc_odc_w_prost((pbox_[2], opis_c[1]), kon_, pbox_))
                    _kand_out.sort(key=lambda kk: (round(_przeciecie_wys(kk), 2),
                                                   math.hypot(kk[0], kk[1])))
                    wybor = _kand_out[0]
            if wybor is None:
                el["wewnatrz"] = False
                el["wysiegnik"] = True
                continue
            dx, dy, prost = wybor
            if abs(dx - el["offset"][0]) > 1e-9 or abs(dy - el["offset"][1]) > 1e-9:
                zmiany += 1
            el["offset"] = (dx, dy)
            el["prost"] = prost
            # „w środku" liczone CAŁYM prostokątem z marginesem — jeśli opis
            # wyszedł poza wydzielenie, MUSI dostać wysięgnik
            el["wewnatrz"] = _box_w_srodku(prost, pts)
            el["wysiegnik"] = not el["wewnatrz"]
            if zn_lit is not None:
                li, (lx, ly, lb) = zn_lit
                litery[li] = lb
                el["offset_litery"] = (lx - base[0], ly - base[1])
        if zmiany == 0:
            break

    # ---- na koniec: żadna litera nie może leżeć pod opisem ----
    for _ in range(6):
        poprawki = 0
        for e in elementy:
            li = e.get("lit_info")
            if not li:
                continue
            li_i, ol, roz_l = li
            akt = e.get("offset_litery") or ol
            lpoz = (e["srodek"][0] + akt[0], e["srodek"][1] + akt[1])
            lb = _prost(lpoz, roz_l)
            pod = False
            for e2 in elementy:
                if _nakladka(_rozszerz(e2["prost"], margines_litery), lb) > 0:
                    pod = True
                    break
            if not pod and _box_w_srodku(_rozszerz(lb, margines_litery), e["pts"]):
                continue
            stara = litery[li_i]
            litery[li_i] = lb
            zn = None
            for (lx, ly) in _kand_litery_prio(e["prost"], roz_l, lpoz):
                lb2 = _prost((lx, ly), roz_l)
                if not _box_w_srodku(_rozszerz(lb2, margines_litery), e["pts"]):
                    continue
                if _przecina(lb2):
                    continue
                zle = False
                for e2 in elementy:
                    if _nakladka(_rozszerz(e2["prost"], margines_litery), lb2) > 0:
                        zle = True
                        break
                if zle:
                    continue
                for jj in range(len(litery)):
                    if jj != li_i and _nakladka(litery[jj], lb2) > 0:
                        zle = True
                        break
                if zle:
                    continue
                zn = (lx, ly, lb2)
                break
            if zn:
                lx, ly, lb2 = zn
                litery[li_i] = lb2
                e["offset_litery"] = (lx - e["srodek"][0], ly - e["srodek"][1])
                poprawki += 1
            else:
                litery[li_i] = stara
        if poprawki == 0:
            break
    # ---- na koniec: WYSIĘGNIK nie może przecinać własnego opisu ----
    # (gdy opis wyszedł nad/pod literę, linia z końca kreski idzie przez
    #  wiersz tekstu; tu szukamy innego miejsca poza wydzieleniem, przy którym
    #  wysięgnik jest czysty). Robimy to PO ułożeniu liter, bo ich ruch
    #  zmienia geometrię wysięgnika.
    for _ in range(3):
        poprawki = 0
        for el in elementy:
            if not el.get("wysiegnik"):
                continue
            li = el.get("lit_info")
            if not li:
                continue
            base = el["srodek"]
            roz = el["rozmiar"]
            pts = el["pts"]
            li_i, ol, roz_l = li
            akt = el.get("offset_litery") or ol
            lit_c = (base[0] + akt[0], base[1] + akt[1])

            def _przec(dx_, dy_):
                opis_c = (base[0] + dx_, base[1] + dy_)
                pbox = _prost(opis_c, roz)
                kon_ = _koniec_wys(lit_c, roz_l, opis_c)
                return min(
                    _dlugosc_odc_w_prost((pbox[0], opis_c[1]), kon_, pbox),
                    _dlugosc_odc_w_prost((pbox[2], opis_c[1]), kon_, pbox))

            cur_dx, cur_dy = el["offset"]
            cur_prz = _przec(cur_dx, cur_dy)
            if cur_prz <= 0.3:
                continue
            best = None
            for r in range(1, 15):
                for k in range(48):
                    a = k * math.pi / 24
                    dxx = math.cos(a) * (25.0 + r * 6.0)
                    dyy = math.sin(a) * (25.0 + r * 6.0)
                    p2 = _prost((base[0] + dxx, base[1] + dyy), roz)
                    if _przecina(p2):
                        continue
                    if _box_w_srodku(p2, pts):
                        continue          # to ma być poza wydzieleniem
                    zle = False
                    for e2 in elementy:
                        if e2 is el:
                            continue
                        if _nakladka(p2, e2["prost"]) > 0:
                            zle = True
                            break
                    if zle:
                        continue
                    for jj in range(len(litery)):
                        if jj != li_i and _nakladka(p2, litery[jj]) > 0:
                            zle = True
                            break
                    if zle:
                        continue
                    pr = _przec(dxx, dyy)
                    if pr < 0.3:
                        best = (dxx, dyy, p2, pr)
                        break
                    if best is None or pr < best[3]:
                        best = (dxx, dyy, p2, pr)
                if best is not None and best[3] < 0.3:
                    break
            if best is not None and best[3] < cur_prz - 0.2:
                el["offset"] = (best[0], best[1])
                el["prost"] = best[2]
                poprawki += 1
        if not poprawki:
            break


    return elementy
