"""
Forestly — moduł WYDRUKI

Generuje wszystkie pliki wydrukowe MIETEKA bezpośrednio z plików DBF mietka
(bez uruchamiania MS-DOS):

  * HALIZNY.TXT  — zestawienie powierzchni niezalesionych (PRZED przeniesieniem)
  * OPTAX.TXT    — opis lasów i gruntów (PO przeniesieniu halizn)
  * TAB_KLW3.TXT — zestawienie wg klas i podklas wieku (PO przeniesieniu)
  * ZEST1.TXT    — skorowidz działek
  * REJESTR1.TXT — rejestr działek wg właścicieli (PO przeniesieniu)
  * WSKAZ1.TXT   — wykaz wskaźników (pusty, gdy brak danych)

Formaty odtworzone 1:1 z wydruków MIETEKA (cp852, CRLF, ramki, sekwencje PCL).
"""

import re
import struct
from pathlib import Path


# ----------------------------------------------------------------------------
# Czytanie DBF / NTX
# ----------------------------------------------------------------------------

def czytaj_dbf(path):
    """Czyta DBF. Zwraca listę słowników {pole: wartość} (cp852)."""
    data = Path(path).read_bytes()
    nrec = struct.unpack('<I', data[4:8])[0]
    hlen = struct.unpack('<H', data[8:10])[0]
    rlen = struct.unpack('<H', data[10:12])[0]
    fields = []
    off = 32
    while off < len(data) and data[off:off + 1] != b'\x0d':
        raw = data[off:off + 32]
        fields.append((
            raw[0:11].split(b'\x00')[0].decode('ascii', errors='replace'),
            chr(raw[11]), raw[16], raw[17],
        ))
        off += 32
    recs = []
    pos = hlen
    for _ in range(nrec):
        row = data[pos:pos + rlen]
        pos += rlen
        if not row or len(row) < rlen or row[0:1] == b'*':
            continue
        rec = {}
        o = 1
        for fname, ftype, flen, fdec in fields:
            raw = row[o:o + flen].replace(b'\x00', b' ').strip()
            o += flen
            s = raw.decode('cp852', errors='replace')
            if ftype == 'N':
                s2 = s.replace(',', '.').strip()
                try:
                    rec[fname] = float(s2) if ('.' in s2 or fdec) else int(s2)
                except ValueError:
                    rec[fname] = None
            elif ftype == 'L':
                rec[fname] = (s == 'T') if s else None
            else:
                rec[fname] = s
        recs.append(rec)
    return recs


def znajdz_dbf(katalog, prefiks):
    """Znajduje (rekurencyjnie) pierwszy plik DBF o danym prefiksie nazwy."""
    katalog = Path(katalog)
    trafienia, seen = [], set()
    for w in (f'{prefiks}*.DBF', f'{prefiks}*.dbf'):
        for p in sorted(katalog.rglob(w)):
            if str(p).upper() not in seen:
                seen.add(str(p).upper())
                trafienia.append(p)
    return trafienia[0] if trafienia else None


def czytaj_ntx_kolejnosc(path_ntx, n_rekordow):
    """Czyta kolejność rekordów z jednopoziomowego indeksu NTX (jak MIETEK).

    Zwraca listę numerów rekordów (1-based) w kolejności indeksu,
    albo None gdy pliku nie ma / nie da się sparsować.
    """
    try:
        d = Path(path_ntx).read_bytes()
    except OSError:
        return None
    # strony 512-bajtowe; szukamy strony-liścia z poprawnymi wpisami
    best = None
    for base in range(512, len(d), 512):
        for pgstart in (base, base + 256 if base + 256 + 32 <= len(d) else -1):
            if pgstart < 0:
                continue
            pg = d[pgstart:pgstart + 1024]
            if len(pg) < 64:
                continue
            try:
                n = struct.unpack('<h', pg[0:2])[0]
            except struct.error:
                continue
            if n <= 0 or n > n_rekordow or 2 + 2 * n >= len(pg):
                continue
            offs = [struct.unpack('<h', pg[2 + 2 * i:4 + 2 * i])[0] for i in range(n)]
            if not offs:
                continue
            stride = offs[1] - offs[0] if len(offs) > 1 else 17
            if stride <= 8 or any(o <= 8 or o + stride > len(pg) for o in offs):
                continue
            recs = []
            ok = True
            for O in offs:
                try:
                    rec = struct.unpack('<i', pg[O + 4:O + 8])[0]
                except struct.error:
                    ok = False
                    break
                if not (1 <= rec <= n_rekordow):
                    ok = False
                    break
                recs.append(rec)
            if ok and len(set(recs)) == len(recs):
                if best is None or len(recs) > len(best):
                    best = recs
    return best


# ----------------------------------------------------------------------------
# DATA.CFG / WSIE.DBF — agencja, obiekt, stan na
# ----------------------------------------------------------------------------

def czytaj_agencje(obreb_dir):
    katalog = Path(obreb_dir)
    cfg = next((c for c in katalog.rglob('DATA.CFG')), None) \
        or next((c for c in katalog.rglob('DATA.cfg')), None)
    if cfg is None:
        return None
    try:
        t = cfg.read_bytes().decode('cp852', errors='replace')
        m = re.search(r'\d{2}-\d{3}', t)
        if not m:
            return None
        nazwa = re.sub(r'^[\d\s]+', '', t[:m.start()]).strip()
        return nazwa or None
    except Exception:
        return None


def czytaj_dane_wsi(obreb_dir):
    wsie_path = znajdz_dbf(obreb_dir, 'WSIE')
    if wsie_path is None:
        return '', ''
    try:
        recs = czytaj_dbf(wsie_path)
        if not recs:
            return '', ''
        nazwa = str(recs[0].get('NAZWA', '') or '').strip()
        stan = str(recs[0].get('STAN_NA', '') or '').strip()
        if re.match(r'^\d{8}$', stan):
            stan = f"{stan[6:8]}-{stan[4:6]}-{stan[0:4]}"
        return nazwa, stan
    except Exception:
        return '', ''


# ----------------------------------------------------------------------------
# Wspólne pomocniki
# ----------------------------------------------------------------------------

_STL = {11: 'Bs', 12: 'Bśw', 13: 'Bw', 14: 'Bb', 21: 'BM', 22: 'BMśw', 23: 'BMw', 24: 'BMb',
        31: 'LM', 32: 'LMśw', 33: 'LMw', 34: 'LMb', 42: 'Lśw', 43: 'Lw',
        44: 'Ol', 45: 'OlJ', 46: 'Lł'}
_PROG_NIEZALESIONE = 240


def _f4(v):
    if v is None:
        return ''
    try:
        return f"{float(v):.4f}"
    except (TypeError, ValueError):
        return '0.0000'


def _wsk_kod(wsk, miaz):
    """Kod wskaźnika w wydrukach (jak MIETEK): 'TP' z MIAZ 22 -> 'TP-22m3/ha'.

    Ręcznie wpisany kod z końcówką '-..m3/ha' jest najpierw odzierany z niej,
    potem doklejana końcówka wg aktualnego MIAZ (CS z MIAZ 0 -> 'CS').
    """
    base = re.sub(r'-\d+m3/ha$', '', wsk)
    if miaz and miaz > 0 and 'm3' not in base:
        return f"{base}-{int(miaz)}m3/ha"
    return base


def _wsk_rj_kod(wsk, miaz):
    """Kod wskaźnika w REJESTR1: kod z myślnikiem jest ucinany do części
    przed myślnikiem; czysty kod z MIAZ>0 dostaje końcówkę '-..m3/ha'."""
    if '-' in wsk:
        return wsk.split('-')[0]
    if miaz and miaz > 0:
        return f"{wsk}-{int(miaz)}m3/ha"
    return wsk


def _klucz_wydz(r):
    oddz = str(r.get('ODDZIAL', '') or '').strip()
    poddz = str(r.get('PODODDZ', '') or '').strip()
    try:
        return (int(float(oddz)), oddz, poddz)
    except ValueError:
        return (9999, oddz, poddz)


def _wczytaj_obreb(obreb_dir):
    """Wczytuje W, D, O, R + metadane obrębu."""
    obreb = Path(obreb_dir)
    w_path = znajdz_dbf(obreb, 'W')
    d_path = znajdz_dbf(obreb, 'D')
    o_path = znajdz_dbf(obreb, 'O')
    r_path = znajdz_dbf(obreb, 'R')
    if o_path is None:
        return None
    dane = {
        'W': czytaj_dbf(w_path) if w_path else [],
        'D': czytaj_dbf(d_path) if d_path else [],
        'O': czytaj_dbf(o_path),
        'R': czytaj_dbf(r_path) if r_path else [],
        'w_path': w_path, 'd_path': d_path, 'o_path': o_path, 'r_path': r_path,
    }
    # indeksy: oddz+poddz -> rekord
    dane['O_by'] = {f"{r.get('ODDZIAL', '')}{r.get('PODODDZ', '')}".strip(): r for r in dane['O']}
    dane['R_by'] = {f"{r.get('ODDZIAL', '')}{r.get('PODODDZ', '')}".strip(): r for r in dane['R']}
    dane['R_all'] = {}
    for _r in dane['R']:
        dane['R_all'].setdefault(
            f"{_r.get('ODDZIAL', '')}{_r.get('PODODDZ', '')}".strip(), []).append(_r)
    return dane


def _wsks(r):
    """Lista (wsk, pow_wsk, miaz) dla R; pomija puste."""
    out = []
    for i in range(1, 7):
        wsk = str(r.get(f'WSK{i}', '') or '').strip()
        pw = r.get(f'POW_WSK{i}', 0) or 0
        mz = r.get(f'MIAZ{i}', 0) or 0
        if wsk:
            out.append((wsk, float(pw), mz))
    return out


def _miaz_na_ha(wsk, miaz, zasob):
    """Miąższość na hektar dla wskaźnika (0 = brak)."""
    if miaz and miaz > 0:
        return float(miaz)
    if wsk.startswith('Rb'):
        return float(zasob or 0)
    return 0.0


def _zapisz(path, lines, tail='\r\n'):
    Path(path).write_bytes(('\r\n'.join(lines) + tail).encode('cp852'))


# ----------------------------------------------------------------------------
# HALIZNY.TXT — Zestawienie powierzchni leśnych niezalesionych
# ----------------------------------------------------------------------------

_PCL_HAL = '\r\x1b(s16.67H\x1b&l5E\x1b&a14L'
_HAL_TOP = '┌───────┬─────────┬─────────────────────────────┐'
_HAL_H1 = '│       │         │                             │'
_HAL_H2 = '│Oddział│   Pow.  │           Rodzaj            │ '
_HAL_H3 = '│poddz. │   wydz. │         powierzchni         │'
_HAL_H4 = '│       │   [ha]  │                             │'
_HAL_SEP = '├───────┼─────────┼─────────────────────────────┤'
_HAL_NUM = '│   1   │    2    │              3              │'
_HAL_BOT = '└───────┴─────────┴─────────────────────────────┘'
_HAL_TITLE = ' Zestawienie powierzchni leśnych niezalesionych'


def generuj_halizny_txt(obreb_dir, agencja=None):
    """HALIZNY.TXT (przed przeniesieniem halizn w D*.DBF)."""
    obreb = Path(obreb_dir)
    o_dbf = znajdz_dbf(obreb, 'O')
    if o_dbf is None:
        return None, 0
    wydzielenia = czytaj_dbf(o_dbf)
    if agencja is None:
        agencja = czytaj_agencje(obreb) or ''
    obiekt, stan_na = czytaj_dane_wsi(obreb)

    niezalesione = sorted(
        (r for r in wydzielenia
         if (r.get('RODZ_POW') or 0) and (r.get('RODZ_POW') or 0) >= _PROG_NIEZALESIONE),
        key=_klucz_wydz,
    )
    if not niezalesione:
        return None, 0

    lines = [_PCL_HAL + agencja.ljust(114) + f"Strona {1:4d}",
             _HAL_TITLE.ljust(59) + f"Obiekt: {obiekt}".ljust(49) + f"Stan na: {stan_na}",
             _HAL_TOP, _HAL_H1, _HAL_H2, _HAL_H3, _HAL_H4,
             _HAL_SEP, _HAL_NUM, _HAL_SEP]

    suma_calk, biez_oddz, suma_oddz = 0.0, None, 0.0

    def zamknij():
        lines.append(_HAL_SEP)
        lines.append('│R.oddz.│' + _f4(suma_oddz).rjust(9) + '│' + ' ' * 29 + '│')
        lines.append(_HAL_SEP)

    for r in niezalesione:
        oddz = str(r.get('ODDZIAL', '') or '').strip()
        poddz = str(r.get('PODODDZ', '') or '').strip()
        if biez_oddz is not None and oddz != biez_oddz:
            zamknij()
            suma_calk += suma_oddz
            suma_oddz = 0.0
        opis = f"{int(r.get('RODZ_POW') or 0)}-{str(r.get('OP_TAX', '') or '')[:30].strip()}"[:29]
        lines.append('│' + f"{oddz}{poddz}".rjust(7) + '│' +
                     _f4(r.get('POW_WYDZ') or 0).rjust(9) + '│' +
                     opis.ljust(29) + '│')
        biez_oddz = oddz
        suma_oddz += float(r.get('POW_WYDZ') or 0)

    zamknij()
    suma_calk += suma_oddz
    lines.append('│ Razem │' + _f4(suma_calk).rjust(9) + '│' + ' ' * 29 + '│')
    lines.append(_HAL_BOT)
    out = o_dbf.parent / 'HALIZNY.TXT'
    _zapisz(out, lines)
    return out, len(niezalesione)


# ----------------------------------------------------------------------------
# Pozycje rejestru (kolejność wg nazwisk) — używana przez REJESTR1, ZEST1, OPTAX
# ----------------------------------------------------------------------------

# kolacja MIETEKA: polskie diakrytyki foldowane do litery bazowej (Ą->A, Ó->O),
# natomiast Ł jako osobna litera tuż za L (ŁOKIETEK > LUDWINIAK, ale < MASTALERZ);
# pozostałe znaki (spacje, cyfry, ukośniki, myślniki) wg kodów ASCII
_PL_FOLD = str.maketrans('ąćęńóśźżĄĆĘŃÓŚŹŻ', 'acenoszzACENOSZZ')


def _pl_klucz(s):
    out = []
    for ch in str(s).upper():
        if ch == 'Ł':
            out.append(('L', 1))
        else:
            out.append((ch.translate(_PL_FOLD), 0))
    return out


def _pozycje_rejestru(dane):
    """Zwraca listę pozycji [(nrrej, [W...], [D...])] w kolejności MIETEKA.

    Pozycje numerowane są wg pierwszego wystąpienia w zbiorze W posortowanym
    po NAZWISKO+IMIE (kolacja _pl_klucz, sortowanie stabilne) — jak indeks
    NAZWISKO MIETEKA; właściciele pozycji w fizycznej kolejności W.
    """
    grupy = {}
    for w in dane['W']:
        nr = w.get('NRREJ')
        if nr is None:
            continue
        grupy.setdefault(int(nr), []).append(w)
    # stabilny sort W po kluczu -> pozycje w kolejności pierwszego wystąpienia
    porzadek = []
    seen = set()
    for w in sorted(dane['W'], key=lambda w: _pl_klucz(
            str(w.get('NAZWISKO', '') or '') + ' ' + str(w.get('IMIE', '') or ''))):
        nr = w.get('NRREJ')
        if nr is None:
            continue
        nr = int(nr)
        if nr not in seen:
            seen.add(nr)
            porzadek.append(nr)
    pozycje = []
    for nr in porzadek:
        ws = grupy[nr]          # fizyczna kolejność W w ramach pozycji
        ds = [d for d in dane['D'] if int(d.get('NRREJ') or 0) == nr]
        pozycje.append((nr, ws, ds))
    return pozycje


def _lp_pozycji(pozycje):
    return {nr: i + 1 for i, (nr, _, _) in enumerate(pozycje)}


# ----------------------------------------------------------------------------
# ZEST1.TXT — Skorowidz działek
# ----------------------------------------------------------------------------

_ZS_H = ('|---------|-------------------|----------|-------------|',
         '|nr dział.|   nr rejestru     | oddz/pod | pow. dział  |')


def _klucz_dzialki(dz):
    m = re.match(r'^(\d+)(.*)$', dz.strip())
    if not m:
        return (9999, dz.strip(), '')
    return (int(m.group(1)), m.group(2), '')


def generuj_zest1_txt(obreb_dir, dane=None, pozycje=None, agencja=None):
    """ZEST1.TXT — skorowidz działek (sortowanie numeryczne działek)."""
    obreb = Path(obreb_dir)
    if dane is None:
        dane = _wczytaj_obreb(obreb)
    if dane is None or not dane['D']:
        return None
    if pozycje is None:
        pozycje = _pozycje_rejestru(dane)
    lp_map = _lp_pozycji(pozycje)
    obiekt, stan = czytaj_dane_wsi(obreb)

    # kolejność w obrębie działki i właściciela: permutacja MIETEKA —
    # dolna połowa grupy rekordów (fizycznych) odwrócona i wyprowadzona
    # przed górną (artefakt wewnętrznego indeksu B-drzewa;
    # weryfikowane 1:1 na CHORZEWIE i WOL; przybliżenie dla grup >=6)
    grupy = {}
    for d in dane['D']:
        dz = str(d.get('NR_DZIAL', '') or '').strip()
        grupy.setdefault(dz, []).append(d)

    def _perm(g):
        n = len(g)
        k = n - n // 2          # rozmiar górnej połowy
        return g[k:][::-1] + g[:k]

    recs = []
    for dz in sorted(grupy, key=_klucz_dzialki):
        # w ramach działki: podgrupy wg właściciela (NRREJ) w kolejności fizycznej
        podgr = {}
        for d in grupy[dz]:
            podgr.setdefault(int(d.get('NRREJ') or 0), []).append(d)
        for nr in podgr:
            recs.extend(_perm(podgr[nr]))

    lines = ['\r Skorowidz działek'.ljust(35) + f"Obiekt: {obiekt}".ljust(49) +
             f"Stan na: {stan}", '',
             _ZS_H[0], _ZS_H[1], _ZS_H[0]]
    for d in recs:
        dz = str(d.get('NR_DZIAL', '') or '').strip()
        nr = int(d.get('NRREJ') or 0)
        lp = lp_map.get(nr, 0)
        oddzp = f"{str(d.get('ODDZIAL', '') or '').strip()}{str(d.get('PODODDZ', '') or '').strip()}"
        lines.append('|' + dz.ljust(9) + '|' + f"{lp}/{nr}".rjust(18) + ' ' + '|' +
                     oddzp.center(10) + '|' + _f4(d.get('POW') or 0).rjust(13) + '|' +
                     ' ' * 45 + '|')
    lines.append(_ZS_H[0])
    out = (dane['d_path'].parent if dane['d_path'] else obreb) / 'ZEST1.TXT'
    _zapisz(out, lines)
    return out


def _idx(lst, item):
    for i, x in enumerate(lst):
        if x is item:
            return i
    return 0


# ----------------------------------------------------------------------------
# OPTAX.TXT — Opis lasów i gruntów przeznaczonych do zalesienia
# ----------------------------------------------------------------------------

_OP_PCL = '\x1b(s16.67H\x1b&l5E\x1b&a9L'
_OP_H = (
    '┌───────┬─────────┬───────────────────────────────────┬──────────────────────────────────────┬──────────────────────────┬───────────────────────┐',
    '│       │         │ Opis taksacyjny lasu,             │          Elementy taksacyjne         │   Wskazania gospodarcze  │       Wykonanie       │',
    '│Oddział│  Pow.   │ gruntu przeznaczonego             ├────├────┬─────┬───┬────┬───┬───┬─────┼──────────┬─────────┬─────┼─────────┬───────┬─────┤',
    '│poddz. │  [ha]   │ do zalesienia                     │Gat.│Wiek│Klasa│Wys│Pier│Bon│Zad│Miąż.│Rodzaj    │  Pow.   │Maks.│  Czynn. │ [ha]  │[m3] │',
    '│       │         │                                   │ gł.│    │wieku│[m]│[cm]│   │   │ na  │wskazania │  [ha]   │miąż.│         │       │     │',
    '│       │         │                                   │    │    │     │   │    │   │   │ pow.│          │         │ do  │         │       │     │',
    '│       │         │                                   │    │    │     │   │    │   │   │ [m3]│          │         │pozys│         │       │     │',
    '│       │         │                                   │    │    │     │   │    │   │   │     │          │         │ [m3]│         │       │     │',
    '├───────┼─────────┼───────────────────────────────────┼────┼────┼─────┼───┼────┼───┼───┼─────┼──────────┼─────────┼─────┼─────────┼───────┼─────┤',
    '│   1   │    2    │                3                  │ 4  │ 5  │  6  │ 7 │  8 │ 9 │ 10│ 11  │    12    │    13   │  14 │   15    │  16   │ 17  │',
    '├───────┼─────────┼───────────────────────────────────┼────┼────┼─────┼───┼────┼───┼───┼─────┼──────────┼─────────┼─────┼─────────┼───────┼─────┤',
)
_OP_BOT = '└───────┴─────────┴───────────────────────────────────┴────┴────┴─────┴───┴────┴───┴───┴─────┴──────────┴─────────┴─────┴─────────┴───────┴─────┘'


def _seg_optax(row1, desc, wsk, pow_wydz=0.0, tail=None):
    """Wiersz OPTAX.

    row1 – None (wiersz kontynuacji) albo 12-elementowa krotka:
           (oddzp, pow_str, gat, wiek, klasa, wys, pier, bon, zad, miaz_pow)
    desc  – tekst opisu (kolumna 3)
    wsk   – None albo (wskaznik, pow_wsk_proc, miaz_na_ha)
    tail  – elementy taksacyjne na wierszu kontynuacji (drugi rekord R,
            np. taksacja płazowiny) – 8-elementowa lista
    """
    def c(v, w, just='>'):
        v = '' if v is None else str(v)
        return v.ljust(w) if just == '<' else v.rjust(w)

    if row1 is None and tail:
        t = ['' if v is None else str(v) for v in tail]
        s = ('│' + ' ' * 7 + '│' + ' ' * 9 + '│' + c(desc, 35, '<') +
             '│' + c(t[0], 4, '<') + '│' + c(t[1], 4) + '│' + ' ' + c(t[2], 4, '<') +
             '│' + c(t[3], 3) + '│' + c(t[4], 4) + '│' + c(t[5], 3, '<') +
             '│' + c(t[6], 3, '<') + '│' + c(t[7], 5))
    elif row1 is None:
        s = ('│' + ' ' * 7 + '│' + ' ' * 9 + '│' + c(desc, 35, '<') +
             '│    │    │     │   │    │   │   │     ')
    else:
        r = row1
        s = ('│' + c(r[0], 7) + '│' + c(r[1], 9) + '│' + c(desc, 35, '<') +
             '│' + c(r[2], 4, '<') + '│' + c(r[3], 4) + '│' + ' ' + c(r[4], 4, '<') +
             '│' + c(r[5], 3) + '│' + c(r[6], 4) + '│' + c(r[7], 3, '<') +
             '│' + c(r[8], 3, '<') + '│' + c(r[9], 5))
    if wsk is None:
        s += ('│' + ' ' * 10 + '│' + ' ' * 9 + '│' + ' ' * 5 +
              '│' + ' ' * 9 + '│' + ' ' * 7 + '│' + ' ' * 5 + '│')
    else:
        w, pw, mh = wsk
        if pw > 0:
            pow_w = pow_wydz * pw / 100.0   # kolejność jak MIETEK (arytmetyka binarna!)
            maks = round(mh * pow_w) if mh > 0 else ''
            if not maks:
                maks = ''           # wyrażone w całości < 0,5 m3 - jak MIETEK
        else:                      # wskaźnik bez udziału powierzchni
            # przestoje: miążdżość podana w m3 całkowitych
            pow_w, maks = None, (str(int(mh)) if mh > 0 else '')
        s += ('│' + c(w, 10, '<') + '│' + c(_f4(pow_w), 9) + '│' + c(maks, 5) +
              '│' + ' ' * 9 + '│' + ' ' * 7 + '│' + ' ' * 5 + '│')
    return s


def generuj_optax_txt(obreb_dir, dane=None, pozycje=None, agencja=None):
    """OPTAX.TXT — opis lasów i gruntów (po przeniesieniu halizn)."""
    obreb = Path(obreb_dir)
    if dane is None:
        dane = _wczytaj_obreb(obreb)
    if dane is None:
        return None
    if pozycje is None:
        pozycje = _pozycje_rejestru(dane)
    lp_map = _lp_pozycji(pozycje)
    if agencja is None:
        agencja = czytaj_agencje(obreb) or ''
    obiekt, stan = czytaj_dane_wsi(obreb)

    def _klucz_optax(o):
        oddz = str(o.get('ODDZIAL', '') or '').strip()
        poddz = str(o.get('PODODDZ', '') or '').strip()
        # warianty z sufiksem 'x' (np. 1ax) MIETEK wyprowadza po wszystkich
        # normalnych pododdziałach danego oddziału
        wariant = 1 if len(poddz) > 1 and poddz.endswith('x') else 0
        try:
            return (int(float(oddz)), wariant, poddz, oddz)
        except ValueError:
            return (9999, wariant, poddz, oddz)

    wydzs = sorted(dane['O'], key=_klucz_optax)

    # bloki: (wiersze) per wydzielenie; na końcu każdego oddziału wiersz R.oddz.
    bloki = []
    suma_pow = 0.0
    suma_miaz = 0
    suma_maks = 0
    oddz_pow, oddz_miaz, oddz_maks = 0.0, 0, 0
    cur_oddz = None

    def _tot_row(name, p, mz, mk):
        return ('│' + f"{name}".ljust(7) + '│' + _f4(p).rjust(9) + '│' + ' ' * 35 +
                '│    │    │     │   │    │   │   │' + str(int(mz)).rjust(5) + '│' +
                ' ' * 10 + '│' + ' ' * 9 + '│' + str(int(mk)).rjust(5) +
                '│' + ' ' * 9 + '│' + ' ' * 7 + '│' + ' ' * 5 + '│')

    def _koniec_oddzialu():
        if cur_oddz is None:
            return
        bloki.append([_tot_row('R.oddz.', oddz_pow, oddz_miaz, oddz_maks)])

    for o in wydzs:
        key = f"{o.get('ODDZIAL', '')}{o.get('PODODDZ', '')}".strip()
        oddz = str(o.get('ODDZIAL', '') or '').strip()
        if cur_oddz is not None and oddz != cur_oddz:
            _koniec_oddzialu()
            oddz_pow, oddz_miaz, oddz_maks = 0.0, 0, 0
        cur_oddz = oddz
        r = dane['R_by'].get(key, {})
        pow_w = float(o.get('POW_WYDZ') or 0)
        zasob = float(r.get('ZASOB') or 0)
        suma_pow += pow_w
        suma_miaz += round(zasob * pow_w)
        oddz_pow += pow_w
        oddz_miaz += round(zasob * pow_w)
        # opis: siedlisko + gatunki, potem segmenty OP_TAX (po 35 znaków)
        stl = _STL.get(int(o.get('TYP_SIED') or 0), '')
        op_t = str(o.get('OP_TAX', '') or '')
        op_t1 = str(o.get('OP_TAX1', '') or '')
        segs = [op_t[i * 35:(i + 1) * 35].strip() for i in range(7)]
        segs += [op_t1[i * 35:(i + 1) * 35].strip() for i in range(7)]
        segs = [s for s in segs if s]
        # nr-y rejestrów właścicieli tego wydzielenia (wg lp pozycji)
        tmp = {}
        for d in dane['D']:
            if f"{d.get('ODDZIAL', '')}{d.get('PODODDZ', '')}".strip() == key:
                nr = int(d.get('NRREJ') or 0)
                tmp.setdefault(nr, lp_map.get(nr, 0))
        lp_strs = [f"{lp}/{nr}" for nr, lp in sorted(tmp.items(), key=lambda kv: kv[1])]
        # nr-y rejestrów zawijane po elementach do 35 znaków (jak MIETEK)
        nry_rows = []
        if lp_strs:
            line = 'nr-y.Rej. '
            for it in lp_strs:
                item = it + ','
                if len(line) + len(item) > 35:
                    nry_rows.append(line)
                    line = ''
                line += item
            if line:
                nry_rows.append(line)
        desc = [f"{stl:<7}{str(o.get('GTD1', '') or ''):<4}{str(o.get('GTD2', '') or ''):<4}"
                f"{str(o.get('GTD3', '') or '')}".rstrip()] + segs + nry_rows

        # elementy taksacyjne; wydzielenie może mieć DWA rekordy R:
        # warstwę wskaźnikową (kl_wiek='') i taksację (z kl_wiek) — wtedy
        # wiersz główny z warstwy wskaźnikowej, a taksacja wiersz dalej
        rs = dane.get('R_all', {}).get(key, [])
        r_prz = next((x for x in rs if not str(x.get('KL_WIEK', '') or '').strip()), None)
        r_tak = next((x for x in rs if str(x.get('KL_WIEK', '') or '').strip()), None)
        dual = r_prz is not None and r_tak is not None
        r_main = r_prz if dual else r
        zasob_m = float(r_main.get('ZASOB') or 0)

        wiek = r_main.get('WIEK')
        miaz_pow = round(zasob_m * pow_w)
        if r_main.get('PRZES'):
            miaz_pow = round(zasob_m)
        row1_tail = [
            str(r_main.get('GATUNEK', '') or '').strip(),          # 4 Gat. gł
            str(wiek) if wiek else '',                        # 5 Wiek
            str(r_main.get('KL_WIEK', '') or '').strip(),         # 6 Klasa
            str(r_main.get('WYS') or '') if r_main.get('WYS') else '',  # 7 Wys
            str(r_main.get('PIERS') or '') if r_main.get('PIERS') else '',  # 8 Pier
            str(r_main.get('BONIT', '') or '').strip(),            # 9 Bon
            f"{r_main.get('ZADRZEW', 0) or 0:.1f}" if wiek else '',  # 10 Zad
            str(miaz_pow) if wiek else '',                  # 11 Miąż na pow
        ]

        # taksacja drugiego rekordu (wiersz po ostatnim wskaźniku)
        tail2 = None
        if dual:
            w2 = r_tak.get('WIEK')
            m2 = round(float(r_tak.get('ZASOB') or 0) * pow_w)
            tail2 = [
                str(r_tak.get('GATUNEK', '') or '').strip(),
                str(w2) if w2 else '',
                str(r_tak.get('KL_WIEK', '') or '').strip(),
                str(r_tak.get('WYS') or '') if r_tak.get('WYS') else '',
                str(r_tak.get('PIERS') or '') if r_tak.get('PIERS') else '',
                str(r_tak.get('BONIT', '') or '').strip(),
                f"{r_tak.get('ZADRZEW', 0) or 0:.1f}" if w2 else '',
                str(m2) if w2 else '',
            ]

        wsks = [(_wsk_kod(w, mz) if pw > 0 else re.sub(r'-\d+m3/ha$', '', w),
                 pw, _miaz_na_ha(w, mz, zasob)) for w, pw, mz in _wsks(r_main)]
        for w, pw, mh in wsks:
            if mh > 0:
                if pw > 0:
                    mk = round(mh * (pw / 100.0) * pow_w)
                else:              # przestój: m3 całkowite
                    mk = int(mh)
                suma_maks += mk
                oddz_maks += mk
                break

        # zbuduj wiersze bloku
        rows = []
        wsk_i = 0
        for di, dline in enumerate(desc):
            wsk = None
            if wsk_i < len(wsks):
                wsk = wsks[wsk_i]
                wsk_i += 1
            if di == 0:
                rows.append(_seg_optax(
                    (f"{o.get('ODDZIAL', '')}{o.get('PODODDZ', '')}".strip(),
                     _f4(pow_w)) + tuple(row1_tail), dline, wsk, pow_w))
            else:
                rows.append(_seg_optax(None, dline, wsk, pow_w))
        # dodatkowe wskaźniki bez linii opisu
        while wsk_i < len(wsks):
            rows.append(_seg_optax(None, '', wsks[wsk_i], pow_w))
            wsk_i += 1
        # taksacja drugiego rekordu R
        if dual:
            pos = max(len(wsks), 1)
            if pos < len(rows):
                # wiersz opisu na tej pozycji — przerenderuj z elementami
                dline = desc[pos] if pos < len(desc) else ''
                rows[pos] = _seg_optax(None, dline, None, pow_w, tail2)
            else:
                rows.append(_seg_optax(None, '', None, pow_w, tail2))
        bloki.append(rows)

    # --- składanie stron: budżet 27 linii danych (bloki + separatory),
    # blok wydzielenia nigdy nie jest dzielony między strony; końcowa stopka
    # (R.oddz. ostatniego oddziału + Razem) nie podlega paginacji ---
    lines = []
    pages, cur, used = [], [], 0

    def eject():
        nonlocal cur, used
        if cur:
            pages.append(cur)
        cur, used = [], 0

    for rows in bloki:
        add = len(rows) + (1 if used > 0 else 0)
        if used > 0 and used + add > 27:
            eject()
            add = len(rows)
        if used:
            cur.append(_OP_H[-1])
        cur.extend(rows)
        used += add

    # stopka końcowa na ostatniej stronie (bez limitu — tak MIETEK)
    if bloki:
        if cur:
            cur.append(_OP_H[-1])
        cur.append(_tot_row('R.oddz.', oddz_pow, oddz_miaz, oddz_maks))
        cur.append(_OP_H[-1])
        cur.append(_tot_row(' Razem ', suma_pow, suma_miaz, suma_maks))
    if cur:
        pages.append(cur)

    _bot_last = _OP_BOT[:59] + '┼' + _OP_BOT[60:]
    pageno = 1
    all_pages = pages
    for pg in all_pages:
        is_last = pg is all_pages[-1]
        if is_last:
            # nagłówek ostatniej strony: 14L gdy strona się mieści, czysty przy przepełnieniu
            nl = 2 + len(_OP_H) + len(pg) + 3
            pcl = _OP_PCL.replace('&a9L', '&a14L') if nl <= 42 else ''
        else:
            pcl = _OP_PCL
        hdr = ('\r' if pageno == 1 else '\f\r') + pcl + agencja.ljust(114) + f"Strona {pageno:4d}"
        lines.append(hdr)
        lines.append(' Opis lasów i gruntów przeznaczonych do zalesienia'.ljust(56) +
                     f"Obiekt: {obiekt}".ljust(49) + f"Stan na: {stan}")
        lines.extend(_OP_H)
        lines.extend(pg)
        lines.append(_bot_last if is_last else _OP_BOT)
        pageno += 1
    lines.append('\f')

    if not all_pages:
        return None, 0
    out = dane['o_path'].parent / 'OPTAX.TXT'
    _zapisz(out, lines, tail='')
    return out, pageno - 1


# ----------------------------------------------------------------------------
# TAB_KLW3.TXT — Zestawienie wg klas i podklas wieku (+ siedliska, ochrona, przebudowa)
# ----------------------------------------------------------------------------

_TK_PCL = '\r\x1b(s16.67H\x1b&l5E\x1b&a0L'
_TK_H = (
    '┌───────┬──────┬───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐',
    '│Gatunek│Przes.│                                                     Powierzchnia w ha / miąższośc w m3                                                                                    │',
    '│panuj. │i nas.├─────────┬───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┬─────────┬─────────┬──────────┬──────────┤',
    '│       │      │  Pow.   │                                                 Klasy, podklasy wieku                                                 │         │         │          │          │',
    '│       │      │ leśna   ├───────────────────┬───────────────────┬───────────────────┬───────────────────┬───────────────────┬─────────┬─────────┤         │         │   Razem  │  Ogółem  │',
    '│       │      │  nie    │         I         │        II         │        III        │        IV         │         V         │   VI    │   VII   │ K.D.O.  │   K.O.  │   kol.   │   kol.   │',
    '│       │      │zalesiona├─────────┬─────────┼─────────┬─────────┼─────────┬─────────┼─────────┬─────────┼─────────┬─────────┤         │    i    │         │         │   4-17   │   3-17   │',
    '│       │      │         │    a    │    b    │    a    │    b    │    a    │    b    │    a    │    b    │    a    │    b    │         │   wyż.  │         │         │          │          │',
    '├───────┼──────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼──────────┼──────────┤',
    '│   1   │  2   │    3    │    4    │    5    │    6    │    7    │    8    │    9    │   10    │   11    │   12    │   13    │   14    │   15    │   16    │   17    │    18    │    19    │',
    '├───────┼──────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼──────────┼──────────┤',
)
_TK_BOT = '└───────┴──────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴──────────┴──────────┘'
_TK_KLASY = ['Ia', 'Ib', 'IIa', 'IIb', 'IIIa', 'IIIb', 'IVa', 'IVb', 'Va', 'Vb', 'VI', 'VII',
             'KDO', 'KO']

_OCHR = {0: '0-brak ochronności', 1: '1-wod.', 2: '2-pow. gleb.', 3: '3-nieuż.roln.',
         4: '4-poz. przyg.', 5: '5-nauk.', 6: '6-pow. pojar.', 7: '7-uzdrow.',
         8: '8-stref. ost.', 9: '9-gosp. chełm.'}


def generuj_tab_klw3_txt(obreb_dir, dane=None, agencja=None, strona_start=1):
    # numer strony: MIETEK kontynuuje licznik OPTAX obcięty do 1 cyfry (5 i 75 -> '5')
    strona_start = (int(strona_start) % 10) or 10
    """TAB_KLW3.TXT (po przeniesieniu halizn)."""
    obreb = Path(obreb_dir)
    if dane is None:
        dane = _wczytaj_obreb(obreb)
    if dane is None:
        return None
    if agencja is None:
        agencja = czytaj_agencje(obreb) or ''
    obiekt, _ = czytaj_dane_wsi(obreb)

    # --- agregacja po gatunkach i klasach wieku ---
    # Kolejność gatunków: pierwsze wystąpienie w R wg oddziału/pododdziału
    # (jak indeks odpod MIETEKA). Wydzielenie może mieć DWA rekordy R:
    # warstwę przestojową (klw='', wsk zwykle 'Us.przest.') i taksację
    # (klw + ZASOB). Miążdżość liczymy per wydzielenie z zaokrągleniem
    # PO zsumowaniu powierzchni działek (weryfikowane na JAŹWI).
    gat_order = []
    for r in sorted(dane['R'], key=_klucz_wydz):
        g = str(r.get('GATUNEK', '') or '').strip()
        if g and g not in gat_order:
            gat_order.append(g)

    R_w = {}
    for r in dane['R']:
        if r.get('PRZES'):
            continue
        R_w.setdefault(f"{r.get('ODDZIAL', '')}{r.get('PODODDZ', '')}".strip(), []).append(r)
    dz_lz = {}
    for d in dane['D']:
        k = f"{d.get('ODDZIAL', '')}{d.get('PODODDZ', '')}".strip()
        lz, lnz = float(d.get('POW_L_ZAL') or 0), float(d.get('POW_L_NZAL') or 0)
        if k in dz_lz:
            dz_lz[k] = (dz_lz[k][0] + lz, dz_lz[k][1] + lnz)
        else:
            dz_lz[k] = (lz, lnz)

    gat_pow = {g: [0.0] * 15 for g in gat_order}
    gat_m3 = {g: [0] * 15 for g in gat_order}
    for key, recs in R_w.items():
        drz = next((r for r in recs if str(r.get('KL_WIEK', '') or '').strip()), None)
        prz = next((r for r in recs if not str(r.get('KL_WIEK', '') or '').strip()), None)
        lz, lnz = dz_lz.get(key, (0.0, 0.0))
        # powierzchnie niezalesione: gatunek warstwy przestojowej, a przy jej
        # braku - drzewostanu (działki częściowo niezalesione)
        g_n = str((prz or drz or {}).get('GATUNEK', '') or '').strip()
        if g_n in gat_pow:
            gat_pow[g_n][0] += lnz
            if drz is not None:
                # miążdżość przestojów: 2x round(zasob x lnz) - tak liczy MIETEK
                gat_m3[g_n][0] += 2 * round(float(drz.get('ZASOB') or 0) * lnz)
        if drz is not None:
            g_d = str(drz.get('GATUNEK', '') or '').strip()
            klw = str(drz.get('KL_WIEK', '') or '').strip()
            if g_d in gat_pow and klw in _TK_KLASY:
                i = 1 + _TK_KLASY.index(klw)
                gat_pow[g_d][i] += lz
                gat_m3[g_d][i] += round(float(drz.get('ZASOB') or 0) * lz)

    def wiersz(gat, m3):
        if m3:
            # komórki m3 zaokrąglone per gatunek; Razem sumuje zaokrąglone
            cells = [round(v) for v in gat_m3[gat]]
            razem = sum(cells[1:])
            ogolem = razem + cells[0]
            s = ('│' + ' ' * 7 + '│' + '0'.rjust(6) + '│' +
                 '│'.join(f"{v}".rjust(9) for v in cells) + '│' +
                 f"{razem}".rjust(10) + '│' + f"{ogolem}".rjust(10) + '│')
        else:
            cells = gat_pow[gat]
            razem = sum(cells[1:])
            ogolem = razem + cells[0]
            s = ('│' + ' ' * 7 + '│' + ' ' * 6 + '│' +
                 '│'.join(_f4(v).rjust(9) for v in cells) + '│' +
                 _f4(razem).rjust(10) + '│' + _f4(ogolem).rjust(10) + '│')
        return s

    lines = [_TK_PCL + agencja.ljust(125) + f"Strona {strona_start:2d}",
             (' Zestawienie powierzchni gruntów i miąższości drzewostanu wg gatunków'
              ' panujących (głównych) wg klas i podklas wieku.'.ljust(118) +
              f"Obiekt: {obiekt}").ljust(166)]
    lines.extend(_TK_H)

    # wiersze gatunków
    for gat in gat_order + ['__razem__']:
        if gat == '__razem__':
            gat_pow[gat] = [sum(gat_pow[g][i] for g in gat_order) for i in range(15)]
            gat_m3[gat] = [sum(gat_m3[g][i] for g in gat_order) for i in range(15)]
            name = ' Razem '
        else:
            pad = 7 - len(gat)
            name = ' ' * (pad // 2) + gat + ' ' * (pad - pad // 2)
        lines.append(wiersz(gat, m3=False))
        part = '│' + name + '│      ├---------┼---------┼---------┼---------┼---------┼---------┼---------┼---------┼---------┼---------┼---------┼---------┼---------┼---------┼---------┼----------┼----------┤'
        lines.append(part)
        lines.append(wiersz(gat, m3=True))
        lines.append(_TK_H[-1])
    lines[-1] = _TK_BOT

    # stopki części 1: „inne grunty (staw, jez.)" = suma POW_INNE z działek rejestrowych
    inne = sum(float(d.get('POW_INNE') or 0) for d in dane['D'])
    dozal = sum(float(o.get('POW_WYDZ') or 0) for o in dane['O']
                if 400 <= (o.get('RODZ_POW') or 0) <= 409)
    lasy = sum(gat_pow['__razem__'])
    for label, v in (('inne grunty (staw, jez.)', inne), ('grunty do zal.', dozal),
                     ('Razem lasy', lasy)):
        lines.append(' ' * 146 + label.ljust(26) + '-' + _f4(v).rjust(10))

    # --- część 2: siedliskowe typy lasu (bez niezalesionych bagien 321) ---
    stl_pow = {}
    inne_siedl = 0.0
    for o in dane['O']:
        kod = int(o.get('TYP_SIED') or 0)
        if str(o.get('RODZ_POW') or '').strip() == '321':
            inne_siedl += float(o.get('POW_WYDZ') or 0)
            continue
        stl_pow[kod] = stl_pow.get(kod, 0.0) + float(o.get('POW_WYDZ') or 0)
    kody = sorted(stl_pow)
    nazwy = [_STL.get(k, str(k)) for k in kody]
    ncol = len(kody) + 1
    w2 = max(39, ncol * 10 - 1)   # nagłówek rozciąga się, gdy kolumn > 4
    lines.append('\f\r\x1b(s16.67H\x1b&l9E\x1b&a8L 2. Zestawienie powierzchni siedliskowych typów lasu - ' + obiekt)
    lines.append('┌──────────────────┬' + '─' * w2 + '┐')
    lines.append('│ Wyszczególnienie │' + 'Siedliskowe typy lasu'.center(w2) + '│')
    lines.append('│                  ├' + '┬'.join(['─' * 9] * ncol) + '┤')
    cells = ['  ' + n.ljust(7) for n in nazwy] + ['  Razem  ']
    lines.append('│                  │' + '│'.join(cells) + '│')
    lines.append('├──────────────────┼' + '┼'.join(['─' * 9] * ncol) + '┤')
    vals = [_f4(stl_pow[k]).rjust(9) for k in kody] + [_f4(sum(stl_pow.values())).rjust(9)]
    lines.append('│ Powierzchnia ha  │' + '│'.join(vals) + '│')
    lines.append('└──────────────────┴' + '┴'.join(['─' * 9] * ncol) + '┘')
    lines.append('')
    lines.append('  powierzchnie inne = ' + _f4(inne_siedl).rjust(9) + ' ha')
    lines.append('')
    lines.append('')

    # --- część 3: lasy ochronne ---
    ochr_pow = {}
    for o in dane['O']:
        k = int(o.get('KAT_OCH') or 0)
        ochr_pow[k] = ochr_pow.get(k, 0.0) + float(o.get('POW_WYDZ') or 0)
    kody = sorted(ochr_pow)
    lines.append('3. Zestawienie powierzchni lasów ochronnych')
    cat_w = 20 * len(kody) + max(0, len(kody) - 1) + 1 + 9
    inner = '┬'.join(['─' * 20] * len(kody) + ['─' * 9])
    lines.append('┌──────────────────┬' + '─' * cat_w + '┐')
    lines.append('│ Wyszczególnienie │' +
                 'Kategoria ochronnosci'.center(cat_w) + '│')
    lines.append('│                  ├' + inner + '┤')
    lines.append('│                  │' + '│'.join(
        _OCHR.get(k, str(k)).ljust(20) for k in kody) + '│  Razem  ' + '│')
    lines.append('├──────────────────┼' + inner.replace('┬', '┼') + '┤')
    lines.append('│ Powierzchnia ha  │' + '│'.join(
        [_f4(ochr_pow[k]).rjust(14).ljust(20) for k in kody] +
        [_f4(sum(ochr_pow.values())).rjust(9)]) + '│')
    lines.append('└──────────────────┴' + inner.replace('┬', '┴') + '┘')
    lines.append('')
    lines.append('')

    # --- część 4: drzewostany do przebudowy (RODZ_POW 199) ---
    przebud_pow, przebud_zasob = 0.0, 0
    for o in dane['O']:
        if str(o.get('RODZ_POW') or '').strip() != '199':
            continue
        key = f"{o.get('ODDZIAL', '')}{o.get('PODODDZ', '')}".strip()
        p = float(o.get('POW_WYDZ') or 0)
        przebud_pow += p
        m3 = 0
        for r in dane.get('R_all', {}).get(key, []):
            z = float(r.get('ZASOB') or 0)
            if z:
                m3 = max(m3, round(z * p))
        przebud_zasob += m3
    lines.append('4. Zestawienie powierzchni dla drzewostanów do przebudowy')
    lines.append('')
    lines.append(f"      Razem  pow:{_f4(przebud_pow).rjust(10)} ha        "
                 f"Zasobność:{round(przebud_zasob):>5} m3")
    lines.append('\f')

    out = dane['o_path'].parent / 'TAB_KLW3.TXT'
    _zapisz(out, lines, tail='')
    return out


# ----------------------------------------------------------------------------
# REJESTR1.TXT — Rejestr działek wg właścicieli
# ----------------------------------------------------------------------------

_RJ_PCL = '\x1b(s16.67H\x1b&l4E\x1b&a1L'
_RJ_H = (
    '┌────────┬──────────────────────────────────────────────────┬───────┬──────┬──────────────────────────────────────────────────────────────┬─────────┬──┬───────────────────────────┬───────┐',
    '│        │                                                  │       │      │                  Opis i powierzchnia lasów                   │  Pow.   │ O│   Wskazania godspodarcze  │       │',
    '│   nr   │  Nazwisko i imię                                 │   Nr  │Oddz. ├──────────────────────┬───────────────────┬─────────┬─────────┤ gruntów │ c├──────────┬─────────┬──────┤ Wykon.│',
    '│rejestru│  Adres właściciela                               │działki│poddz.│       zalesiona      │   nie zalesiona   │  inne   │  razem  │   do    │ h│  Rodzaj  │  Pow.   │ Miąż.│       │',
    '│        │  Współwłaściciele                                │       │      ├────┬───┬───┬─────────┼─────────┬─────────┤ grunty  │  lasy   │  zal.   │ r│  zabiegu │  [ha]   │  m3  │       │',
    '│        │                                                  │       │      │gat.│ W │Bon│   Pow.  │ do odn. │ pozost. │         │         │         │ .│          │         │      │       │',
    '├────────┼──────────────────────────────────────────────────┼───────┼──────┼────┼───┼───┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼──┼──────────┼─────────┼──────┼───────┤',
    '│   1    │                    2                             │   3   │  4   │ 5  │ 6 │ 7 │    8    │    9    │   10    │   11    │   12    │   13    │14│    15    │    16   │  17  │  18   │',
    '├────────┼──────────────────────────────────────────────────┼───────┼──────┼────┼───┼───┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼──┼──────────┼─────────┼──────┼───────┤',
)
_RJ_BOT = '└────────┴──────────────────────────────────────────────────┴───────┴──────┴────┴───┴───┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴──┴──────────┴─────────┴──────┴───────┘'


def _rj_row(nr_poz, nazwa, d_item, wsk_item, dane):
    """Wiersz danych REJESTR1 (16 segmentów, kolumny 3-18)."""
    seg = [''] * 16
    if d_item is not None:
        d, r, o = d_item
        lz = float(d.get('POW_L_ZAL') or 0)
        lnz = float(d.get('POW_L_NZAL') or 0)
        hal = lnz > 0
        gat = 'hal.' if hal else (str(r.get('GATUNEK', '') or '').strip() if r else '')
        wiek = '' if (hal or not r or not r.get('WIEK')) else str(r.get('WIEK'))
        bon = str(r.get('BONIT', '') or '').strip() if r else ''
        seg = [
            str(d.get('NR_DZIAL', '') or '').strip(),
            f"{str(d.get('ODDZIAL', '') or '').strip()}{str(d.get('PODODDZ', '') or '').strip()}",
            gat, wiek, bon,
            _f4(lz) if lz > 0 else '',
            _f4(lnz) if lnz > 0 else '',
            _f4(d.get('POW_N_ZAL') or 0) if (d.get('POW_N_ZAL') or 0) > 0 else '',
            _f4(d.get('POW_INNE') or 0) if (d.get('POW_INNE') or 0) > 0 else '',
            _f4(d.get('POW') or 0),
            '', '',
        ] + [''] * 4
    if wsk_item is not None:
        w, pw, mh, mz, d = wsk_item
        if pw > 0:
            pow_z = pw / 100.0 * float(d.get('POW') or 0)
            miaz = f"{mh * pow_z:.1f}" if mh > 0 else ''
        else:                      # wskaźnik bez udziału powierzchni
            pow_z, miaz = None, ''
        seg[12:16] = [_wsk_rj_kod(w, mz), _f4(pow_z), miaz, '']
    if nr_poz:
        _lp, _nr = nr_poz.split('/')
        _s = f"{_lp}/{_nr}"
        _L = 4 - len(_lp) - (1 if len(_nr) >= 3 else 0)
        _R = 8 - _L - len(_s)
        c1 = ' ' * _L + _s + ' ' * _R
    else:
        c1 = ' ' * 8
    row = '│' + c1 + '│' + f"{nazwa}".ljust(50)
    widths = [7, 6, 4, 3, 3, 9, 9, 9, 9, 9, 9, 2, 10, 9, 6, 7]
    for k, (v, wd) in enumerate(zip(seg, widths)):
        v = '' if v is None else str(v)
        row += '│' + (v.ljust(wd) if k in (2, 4, 12) else v.rjust(wd))
    return row + '│'


def generuj_rejestr1_txt(obreb_dir, dane=None, agencja=None):
    """REJESTR1.TXT (po przeniesieniu halizn w D*.DBF)."""
    obreb = Path(obreb_dir)
    if dane is None:
        dane = _wczytaj_obreb(obreb)
    if dane is None:
        return None
    if agencja is None:
        agencja = czytaj_agencje(obreb) or ''
    obiekt, _ = czytaj_dane_wsi(obreb)
    pozycje = _pozycje_rejestru(dane)

    els = []          # ('sep',) | ('row',...,is_pool) | ('dash',) | ('razem', label, sums)
    tot_obj = [0.0] * 5
    for lp, (nrrej, ws, ds) in enumerate(pozycje, start=1):
        els.append(('sep',))
        # strumień: każdy D + jego WSK1, potem jego WSK2..WSK6
        stream = []
        for d in ds:
            key = f"{d.get('ODDZIAL', '')}{d.get('PODODDZ', '')}".strip()
            r = dane['R_by'].get(key, {})
            o = dane['O_by'].get(key, {})
            zasob = float(r.get('ZASOB') or 0)
            wsks = _wsks(r)
            if wsks:
                w, pw, mz = wsks[0]
                stream.append(('D', d, r, o, (w, pw, _miaz_na_ha(w, mz, zasob), mz, d)))
            else:
                stream.append(('D', d, r, o, None))
            for w, pw, mz in wsks[1:]:
                stream.append(('Z', d, r, o, (w, pw, _miaz_na_ha(w, mz, zasob), mz, d)))
        # wylewanie na pulę wierszy właścicieli (nazwisko/adres), potem dopisywane
        pool = []
        for w in ws:
            pool.append((w, 'name'))
            pool.append((w, 'addr'))
        for i, item in enumerate(stream):
            if i < len(pool):
                w, kind = pool[i]
                if kind == 'name':
                    nazwa = f"{str(w.get('NAZWISKO', '') or '').strip()} " \
                            f"{str(w.get('IMIE', '') or '').strip()} " \
                            f"{str(w.get('RODZICE', '') or '').strip()}".strip()
                    nr_poz = f"{lp}/{nrrej}"
                else:
                    nazwa = str(w.get('ADRES', '') or '').strip()
                    nr_poz = ''
                is_pool = True
            else:
                nazwa, nr_poz, is_pool = '', '', False
            if item[0] == 'D':
                els.append(('row', nr_poz, nazwa, (item[1], item[2], item[3]),
                            item[4], is_pool))
            else:
                els.append(('row', nr_poz, nazwa, None, item[4], is_pool))
        # nadmiarowi właściciele (bez działek w tej pozycji) — same puste pola
        for w, kind in pool[len(stream):]:
            if kind == 'name':
                nazwa = f"{str(w.get('NAZWISKO', '') or '').strip()} " \
                        f"{str(w.get('IMIE', '') or '').strip()} " \
                        f"{str(w.get('RODZICE', '') or '').strip()}".strip()
                nr_poz = f"{lp}/{nrrej}"
            else:
                nazwa = str(w.get('ADRES', '') or '').strip()
                nr_poz = ''
            els.append(('row', nr_poz, nazwa, None, None, True))
        # blok "Razem"
        seen_dz, dz_sum, poz_sums = [], {}, [0.0] * 5
        for d in ds:
            dz = str(d.get('NR_DZIAL', '') or '').strip()
            if dz not in dz_sum:
                dz_sum[dz] = 0.0
                seen_dz.append(dz)
            dz_sum[dz] += float(d.get('POW') or 0)
            for k, f in enumerate(('POW_L_ZAL', 'POW_L_NZAL', 'POW_N_ZAL', 'POW_INNE')):
                poz_sums[k] += float(d.get(f) or 0)
            poz_sums[4] += float(d.get('POW') or 0)
        els.append(('dash',))
        for dz in seen_dz:
            els.append(('razem', f" Razem dzialka {dz}".ljust(28) +
                        _f4(dz_sum[dz]) + ' ha', None))
        els.append(('razem', f" Razem pozycja {lp:>3}/{nrrej}".ljust(28) +
                    _f4(poz_sums[4]) + ' ha', poz_sums))
        for k in range(5):
            tot_obj[k] += poz_sums[k]
    els.append(('sep',))
    els.append(('razem', " Razem obiekt".ljust(22) + _f4(tot_obj[4]) + ' ha', tot_obj))

    # ---- paginacja: łamanie po wierszach puli (>=24) i przed pozycjami ----
    def _row_out(el):
        typ = el[0]
        if typ == 'sep':
            return _RJ_H[-1]
        if typ == 'dash':
            return ('│' + ' ' * 8 + '│' + ' ' * 50 + '│' + ' ' * 7 + '│' + ' ' * 6 +
                    '│    │   │   │' + '---------│' * 6 + '  │' + ' ' * 10 + '│' +
                    ' ' * 9 + '│' + ' ' * 6 + '│' + ' ' * 7 + '│')
        if typ == 'razem':
            _, label, sums = el
            s = ('│' + ' ' * 8 + '│' + label.ljust(50) + '│' + ' ' * 7 + '│' + ' ' * 6 +
                 '│    │   │   │')
            if sums is None:
                s += '---------│' * 6
            else:
                s += ''.join(_f4(v).rjust(9) + '│' for v in sums)
                s += '0.0000'.rjust(9) + '│'
            s += '  │' + ' ' * 10 + '│' + ' ' * 9 + '│' + ' ' * 6 + '│' + ' ' * 7 + '│'
            return s
        _, nr_poz, nazwa, d_item, wsk_item, is_pool = el
        return _rj_row(nr_poz, nazwa, d_item, wsk_item, dane)

    pages, cur, rows = [], [], 0

    def eject():
        nonlocal cur, rows
        if cur:
            pages.append(cur)
        cur, rows = [], 0

    for el in els:
        if el[0] == 'sep':
            if rows >= 24:
                eject()
            if rows == 0:
                continue          # nagłówek strony już kończy się SEP
            cur.append(_row_out(el))
            rows += 1
            continue
        cur.append(_row_out(el))
        rows += 1
        if el[0] == 'row' and el[5] and rows >= 24:
            eject()

    out_lines = []
    for pageno, pg in enumerate(pages + [cur] if cur else pages, start=1):
        out_lines.extend(_strona_rej(pg, pageno, agencja, obiekt))
    out_lines.append('\f\x1bE')

    out = dane['d_path'].parent / 'REJESTR1.TXT' if dane['d_path'] else obreb / 'REJESTR1.TXT'
    _zapisz(out, out_lines, tail='')
    return out


def _strona_rej(rows, pageno, agencja, obiekt):
    hdr = ('\r' if pageno == 1 else '\f\r') + _RJ_PCL + agencja.ljust(134) + f"Strona {pageno:5d}"
    return [hdr,
            (' Rejestr działek leśnych i gruntów do zalesienia wg. właścicieli'.ljust(75) +
             f"Obiekt: {obiekt}").ljust(123)] + list(_RJ_H) + list(rows) + [_RJ_BOT]


# ----------------------------------------------------------------------------
# WSKAZ1.TXT — wykaz wskaźników (pusty przy braku danych)
# ----------------------------------------------------------------------------

# Słownik rozwinięć kodów wskaźników (SWSKAZA.LST mietka;
# przy braku pliku używany jest wbudowany odpowiednik)
_ZADANIA = {
    'CP': 'czyszczenia późne',
    'CP w 2naw.': 'CP w 2 nawrotach',
    'CP z m3': 'CP z masą',
    'CS': 'cięcia sanitarne',
    'CW': 'czyszczenia wczesne',
    'Dol.': 'dolesienia',
    'Inne': 'pozostałe wskazania',
    'Magr.oczyś': 'mel.agrotech. - oczyścić',
    'Magr.wyrów': 'mel.agrotech. - wyrównać',
    'Mel.agr.': 'melioracje agrotechniczne',
    'Mel.wodne': 'melioracje wodne',
    'Naw.': 'nawożenie',
    'Oczyścić': 'oczyścić',
    'Odn.': 'odnowić',
    'Piel.': 'pielęgnowanie uprawy',
    'Piel.p.poz': 'pielegnowac pasy p.poz.',
    'Pods.': 'podsadzenia',
    'Popr.': 'poprawki',
    'Pozostawić': 'pozostawić',
    'Przeklas.': 'przeklasyfikowac',
    'Rb I': 'rębnia I',
    'Rb II': 'rębnia II',
    'Rb III': 'rębnia III',
    'Rb IV': 'rębnia IV',
    'TP': 'trzebież późna',
    'TW': 'trzebież wczesna',
    'TW w 2naw.': 'TW w 2 nawrotach',
    'Uprzątnąc': 'uprzątnąć',
    'Us.nas.': 'usunąć nasienniki',
    'Us.przedr.': 'usunąć przedrosty',
    'Us.przest.': 'usunąć przestoje',
    'Uzup.': 'uzupełnienia',
    'Wpr.podsz.': 'wprowadzenie podszytu',
    'Wyrównać': 'wyrównać',
    'Zalesić': 'zalesić',
    '16Xdo29II': 'wykonywać pom. 16X-29II',
    'drz.dziupl': 'pozostawić drz. dziuplast',
    'do5l po Rb': 'do 5 lat po Rb',
    'Nat.2000': 'obszar Natura 2000',
    'O.chr.kr.': 'Obszar Chron.Krajobr.',
    'Wykonyw.': 'wykonywać',
}

def _czytaj_zadania(obreb_dir):
    """Słownik kod -> nazwa zadania; z SWSKAZA.LST mietka albo wbudowany."""
    kandydaci = []
    obreb = Path(obreb_dir)
    for bazowy in (obreb, obreb.parent):
        for nazwa in ('SWSKAZA.LST', 'SWSKAZA .LST'):
            kandydaci.append(bazowy / nazwa)
    for path in kandydaci:
        try:
            if not path.exists():
                continue
            recs = czytaj_dbf(path)
            if recs and recs[0].get('SKROT') is not None:
                sl = {str(r.get('SKROT', '') or '').strip():
                      str(r.get('NAZWA', '') or '').strip() for r in recs}
                sl.pop('', None)
                if sl:
                    return sl
        except Exception:
            continue
    return dict(_ZADANIA)

_W1_TOP = '┌─────────┬───────┬───────────────────┬───────────────────┬─────────────────────────────────────────────────────────────────┬────────────┐'
_W1_H1 = '│         │Oddział│   Powierzchnia    │Skrócony opis lasu │           Zadania w zakresie gospodarki leśnej                  │            │'
_W1_H2 = '│  Numer  │poddz. ├─────────┬─────────┤(gat.gł.,wiek,bon.,├────────────────────────────────────┬────────────┬───────────────┤            │'
_W1_H3 = '│ działki │na     │  lasu   │ gruntu  │ rodzaj pow. ochr.)│                Rodzaj              │Powierzchnia│Maks.miąż. (m3)│    Uwagi   │'
_W1_H4 = '│         │mapie  │         │ do zal. │według stanu na:   │                zadania             │     w      ├───────┬───────┤            │'
_W1_H5 = '│         │gospod.├─────────┴─────────┤                   │                                    │   [ ha ]   │ p.ręb.│rębnym │            │'
_W1_SEP1 = '├─────────┼───────┼─────────┬─────────┼───────────────────┼────────────────────────────────────┼────────────┼───────┼───────┼────────────┤'
_W1_SEP = '├─────────┼───────┼─────────┼─────────┼───────────────────┼────────────────────────────────────┼────────────┼───────┼───────┼────────────┤'
_W1_NUM = '│    1    │   2   │    3    │    4    │         5         │                   6                │      7     │   8   │   9   │     10     │'
_W1_RSEP = '├─────────────────┼─────────┼─────────┼───────────────────┴────────────────────────────────────┴────────────┼───────┼───────┼────────────┤'
_W1_BOT = '└─────────────────┴─────────┴─────────┴─────────────────────────────────────────────────────────────────────┴───────┴───────┴────────────┘'

def generuj_wskaz1_txt(obreb_dir, dane=None, agencja=None):
    """WSKAZ1.TXT — zadania w zakresie gospodarki leśnej na 10-lecie,
    rozbite na pozycje rejestru (jedna strona = jedna pozycja)."""
    obreb = Path(obreb_dir)
    if dane is None:
        dane = _wczytaj_obreb(obreb)
    if dane is None:
        return None
    o_path = dane['o_path']
    out_dir = o_path.parent if o_path else obreb
    out = out_dir / 'WSKAZ1.TXT'
    if agencja is None:
        agencja = czytaj_agencje(obreb) or ''
    obiekt, stan = czytaj_dane_wsi(obreb)

    # daty 10-lecia z WSIE.DBF
    od_txt = do_txt = ''
    wsie_path = znajdz_dbf(obreb, 'WSIE')
    if wsie_path is not None:
        try:
            w = czytaj_dbf(wsie_path)
            if w:
                v = str(w[0].get('OBOW_OD', '') or '').strip()
                if re.match(r'^\d{8}$', v):
                    od_txt = f"{v[6:8]}-{v[4:6]}-{v[0:4]}"
                v = str(w[0].get('OBOW_DO', '') or '').strip()
                if re.match(r'^\d{8}$', v):
                    do_txt = f"{v[6:8]}-{v[4:6]}-{v[0:4]}"
        except Exception:
            pass

    zadania = _czytaj_zadania(obreb)
    pozycje = _pozycje_rejestru(dane)

    def _row(c1, c2, c3, c4, c5, c6, c7, c8, c9):
        return ('│' + c1 + '│' + c2 + '│' + c3 + '│' + c4 + '│' + c5 +
                '│' + c6 + '│' + c7 + '│' + c8 + '│' + c9 + '│' + ' ' * 12 + '│')

    pcl = '\r\x1b(s16.67H\x1b&l5E\x1b&a20L' + (agencja or '').ljust(40)
    chunks = []
    for lp, (nrrej, ws, ds) in enumerate(pozycje, start=1):
        if not ds:
            continue
        rows = []
        razem_l = razem_n = 0.0
        razem8 = razem9 = 0
        for d in ds:
            key = f"{d.get('ODDZIAL', '')}{d.get('PODODDZ', '')}".strip()
            r = dane['R_by'].get(key, {}) or {}
            lz = float(d.get('POW_L_ZAL') or 0)
            ln = float(d.get('POW_L_NZAL') or 0)
            nz = float(d.get('POW_N_ZAL') or 0)
            pow_l = lz + ln
            razem_l += pow_l
            razem_n += nz
            hal = ln > 0
            gat = 'hal.' if hal else str(r.get('GATUNEK', '') or '').strip()
            wiek = '' if hal else str(r.get('WIEK', '') or '').strip()
            bon = str(r.get('BONIT', '') or '').strip()
            c5 = f"{gat:<4}-{wiek:>3} - {bon:<3}".ljust(19)
            zasob = float(r.get('ZASOB') or 0)
            wsks = _wsks(r) if r else []
            for j, (wsk, pw, mz) in enumerate(wsks):
                kod = str(wsk).split('-')[0].strip()
                nazwa = zadania.get(kod, '')
                c6 = f"{kod} : {nazwa}"[:36].ljust(36)
                pow_w = pow_l * pw / 100.0
                c7 = _f4(pow_w).rjust(12)
                c8 = c9 = ' ' * 7
                mh = _miaz_na_ha(wsk, mz, zasob)
                if kod.startswith('Rb'):
                    v = round(zasob * pow_w)
                    if v > 0:
                        c9 = str(v).rjust(7)
                elif kod.startswith(('TP', 'TW')):
                    v = round(mh * pow_w)
                    if v > 0:
                        c8 = str(v).rjust(7)
                razem8 += int(c8) if c8.strip() else 0
                razem9 += int(c9) if c9.strip() else 0
                if j == 0:
                    c1 = str(d.get('NR_DZIAL', '') or '').strip().ljust(9)
                    c2 = key.rjust(7)
                    c3 = _f4(pow_l).rjust(9)
                    c4 = _f4(nz).rjust(9) if nz > 0 else ' ' * 9
                else:
                    c1 = ' ' * 9
                    c2 = ' ' * 7
                    c3 = ' ' * 9
                    c4 = ' ' * 9
                    c5 = ' ' * 19
                rows.append(_row(c1, c2, c3, c4, c5, c6, c7, c8, c9))
        razem = ('│ Razem:          │' + _f4(razem_l).rjust(9) + '│' +
                 _f4(razem_n).rjust(9) + '│' + ' ' * 69 + '│' +
                 str(razem8).rjust(7) + '│' + str(razem9).rjust(7) + '│' +
                 ' ' * 12 + '│')
        h6 = ('│         │       │        [ha]       │ ' + (stan or '').ljust(18) +
              '│                                    │            │       │       │            │')
        tab = '\r\n'.join([_W1_TOP, _W1_H1, _W1_H2, _W1_H3, _W1_H4, _W1_H5, h6,
                            _W1_SEP1, _W1_NUM, _W1_SEP] + rows +
                           [_W1_RSEP, razem, _W1_BOT])
        head = pcl + '\r\n Obiekt: ' + obiekt + '\r\n' + ' ' * 38 + \
            'ZADANIA W ZAKRESIE GOSPODARKI LEŚNEJ\r\n' + ' ' * 38 + \
            f'na okres od {od_txt} do {do_txt}' + '\r\n\r\n'
        wlasc = ''
        for w in ws:
            nazw = str(w.get('NAZWISKO', '') or '').strip()
            imie = str(w.get('IMIE', '') or '').strip()
            rodz = str(w.get('RODZICE', '') or '').strip()
            nazwa = 'P. ' + ' '.join(x for x in (nazw, imie, rodz) if x)
            wlasc += nazwa.ljust(79) + f"Nr rej:{int(w.get('NRREJ') or 0):>6}" + \
                '\r\n' + f"Adres  :{str(w.get('ADRES', '') or '').strip()}".ljust(68) + '\n\r'
        chunks.append(('\x0c' if chunks else '') + head + wlasc + tab + '\r\n')
    out.write_bytes((''.join(chunks) + '\x0c').encode('cp852'))
    return out


# ----------------------------------------------------------------------------
# WYK_NEG.TXT — zestawienie drzewostanów negatywnych i źle produkujących
# ----------------------------------------------------------------------------

_WN_PCL = '\r\x1b(s16.67H\x1b&l9E\x1b&a8L'
_WN_GORA = '┌───────┬────────────┬─────────┬────┬──────────┐'
_WN_H1 = '│ Oddz. │  Skrócony  │   Pow.  │ Zas│   Uwagi  │'
_WN_H2 = '│ Podod.│  opis lasu │   [ha]  │ m3 │          │'
_WN_SEP = '├───────┼────────────┼─────────┼────┼──────────┤'
_WN_SEP_RAZEM = '├───────┴────────────┼─────────┼────┼──────────┤'
_WN_DOL = '└────────────────────┴─────────┴────┴──────────┘'


def generuj_wyk_neg_txt(obreb_dir, dane=None, agencja=None):
    """WYK_NEG.TXT — drzewostany negatywne (RODZ_POW 198) i źle produkujące (199).

    Brak takich wydzieleń => plik pusty (0 B), jak w MIETEKU.
    """
    obreb = Path(obreb_dir)
    if dane is None:
        dane = _wczytaj_obreb(obreb)
    if dane is None:
        return None
    o_path = dane['o_path']
    out = o_path.parent / 'WYK_NEG.TXT' if o_path else obreb / 'WYK_NEG.TXT'

    neg = [r for r in dane['O'] if (r.get('RODZ_POW') or 0) in (198, 199)]
    if not neg:
        out.write_bytes(b'')
        return out

    obiekt, _ = czytaj_dane_wsi(obreb)
    neg.sort(key=_klucz_wydz)

    lines = [
        _WN_PCL + 'Zestawienie powierzchni i zasobnosci dla drzewostanów '
                  'neg. i zle produkujacych  - ' + obiekt,
        _WN_GORA, _WN_H1, _WN_H2, _WN_SEP,
    ]
    suma_pow, suma_zas = 0.0, 0
    for o in neg:
        key = f"{o.get('ODDZIAL', '')}{o.get('PODODDZ', '')}".strip()
        r = dane['R_by'].get(key, {})
        oddz = f"{str(o.get('ODDZIAL', '') or '').strip()}{str(o.get('PODODDZ', '') or '').strip()}"
        gat = str(r.get('GATUNEK', '') or '').strip()
        bonit = str(r.get('BONIT', '') or '').strip()[:3]
        wiek = int(r.get('WIEK') or 0)
        opis = f"{gat:<4}-{bonit:<3}-{wiek:>3}"
        pow_w = o.get('POW_WYDZ') or 0
        zas = round((r.get('ZASOB') or 0) * pow_w)
        suma_pow += pow_w
        suma_zas += zas
        lines.append('│' + f"{oddz:<7}│{opis}│{pow_w:>9.4f}│{zas:>4}│          │")

    lines.append(_WN_SEP_RAZEM)
    lines.append('│           Razem    │' + f"{suma_pow:>9.4f}│{suma_zas:>4}│          │")
    lines.append(_WN_DOL)
    out.write_bytes(('\r\n'.join(lines) + '\r\n').encode('cp852'))
    return out


# ----------------------------------------------------------------------------
# WSK_ZB.TXT — zestawienie czynności gospodarczych (zbiorcze wskaźniki)
# ----------------------------------------------------------------------------

_WZ_PCL = '\r\x1b&l6E\x1b&a10L\x1b(s3T'

# grupy wskaźników (kody jak w SWSKAZA.LST / kodzie MIETEKA)
_WZ_REBNE = {'Rb I', 'Rb II', 'Rb III', 'Rb IV'}
_WZ_POZ_REBNE = {'Uprzątnąć', 'uprz.płaz', 'Us.przest.', 'Us.nas.', 'Us.przedr.'}
_WZ_CZ_POLNE_Z_MASA = {'CP z m3'}
_WZ_TRZ_WCZESNE = {'TW', 'TW w 2naw.'}
_WZ_TRZ_POZNE = {'TP'}
_WZ_SANITARNE = {'CS'}
_WZ_ZALES = {'Zalesić'}
_WZ_ODN = {'Odn.'}
_WZ_POPR_UZUP = {'Popr.', 'Uzup.'}
_WZ_DOL = {'Dol.'}
_WZ_PIEL_UPR = {'Piel.', 'CW'}
_WZ_PIEL_MLOD = {'CP', 'CP w 2naw.'}
_WZ_PODSZYT = {'Wpr.podsz.'}
_WZ_PODSADZ = {'Pods', 'Pods.'}
_WZ_MEL_AGR = {'Mel.agr.', 'Magr.', 'Magr.oczyś', 'Magr.wyrówn.'}
_WZ_MEL_WODNE = {'Mel.wodne'}


def _obow_daty(obreb_dir):
    """Daty obowiązywania planu z WSIE.DBF (OBOW_OD/OBOW_DO) jako dd-mm-yyyy."""
    wsie_path = znajdz_dbf(Path(obreb_dir), 'WSIE')
    if wsie_path is not None:
        try:
            recs = czytaj_dbf(wsie_path)
            if recs:
                od = str(recs[0].get('OBOW_OD', '') or '').strip()
                do = str(recs[0].get('OBOW_DO', '') or '').strip()
                if re.match(r'^\d{8}$', od) and re.match(r'^\d{8}$', do):
                    return (f"{od[6:8]}-{od[4:6]}-{od[0:4]}",
                            f"{do[6:8]}-{do[4:6]}-{do[0:4]}")
        except Exception:
            pass
    # fallback: 10-lecie od roku po dacie stanu
    _, stan = czytaj_dane_wsi(Path(obreb_dir))
    rok = int(stan[:4]) + 1 if stan[:4].isdigit() else 2027
    return (f'01-01-{rok}', f'31-12-{rok + 9}')


def _wz_wystapienia(dane):
    """Wystąpienia wskaźników: (wsk, pw%, miaz, pow_wydz, zasob, rodz_pow)."""
    out = []
    o_by = dane['O_by']
    for r in dane['R']:
        key = f"{r.get('ODDZIAL', '')}{r.get('PODODDZ', '')}".strip()
        o = o_by.get(key)
        pow_w = o.get('POW_WYDZ') or 0 if o else 0
        rodz = o.get('RODZ_POW') or 0 if o else 0
        for i in range(1, 7):
            w = str(r.get(f'WSK{i}', '') or '').strip()
            if not w:
                continue
            w = re.sub(r'-\d+m3/ha$', '', w)
            pw = r.get(f'POW_WSK{i}') or 0
            mz = r.get(f'MIAZ{i}') or 0
            out.append((w, pw, mz, pow_w, r.get('ZASOB') or 0, rodz))
    return out


def _wz_pow(occ, kody):
    """Suma powierzchni [ha] wskazań z danej grupy (ważona procentem)."""
    return sum(pw / 100.0 * pow_w for w, pw, mz, pow_w, z, rodz in occ if w in kody)


def _wz_m3(occ, kody, tryb='wazony'):
    """Suma miążdżności [m3] dla grupy.

    tryb 'wazony'  — m3/ha × procent × powierzchnia (m3/ha: MIAZ, dla Rb — ZASOB,
                      dla CS bez MIAZ — 5% ZASOB);
    tryb 'raw'     — suma pol MIAZ bez ważenia (przestoje/p przedrostki).
    """
    s = 0.0
    for w, pw, mz, pow_w, z, rodz in occ:
        if w not in kody:
            continue
        if tryb == 'raw':
            s += mz
            continue
        m3ha = mz
        if m3ha <= 0:
            if w.startswith('Rb'):
                m3ha = z
            elif w == 'CS':
                m3ha = 0.05 * z
        s += m3ha * (pw / 100.0) * pow_w
    return s


def _wz_linia(label, pow_ha, m3=None):
    wiersz = f"{label:<37}" + f"{pow_ha:.4f} ha".rjust(10)
    if m3 is not None:
        wiersz += f"{round(m3)} m3".rjust(14)
    else:
        wiersz += '  '
    return wiersz


def generuj_wsk_zb_txt(obreb_dir, dane=None, agencja=None):
    """WSK_ZB.TXT — zestawienie czynności gospodarczych na 10-lecie."""
    obreb = Path(obreb_dir)
    if dane is None:
        dane = _wczytaj_obreb(obreb)
    if dane is None:
        return None
    o_path = dane['o_path']
    out = o_path.parent / 'WSK_ZB.TXT' if o_path else obreb / 'WSK_ZB.TXT'
    if agencja is None:
        agencja = czytaj_agencje(obreb) or ''
    obiekt, _ = czytaj_dane_wsi(obreb)
    od_dat, do_dat = _obow_daty(obreb)

    occ = _wz_wystapienia(dane)

    # --- I. użytkowanie ---
    rebne_pow = _wz_pow(occ, _WZ_REBNE)
    rebne_m3 = _wz_m3(occ, _WZ_REBNE)
    poz_pow = _wz_pow(occ, _WZ_POZ_REBNE)
    poz_m3 = _wz_m3(occ, _WZ_POZ_REBNE, tryb='raw')
    czp_pow = _wz_pow(occ, _WZ_CZ_POLNE_Z_MASA)
    czp_m3 = _wz_m3(occ, _WZ_CZ_POLNE_Z_MASA)
    tw_pow = _wz_pow(occ, _WZ_TRZ_WCZESNE)
    tw_m3 = _wz_m3(occ, _WZ_TRZ_WCZESNE)
    tp_pow = _wz_pow(occ, _WZ_TRZ_POZNE)
    tp_m3 = _wz_m3(occ, _WZ_TRZ_POZNE)
    cs_pow = _wz_pow(occ, _WZ_SANITARNE)
    cs_m3 = _wz_m3(occ, _WZ_SANITARNE)

    razem_rebne_pow = rebne_pow + poz_pow
    razem_rebne_m3 = rebne_m3 + poz_m3
    razem_przed_pow = czp_pow + tw_pow + tp_pow + cs_pow
    razem_przed_m3 = czp_m3 + tw_m3 + tp_m3 + cs_m3
    ogo_pow = razem_rebne_pow + razem_przed_pow
    ogo_m3 = razem_rebne_m3 + razem_przed_m3

    # --- II. hodowla ---
    zales_pow = _wz_pow(occ, _WZ_ZALES)
    odn_pow = _wz_pow(occ, _WZ_ODN)

    # rozbicie odnowień wg RODZ_POW wydzieleń z wskazaniem Odn.
    zreby = hal_płaz = hal_rol = 0.0
    for w, pw, mz, pow_w, z, rodz in occ:
        if w not in _WZ_ODN:
            continue
        p = pw / 100.0 * pow_w
        if rodz == 240:
            zreby += p
        elif 241 <= rodz <= 245:
            hal_płaz += p
            if rodz in (244, 245):
                hal_rol += p
    # powierzchnie leśne niezalesione pozostałe: z działek minus te w odnowieniach
    lnzal = sum(d.get('POW_L_NZAL') or 0 for d in dane['D'])
    niezal_ost = lnzal - zreby - hal_płaz

    razem_odn_zales = zales_pow + odn_pow
    popr_uzup_pow = _wz_pow(occ, _WZ_POPR_UZUP)
    popr_spodz_pow = round(0.2 * odn_pow, 4)
    dol_pow = _wz_pow(occ, _WZ_DOL)
    piel_upr_pow = _wz_pow(occ, _WZ_PIEL_UPR)
    piel_mlod_pow = _wz_pow(occ, _WZ_PIEL_MLOD)
    podszyt_pow = _wz_pow(occ, _WZ_PODSZYT)
    podsadz_pow = _wz_pow(occ, _WZ_PODSADZ)
    mel_agr_pow = _wz_pow(occ, _WZ_MEL_AGR)
    mel_wod_pow = _wz_pow(occ, _WZ_MEL_WODNE)

    lines = [
        _WZ_PCL + agencja.ljust(40),
        '',
        ' Zestawienie czynności gospodarczych projektowanych do wykonania',
        f' w 10-leciu od {od_dat} do {do_dat} wg. wskazań gospodarczych',
        f' dla obiektu {obiekt}',
        '',
        ' I. Użytkowanie lasu',
        '',
        '   A. Użytkowanie rębne            ',
        _wz_linia('      1. Użytki rębne właściwe', rebne_pow, rebne_m3),
        _wz_linia('      2. Pozostałe użytki rębne', poz_pow, poz_m3),
        '   ' + '-' * 61,
        _wz_linia('   Ogółem użytki rębne', razem_rebne_pow, razem_rebne_m3),
        '   ' + '-' * 61,
        '',
        '   B. Użytkowanie przedrębne       ',
        _wz_linia('      1. cz. późne z masą', czp_pow, czp_m3),
        _wz_linia('      2. trzebieże wczesne', tw_pow, tw_m3),
        _wz_linia('      3. trzebieże późne', tp_pow, tp_m3),
        _wz_linia('      4. cięcia sanitarne', cs_pow, cs_m3),
        '   ' + '-' * 60,
        _wz_linia('   Razem użytki przedrębne', razem_przed_pow, razem_przed_m3),
        '   ' + '=' * 60,
        _wz_linia('   Ogółem użytkowanie w 10-leciu', ogo_pow, ogo_m3),
        '',
        '',
        '',
        'II. Hodowla lasu ',
        _wz_linia('   1. Zalesienia', zales_pow),
        _wz_linia('   2. Odnowienia', odn_pow),
        _wz_linia('       Zręby', zreby),
        _wz_linia('         w tym zręby bież.', 0.0),
        _wz_linia('       Halizny i płazowiny', hal_płaz),
        _wz_linia('         w tym halizny uż. rol.', hal_rol),
        _wz_linia('   3. Pow. leśne niezal. pozostałe', niezal_ost),
        '   ' + '-' * 60,
        _wz_linia('   Razem odnowienia i zalesienia', razem_odn_zales),
        '   ' + '-' * 60,
        _wz_linia('   3. Poprawki i uzupełnienia', popr_uzup_pow),
        _wz_linia('   4. Poprawki spodziewane', popr_spodz_pow),
        _wz_linia('   5. Dolesienie luk', dol_pow),
        _wz_linia('   6. Pielęgnowanie upraw', piel_upr_pow),
        _wz_linia('   7. Pielęgnowanie młodników', piel_mlod_pow),
        _wz_linia('   8. Wprowadzanie podszytów', podszyt_pow),
        _wz_linia('   9. Podsadzenia produkcyjne', podsadz_pow),
        '',
        _wz_linia('  10. Melioracje agrotechniczne', mel_agr_pow),
        _wz_linia('  11. Melioracje wodne', mel_wod_pow),
        '\x0c',
    ]
    out.write_bytes(('\r\n'.join(lines)).encode('cp852'))
    return out


# ----------------------------------------------------------------------------
# Generowanie całości
# ----------------------------------------------------------------------------

def generuj_wszystkie_po_przeniesieniu(obreb_dir, agencja=None, tylko=None):
    """Generuje wydruki MIETEKA (po przeniesieniu halizn).

    tylko: opcjonalny zbiór nazw plików, które mają powstać
           (np. {'OPTAX.TXT', 'WSKAZ1.TXT'}); None = komplet.
    Numeracja stron TAB_KLW3 kontynuuje OPTAX (jak w MIETEKU).
    Zwraca słownik {nazwa: ścieżka}.
    """
    wybrane = {str(t).upper() for t in tylko} if tylko else None

    def _chk(nazwa):
        return wybrane is None or nazwa in wybrane

    dane = _wczytaj_obreb(obreb_dir)
    if dane is None:
        return {}
    pozycje = _pozycje_rejestru(dane)
    out = {}
    optax_stron = 0
    if _chk('OPTAX.TXT'):
        try:
            p, optax_stron = generuj_optax_txt(obreb_dir, dane=dane, pozycje=pozycje,
                                               agencja=agencja)
            if p:
                out['OPTAX.TXT'] = p
        except Exception:
            import traceback
            traceback.print_exc()
    elif _chk('TAB_KLW3.TXT'):
        # OPTAX nie jest generowany — liczbę stron bierzemy z istniejącego pliku
        d = dane['o_path'].parent if dane.get('o_path') else Path(obreb_dir)
        try:
            optax_stron = (d / 'OPTAX.TXT').read_bytes().count(b'\x0c')
        except Exception:
            optax_stron = 0
    for fn, gen, kw in (
        ('TAB_KLW3.TXT', generuj_tab_klw3_txt, dict(strona_start=optax_stron or 1)),
        ('ZEST1.TXT', generuj_zest1_txt, dict(pozycje=pozycje)),
        ('REJESTR1.TXT', generuj_rejestr1_txt, dict()),
        ('WSKAZ1.TXT', generuj_wskaz1_txt, dict()),
        ('WYK_NEG.TXT', generuj_wyk_neg_txt, dict()),
        ('WSK_ZB.TXT', generuj_wsk_zb_txt, dict()),
    ):
        if not _chk(fn):
            continue
        try:
            p = gen(obreb_dir, dane=dane, agencja=agencja, **kw)
            if p:
                out[fn] = p
        except Exception:
            import traceback
            traceback.print_exc()
    return out
