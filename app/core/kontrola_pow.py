# -*- coding: utf-8 -*-
"""Kontrola powierzchni REJESTR <-> OPTAX.

Skan rozbieżności między Rejestrem (działki, D*.DBF) i opisem taksacyjnym
(wydzielenia, O*.DBF) oraz stosowanie decyzji użytkownika podjętych
w podsumowaniu kreatora Pełnego Automatu:

  usun    - usuwa z Rejestru wszystkie rekordy działek wydzielenia
  dopisz  - dopisuje rekord O (opis taksacyjny wpisany przez użytkownika)
  rozloz  - działki wydzielenia znikają, a ich powierzchnia odejmuje się
            proporcjonalnie od pozostałych działek tego samego oddziału
            (OPTAX mówi prawdę o łącznej powierzchni)
  usunO   - usuwa rekord O (wydzielenie bez działek w Rejestrze)
  zostaw  - bez zmian

Każdy zapis do DBF poprzedza kopia .BAK (konwencja edytora Mietek v2.0:
pierwsza kopia zostaje nienaruszona).
"""

from pathlib import Path

import struct
from datetime import datetime

from app.core.wydruki import czytaj_dbf, znajdz_dbf

_EPS = 0.00005          # tolerancja zgodności sum


def _key(rec):
    return f"{rec.get('ODDZIAL', '')}{rec.get('PODODDZ', '')}".strip()


def _oddz(key):
    """Część numeryczna klucza wydzielenia ('12ab' -> '12')."""
    i = 0
    while i < len(key) and key[i].isdigit():
        i += 1
    return key[:i]


def _sort_key(key):
    try:
        return (int(_oddz(key) or 9999), key)
    except ValueError:
        return (9999, key)


# --------------------------------------------------------------------------
# skan rozbieżności
# --------------------------------------------------------------------------

def policz_rozbieznosci(obreb_dir):
    """Porównanie sum per wydzielenie dla jednego obrębu (folder z DBF).

    Zwraca dict {folder, sumy:{rej,opt}, wydz:[{wydz,rej,opt,uwaga,dzialki}]}
    albo None, gdy brak D/O w folderze.
    """
    obreb = Path(obreb_dir)
    d_path, o_path = znajdz_dbf(obreb, 'D'), znajdz_dbf(obreb, 'O')
    if d_path is None or o_path is None:
        return None
    D, O = czytaj_dbf(d_path), czytaj_dbf(o_path)
    w_path = znajdz_dbf(obreb, 'W')
    W = czytaj_dbf(w_path) if w_path else []

    wlasc = {}
    for w in W:
        nr = w.get('NRREJ')
        if nr is None:
            continue
        nazw = f"{str(w.get('NAZWISKO', '') or '').strip()} " \
               f"{str(w.get('IMIE', '') or '').strip()}".strip()
        if nazw:
            wlasc.setdefault(nr, []).append(nazw)

    rej = {}                      # key -> [suma, [(dz, nrrej, pow, wlasc)]]
    for d in D:
        key = _key(d)
        pow_d = float(d.get('POW') or 0)
        w = rej.setdefault(key, [0.0, []])
        w[0] += pow_d
        nr = d.get('NRREJ')
        w[1].append((str(d.get('NR_DZIAL', '') or '').strip(), nr, pow_d,
                     ' / '.join(wlasc.get(nr, [])) or '(brak nazwiska)'))
    opt = {}
    for o in O:
        key = _key(o)
        opt[key] = opt.get(key, 0.0) + float(o.get('POW_WYDZ') or 0)

    wydz = []
    for key in set(rej) | set(opt):
        r_pow = rej.get(key, (0.0, []))[0]
        o_pow = opt.get(key)
        if o_pow is None:
            uwaga = 'BRAK W OPTAX'
        elif key not in rej:
            uwaga = 'BRAK W REJESTRZE'
        elif abs(r_pow - o_pow) > _EPS:
            uwaga = 'RÓŻNICA'
        else:
            continue
        wydz.append({
            'wydz': key,
            'rej': r_pow if key in rej else None,
            'opt': o_pow,
            'uwaga': uwaga,
            'dzialki': [{'dz': dz, 'nrrej': nr, 'pow': p, 'wlasc': kt}
                        for dz, nr, p, kt in (rej.get(key, (0.0, []))[1])],
        })
    wydz.sort(key=lambda w: _sort_key(w['wydz']))
    return {
        'folder': str(obreb),
        'nazwa': obreb.parent.name or obreb.name,
        'sumy': {'rej': sum(v[0] for v in rej.values()),
                 'opt': sum(opt.values())},
        'wydz': wydz,
    }


# --------------------------------------------------------------------------
# decyzje użytkownika -> DBF
# --------------------------------------------------------------------------

def _bak(path):
    """Kopia .BAK (pierwsza kopia zostaje — jak w edytorze Mietek v2.0)."""
    import shutil
    path = Path(path)
    bak = path.with_suffix(path.suffix + '.BAK')
    if not bak.exists():
        shutil.copy2(path, bak)
        return str(bak)
    return None


def _fmt_num(v, dec):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v or 0)
    return f"{v:.{dec}f}" if dec else f"{round(v):d}"


def zastosuj_decyzje(obreb_dir, decyzje):
    """Zastosuj decyzje [{wydz, akcja, opis}] do D/O jednego obrębu.

    Zwraca (log, nowe_sumy, usuniete). Log to lista komunikatów do
    pokazania użytkownikowi; nowe_sumy = {'rej': ..., 'opt': ...} po
    zmianach; usuniete = lista usuniętych wydzieleń (do raportu kontroli
    powierzchni z dopiskiem „usunięte”):
    [{wydz, akcja, rej, opt, dzialki: [(nr_dzialki, nrrej, pow)]}].
    """

    obreb = Path(obreb_dir)
    d_path, o_path = znajdz_dbf(obreb, 'D'), znajdz_dbf(obreb, 'O')
    if d_path is None or o_path is None:
        return [f"Brak plików D*.DBF / O*.DBF w {obreb} — pomijam."], None

    d_pola, d_rek = _read_dbf(d_path)
    o_pola, o_rek = _read_dbf(o_path)
    log, odrzucone, usuniete = [], [], []

    # aktualny stan (do walidacji decyzji)
    def _sumy_d():
        rej = {}
        for r in d_rek:
            rej[_key(r)] = rej.get(_key(r), 0.0) + float(r.get('POW') or 0)
        return rej
    rej_d = _sumy_d()
    opt_d = {}
    for r in o_rek:
        opt_d[_key(r)] = opt_d.get(_key(r), 0.0) + float(r.get('POW_WYDZ') or 0)

    # --- walidacja i grupowanie decyzji ---
    do_usun, do_usunO, do_dopisz = [], [], []
    for dz in (decyzje or []):
        key = str(dz.get('wydz', '') or '').strip()
        akcja = str(dz.get('akcja', '') or '').strip()
        opis = str(dz.get('opis', '') or '').strip()
        if not key or akcja in ('', 'zostaw'):
            continue
        w_rej, w_opt = key in rej_d, key in opt_d
        if akcja in ('usun', 'rozloz') and w_rej:
            # wariant A: działki wydzielenia znikają z Rejestru,
            # pozostałe działki bez zmian (sumy zrównują się z OPTAX)
            do_usun.append((key, akcja))
        elif akcja == 'usunO' and w_opt:
            do_usunO.append(key)
        elif akcja == 'dopisz' and w_rej and not w_opt:
            do_dopisz.append((key, opis))
        else:
            odrzucone.append(f"{key}/{akcja}")
    if odrzucone:
        log.append("Odrzucone (nie pasują do aktualnych danych): "
                   + ", ".join(odrzucone))

    zmiana_d = zmiana_o = False

    # --- 1. USUŃ / ROZŁÓŻ z Rejestru ---
    if do_usun:
        keys = {k for k, _a in do_usun}
        for key in sorted(keys):
            # szczegóły usuniętych rekordów — do raportu kontroli
            dzl = [(str(r.get('NR_DZIAL', '') or '').strip(),
                    r.get('NRREJ'), float(r.get('POW') or 0))
                   for r in d_rek if _key(r) == key]
            usuniete.append({'wydz': key, 'akcja': 'usun',
                             'rej': rej_d.get(key, 0.0),
                             'opt': opt_d.get(key, 0.0),
                             'dzialki': dzl})
        d_rek = [r for r in d_rek if _key(r) not in keys]
        zmiana_d = True
        for key, akcja in do_usun:
            slowo = ("ROZŁOŻONO (działki znikają z Rejestru, pozostałe "
                     "bez zmian)" if akcja == 'rozloz'
                     else "USUNIĘTO z Rejestru (działki znikają z wydruków)")
            log.append(f"{slowo}: wydzielenie {key} "
                       f"({rej_d.get(key, 0):.4f} ha)")

    # --- 2. USUŃ z OPTAX ---
    if do_usunO:
        for key in sorted(set(do_usunO)):
            usuniete.append({'wydz': key, 'akcja': 'usunO',
                             'rej': rej_d.get(key, 0.0),
                             'opt': opt_d.get(key, 0.0),
                             'dzialki': []})
        o_rek = [r for r in o_rek if _key(r) not in set(do_usunO)]
        zmiana_o = True
        for key in do_usunO:
            log.append(f"USUNIĘTO z opisu taksacyjnego: wydzielenie {key} "
                       f"({opt_d.get(key, 0):.4f} ha)")

    # --- 3. DOPISZ opis taksacyjny ---
    if do_dopisz:
        for key, opis in do_dopisz:
            pow_d = rej_d.get(key, 0.0)
            nowy = {}
            # rekord budowany po nazwach pól (brakujące -> domyślne)
            for nm, typ, dl, dec in o_pola:
                if nm == 'ODDZIAL':
                    nowy[nm] = _oddz(key)
                elif nm == 'PODODDZ':
                    nowy[nm] = key[len(_oddz(key)):]
                elif nm == 'POW_WYDZ':
                    nowy[nm] = _fmt_num(pow_d, dec)
                elif nm == 'OP_TAX':
                    nowy[nm] = opis
                elif nm == 'OP_TAX1':
                    nowy[nm] = ''
                elif typ == 'N':
                    nowy[nm] = '0'
                else:
                    nowy[nm] = ''
            o_rek.append(nowy)
            zmiana_o = True
            log.append(f"DOPISANO opis taksacyjny: wydzielenie {key} "
                       f"({pow_d:.4f} ha) — opis: "
                       f"{(opis[:60] + '…') if len(opis) > 60 else opis or '(pusty)'}")
    # --- zapis z .BAK ---
    if zmiana_d:
        bak = _bak(d_path)
        _write_dbf(d_path, d_pola, d_rek)
        log.append(f"Zapisano {d_path.name}" + (f" (kopia: {bak})" if bak
                                                   else " (.BAK już istniał)"))
    if zmiana_o:
        bak = _bak(o_path)
        _write_dbf(o_path, o_pola, o_rek)
        log.append(f"Zapisano {o_path.name}" + (f" (kopia: {bak})" if bak
                                                   else " (.BAK już istniał)"))
    if not (zmiana_d or zmiana_o):
        log.append("Nic do zmiany.")

    # nowe sumy
    rej_nowe, opt_nowe = 0.0, 0.0
    for r in d_rek:
        rej_nowe += float(r.get('POW') or 0)
    for r in o_rek:
        opt_nowe += float(r.get('POW_WYDZ') or 0)
    return log, {'rej': rej_nowe, 'opt': opt_nowe}, usuniete


# --------------------------------------------------------------------------
# odczyt/zapis DBF (kopia z tab_mietek_plus10 — core nie może ciągnąć GUI)
# --------------------------------------------------------------------------

def _read_dbf(filename):
    """Odczytuje plik dBase III zwracając (fields, records) — wartości tekstowe."""
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
                    rec[name] = raw.decode('ascii', 'replace').replace('\x00', '').strip()
                elif typ == 'D':
                    rec[name] = raw.decode('ascii', 'replace').strip()
                elif typ == 'L':
                    rec[name] = chr(raw[0]) if raw else ''
                else:
                    rec[name] = raw.decode('cp852', 'replace').strip()
            records.append(rec)
    return fields, records


def _write_dbf(filename, fields, records):
    """Zapisuje listę rekordów do dBase III."""
    num_records = len(records)
    header_length = 32 + (len(fields) * 32) + 1
    record_length = 1 + sum(f[2] for f in fields)
    with open(filename, 'wb') as f:
        f.write(struct.pack('<B', 0x03))
        now = datetime.now()
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
                    v = (str(val)[:1] if str(val) else " ") or " "
                    f.write(v.encode('ascii', errors='replace'))
                else:
                    val_bytes = str(val).encode('cp852', errors='replace')[:length].ljust(length, b' ')
                    f.write(val_bytes)
        f.write(struct.pack('<B', 0x1A))
