"""
Forestly — Zadania Excel
==================================
Zależności: config.py (EXCEL_SHEET_DEFAULTS)
Odpowiada za: funkcje pomocnicze dla Excela — parsowanie właścicieli,
              łączenie plików XLS/VAL, wykonywanie makr VBA, formatowanie arkuszy.

Te funkcje są standalone (nie wymagają instancji ModernApp).
Można je testować w izolacji.
"""

import re
import warnings
import json
from pathlib import Path

import pandas as pd
import numpy as np
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.utils import get_column_letter

from app.config import EXCEL_SHEET_DEFAULTS

# Wyciszanie ostrzeżeń z xlrd przy czytaniu starych .xls
warnings.filterwarnings("ignore", message=".*OLE2 inconsistency.*")
warnings.filterwarnings("ignore", message=".*file size.*not.*sector size.*")
warnings.filterwarnings("ignore", message=".*SSCS size.*")

def bezpieczna_liczba(val):
    if pd.isna(val) or val == '' or val == 'nan':
        return 0.0
    s = str(val).replace(',', '.')
    s = re.sub(r'\s+', '', s)
    try:
        return float(s)
    except ValueError:
        return 0.0


def wczytaj_i_przetworz_wlascicieli(sciezka_do_pliku):
    # DYNAMICZNE SZUKANIE NAGŁÓWKA
    df_raw = pd.read_excel(sciezka_do_pliku, header=None, nrows=20)
    header_row = 1
    for i, row in df_raw.iterrows():
        row_str = " ".join([str(val).lower() for val in row.values])
        if 'numer działki' in row_str or 'numer dzialki' in row_str:
            header_row = i
            break

    df = pd.read_excel(sciezka_do_pliku, header=header_row)
    # Odporny rename — łapie znak nowej linii, \n literal, i warianty
    rename_map = {}
    for c in df.columns:
        c_str = str(c)
        if 'klasouż' in c_str.lower() and 'pow' in c_str.lower():
            rename_map[c] = 'Pow. klasouż.'
        if c_str == 'Numer działki':
            rename_map[c] = 'nr_dz'
    df = df.rename(columns=rename_map)

    kolumny_do_wypelnienia = ['nr_dz', 'J. rej.', 'Pow. działki', 'Właściciel']
    istniejace_kolumny = [col for col in kolumny_do_wypelnienia if col in df.columns]

    for col in istniejace_kolumny:
        df[col] = df[col].astype(str).replace(r'^\s*$', np.nan, regex=True)
        df[col] = df[col].replace('nan', np.nan)
        df[col] = df[col].replace('-', np.nan)

    if istniejace_kolumny:
        df[istniejace_kolumny] = df[istniejace_kolumny].ffill()

    if 'nr_dz' in df.columns:
        df['nr_dz'] = df['nr_dz'].astype(str).str.strip()
        df['nr_dz'] = df['nr_dz'].str.replace(r'\.0$', '', regex=True)

    if 'J. rej.' in df.columns:
        def extract_after_g(val):
            v_str = str(val)
            if 'G' in v_str:
                return v_str.split('G')[-1]
            return v_str

        df['J. rej.'] = df['J. rej.'].apply(extract_after_g)

    if 'Właściciel' in df.columns:
        df['Właściciel'] = df['Właściciel'].astype(str).replace(['nan', 'NaN'], 'Brak danych')

    # --- DODANA LOGIKA POBIERANIA WŁAŚCICIELA DO GŁÓWNEJ PAMIĘCI ---
    cols_full = ['nr_dz', 'J. rej.', 'Pow. działki']
    if 'Właściciel' in df.columns:
        cols_full.append('Właściciel')

    df_full = df[cols_full].drop_duplicates(
        ['nr_dz', 'J. rej.']).copy() if 'nr_dz' in df.columns and 'Pow. działki' in df.columns else pd.DataFrame()
    # ---------------------------------------------------------------

    # --- jednostki: MIETEKA zapisuje powierzchnie w m² — przeliczamy na ha.
    # Autodetekcja: jeśli w pliku są już wartości w ha (stare XLS-y), nie dzielimy.
    def _czy_m2(col):
        if col not in df.columns:
            return False
        return bool(df[col].apply(bezpieczna_liczba).dropna().gt(50).any())
    dzielnik = 10000.0 if (_czy_m2('Pow. działki') or _czy_m2('Pow. klasouż.')) else 1.0

    if not df_full.empty:
        df_full['pow dz'] = df_full['Pow. działki'].apply(bezpieczna_liczba) / dzielnik

    wiersze_po_rozbiciu = []
    for _, row in df.iterrows():
        klasy = str(row.get('Klasoużytek', '')).split('\n')
        powierzchnie = str(row.get('Pow. klasouż.', '')).split('\n')
        max_len = max(len(klasy), len(powierzchnie))
        klasy += [''] * (max_len - len(klasy))
        powierzchnie += [''] * (max_len - len(powierzchnie))
        for k, p in zip(klasy, powierzchnie):
            nowy_wiersz = row.copy()
            nowy_wiersz['Klasoużytek'] = k
            nowy_wiersz['Pow. klasouż.'] = p
            wiersze_po_rozbiciu.append(nowy_wiersz)

    df_exploded = pd.DataFrame(wiersze_po_rozbiciu)

    for col in ['Pow. działki', 'Pow. klasouż.']:
        if col in df_exploded.columns:
            df_exploded[col] = df_exploded[col].apply(bezpieczna_liczba) / dzielnik

    if 'Klasoużytek' in df_exploded.columns:
        df_ls = df_exploded[df_exploded['Klasoużytek'].astype(str).str.contains('Ls', case=False, na=False)].copy()
    else:
        df_ls = pd.DataFrame(columns=df_exploded.columns)

    if df_ls.empty:
        wynik = pd.DataFrame(columns=['nr_dz', 'J. rej.', 'pow dz', 'pow ls', 'Właściciel'])
        return wynik, df_full

    wynik = df_ls.groupby(['nr_dz', 'J. rej.', 'Pow. działki', 'Właściciel'], as_index=False)['Pow. klasouż.'].sum()
    wynik = wynik.rename(columns={'Pow. działki': 'pow dz', 'Pow. klasouż.': 'pow ls'})
    wynik['pow dz'] = wynik['pow dz'].round(4)
    wynik['pow ls'] = wynik['pow ls'].round(4)
    wynik = wynik[['nr_dz', 'J. rej.', 'pow dz', 'pow ls', 'Właściciel']]
    return wynik, df_full


def wczytaj_i_przetworz_val(sciezka_do_pliku_val):
    try:
        with open(sciezka_do_pliku_val, 'r', encoding='cp1250') as file:
            linie = file.readlines()
    except Exception as e:
        print(f"Błąd wczytywania VAL: {e}")
        return None

    dane_wyjsciowe = []
    aktualny_nr_dz = None

    for line in reversed(linie):
        line = line.strip()
        if not line or line.startswith(';'):
            continue
        elementy = line.split()
        if not elementy:
            continue
        if elementy[0] == '*':
            if len(elementy) >= 2:
                aktualny_nr_dz = elementy[1]
        elif elementy[0] == '^':
            if len(elementy) >= 3:
                oznaczenie = elementy[1]
                if 'X' not in oznaczenie and re.search(r'[A-Za-z]', oznaczenie):
                    litera = oznaczenie
                    if aktualny_nr_dz:
                        pow_sqm_str = elementy[2]
                        try:
                            pow_geo = float(pow_sqm_str.replace(',', '.')) / 10000.0  # m² -> ha
                        except ValueError:
                            continue
                        if pow_geo >= 0.001:
                            dane_wyjsciowe.append({
                                'nr_dz': aktualny_nr_dz,
                                'litera': litera,
                                'pow geo': round(pow_geo, 4)
                            })

    dane_wyjsciowe.reverse()
    df = pd.DataFrame(dane_wyjsciowe)
    if df.empty:
        return pd.DataFrame(columns=['nr_dz', 'litera', 'pow geo'])
    df = df[['nr_dz', 'litera', 'pow geo']]
    return df


def polacz_xls_i_val(df_xls, df_full, df_val):
    xls = df_xls.copy()
    val = df_val.copy()

    df_merged = pd.merge(val, xls, on='nr_dz', how='left')

    mapping_j_rej = df_full.set_index('nr_dz')['J. rej.']
    mapping_pow_dz = df_full.set_index('nr_dz')['pow dz']

    df_merged['J. rej.'] = df_merged['J. rej.'].fillna(df_merged['nr_dz'].map(mapping_j_rej))
    df_merged['pow dz'] = df_merged['pow dz'].fillna(df_merged['nr_dz'].map(mapping_pow_dz))

    # --- NOWA LOGIKA RATOWANIA NAZWISK DLA DZIAŁEK "PRZYBYŁO" ---
    if 'Właściciel' in df_full.columns:
        mapping_wlasciciel = df_full.set_index('nr_dz')['Właściciel']
        df_merged['Właściciel'] = df_merged['Właściciel'].fillna(df_merged['nr_dz'].map(mapping_wlasciciel))
    # ------------------------------------------------------------

    df_out = pd.DataFrame()
    df_out['Kolumna_A'] = ""
    df_out['J. rej.'] = df_merged['J. rej.']
    df_out['nr_dz'] = df_merged['nr_dz']
    df_out['litery'] = df_merged['litera']
    df_out['pow geo'] = df_merged['pow geo']
    df_out['ROZLICZONE'] = np.nan
    df_out['Kolumna_G'] = ""
    df_out['nr_dz_ewid'] = df_merged['nr_dz']
    df_out['pow ls'] = df_merged['pow ls']
    df_out['pow dz'] = df_merged['pow dz']
    df_out['właściciel'] = df_merged['Właściciel']

    nieotaksowane = xls[~xls['nr_dz'].isin(val['nr_dz'])].copy() if not xls.empty else pd.DataFrame(
        columns=['J. rej.', 'nr_dz', 'Właściciel', 'pow ls', 'pow dz'])
    if not nieotaksowane.empty:
        nieotaksowane = nieotaksowane[['J. rej.', 'nr_dz', 'Właściciel', 'pow ls', 'pow dz']]
        nieotaksowane = nieotaksowane.rename(columns={'Właściciel': 'właściciel'})

    return df_out, nieotaksowane


def wykonaj_makro_vba(df_out, df_braki, tylko_wyrownywanie=False):
    df = df_out.copy()
    df['bg_color'] = ""
    df['font_color'] = ""
    TOLERANCJA = 0.0010

    # 1. WARTOŚCI (bez żadnych kolorów tła)
    # Zmieniamy grupowanie z samego 'nr_dz' na ['nr_dz', 'J. rej.'] aby współwłaściciele nie wpływali na siebie
    for (dz, j_rej), group in df.groupby(['nr_dz', 'J. rej.'], sort=False, dropna=False):
        # Zabezpieczenie przed dublami: sumujemy unikalne kontury, by uniknąć inflacji powierzchni
        unikalne_geo = group.drop_duplicates(subset=['litery'])
        suma_geo = unikalne_geo['pow geo'].sum()

        pow_ewid = group['pow ls'].iloc[0]
        pow_docelowa = group['pow dz'].iloc[0]
        is_new_forest = pd.isna(pow_ewid) or str(pow_ewid).strip() == ""
        df_font = 'FF0000' if is_new_forest else '000000'

        # Checkbox: Wymuszenie wyrównania blokuje tworzenie nadmiarów (zielonych działek)
        if tylko_wyrownywanie:
            nadmiar_sciezka = False
        else:
            nadmiar_sciezka = pd.notna(pow_ewid) and suma_geo > (float(pow_ewid) + 0.1)

        suma_przepisanych = 0.0

        for idx in group.index:
            aktualna_pow = group.at[idx, 'pow geo']
            df.at[idx, 'font_color'] = df_font

            if is_new_forest:
                df.at[idx, 'ROZLICZONE'] = aktualna_pow
                continue

            if nadmiar_sciezka:
                if pd.notna(pow_docelowa):
                    reszta = float(pow_docelowa) - suma_przepisanych
                    if reszta > 0:
                        wartosc = min(reszta, aktualna_pow)
                        df.at[idx, 'ROZLICZONE'] = round(wartosc, 4)
                        suma_przepisanych += wartosc
                    else:
                        df.at[idx, 'ROZLICZONE'] = 0.0000
                else:
                    df.at[idx, 'ROZLICZONE'] = aktualna_pow
            else:
                if pd.notna(pow_ewid) and suma_geo != 0:
                    nowa = (aktualna_pow / suma_geo) * float(pow_ewid)
                    zaokr = round(nowa, 4)
                    df.at[idx, 'ROZLICZONE'] = zaokr if zaokr != 0 else aktualna_pow
                else:
                    df.at[idx, 'ROZLICZONE'] = aktualna_pow

    # 2. DOCIĄGANIE RÓŻNIC ZAOKRĄGLEŃ
    for (dz, j_rej), group in df.groupby(['nr_dz', 'J. rej.'], sort=False, dropna=False):
        pow_ewid = group['pow ls'].iloc[0]
        pow_docelowa = group['pow dz'].iloc[0]

        valid_indices = group[group['ROZLICZONE'].notna()].index
        if len(valid_indices) == 0:
            continue
        suma_f = df.loc[valid_indices, 'ROZLICZONE'].sum()
        roznica = 0.0
        if pd.notna(pow_docelowa) and str(pow_docelowa).strip() != "":
            pow_j = float(pow_docelowa)
            if suma_f > pow_j:
                roznica = pow_j - suma_f
            elif 0 < (pow_j - suma_f) <= TOLERANCJA:
                roznica = pow_j - suma_f
        if roznica == 0.0 and pd.notna(pow_ewid) and str(pow_ewid).strip() != "":
            pow_i = float(pow_ewid)
            if abs(pow_i - suma_f) > 0 and abs(pow_i - suma_f) <= TOLERANCJA:
                roznica = pow_i - suma_f
        if roznica != 0:
            ostatni_wiersz = valid_indices[-1]
            df.at[ostatni_wiersz, 'ROZLICZONE'] = round(
                df.at[ostatni_wiersz, 'ROZLICZONE'] + roznica, 4)

    # 3. SZUM -> RÓŻOWY
    rows_to_drop = []
    for idx in df.index:
        val = df.at[idx, 'ROZLICZONE']
        pow_ewid = df.at[idx, 'pow ls']
        if pd.notna(val) and val <= 0.004:
            if pd.isna(pow_ewid) or str(pow_ewid).strip() == "":
                rows_to_drop.append(idx)
            else:
                df.at[idx, 'bg_color'] = 'FFB6C1'
    if rows_to_drop:
        df = df.drop(index=rows_to_drop)

    # 4. PRZYBYŁO / UBYŁO
    przybylo_data = []
    ubylo_data = []

    # Checkbox wymusza, by ominąć generowanie arkuszy UBYŁO/PRZYBYŁO i nie mazać na zielono
    if not tylko_wyrownywanie:
        for dz, group in df.groupby('nr_dz', sort=False):
            # Unikalne kontury dla całej działki, aby uniknąć zdublowania sumy
            unikalne_geo = group.drop_duplicates(subset=['litery'])
            suma_f = unikalne_geo['ROZLICZONE'].sum() if not unikalne_geo.empty else 0.0

            pow_ewid = group['pow ls'].iloc[0]
            pow_docelowa = group['pow dz'].iloc[0]
            j_rej = group['J. rej.'].iloc[0] if 'J. rej.' in group.columns else ""
            startowy_las = float(pow_ewid) if (pd.notna(pow_ewid) and str(pow_ewid).strip() != "") else 0.0
            roznica = round(suma_f - startowy_las, 4)

            if roznica > 0:
                for idx in group.index:
                    if df.at[idx, 'bg_color'] != 'FFB6C1' and pd.notna(df.at[idx, 'ROZLICZONE']):
                        df.at[idx, 'bg_color'] = '00FF00'
                przybylo_data.append({
                    'J. rej.': j_rej, 'nr działki': dz,
                    'aktualna pow ls': round(suma_f, 4), 'ls ewidenca': startowy_las,
                    'ile przybyło': roznica,
                    'pow dz': pow_docelowa if pd.notna(pow_docelowa) else ""
                })
            elif roznica < 0:
                ubylo_data.append({
                    'J. rej.': j_rej, 'nr działki': dz,
                    'aktualna pow ls': round(suma_f, 4), 'ls ewidenca': startowy_las,
                    'ile ubyło': roznica,
                    'pow dz': pow_docelowa if pd.notna(pow_docelowa) else ""
                })

        if not df_braki.empty:
            for _, row in df_braki.iterrows():
                if '[OP]' not in str(row.get('właściciel', '')):
                    pow_ewid = row.get('pow ls', np.nan)
                    pow_doc = row.get('pow dz', np.nan)
                    j_rej = row.get('J. rej.', "")
                    if pd.notna(pow_ewid) and float(pow_ewid) > 0:
                        ubylo_data.append({
                            'J. rej.': j_rej, 'nr działki': row.get('nr_dz', ''),
                            'aktualna pow ls': 0.0, 'ls ewidenca': pow_ewid,
                            'ile ubyło': -float(pow_ewid),
                            'pow dz': pow_doc if pd.notna(pow_doc) else ""
                        })

    return df, pd.DataFrame(przybylo_data), pd.DataFrame(ubylo_data)



def formatuj_arkusz_raportowy(worksheet, tytul, hex_kolor_tytulu):
    worksheet['A1'] = tytul
    worksheet.merge_cells('A1:F1')
    worksheet['A1'].font = Font(size=18, bold=True, color=hex_kolor_tytulu)
    worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')

    thick_bottom = Border(bottom=Side(style='thick', color='000000'))
    thin_border = Border(left=Side(style='thin', color='000000'),
                         right=Side(style='thin', color='000000'),
                         top=Side(style='thin', color='000000'),
                         bottom=Side(style='thin', color='000000'))

    max_row = worksheet.max_row
    max_col = 6

    for col in range(1, max_col + 1):
        cell = worksheet.cell(row=2, column=col)
        cell.font = Font(bold=True)
        cell.border = thick_bottom

    for row in range(3, max_row + 1):
        for col in range(1, max_col + 1):
            cell = worksheet.cell(row=row, column=col)
            cell.border = thin_border
            if col in [1, 2]:
                cell.alignment = Alignment(horizontal='left')

    for col in range(1, max_col + 1):
        col_letter = get_column_letter(col)
        max_length = 0
        for row in range(2, max_row + 1):
            cell = worksheet.cell(row=row, column=col)
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        worksheet.column_dimensions[col_letter].width = (max_length + 2)

# Struktura WSIE.DBF dokładnie wg specyfikacji MIETEK.EXE (kolejność krytyczna!)
WSIE_FIELDS = [
    ('NAZWA', 'C', 40, 0), ('WOJEW', 'C', 30, 0), ('GMINA', 'C', 30, 0),
    ('STAN_NA', 'D', 8, 0), ('OBOW_OD', 'D', 8, 0), ('OBOW_DO', 'D', 8, 0),
    ('NR_WSI', 'N', 3, 0), ('ROK_ZAL', 'C', 2, 0),
    ('SPR', 'C', 75, 0), ('ZLC', 'C', 40, 0), ('WW', 'C', 20, 0), ('KR', 'C', 5, 0),
    ('DZ1', 'C', 75, 0), ('DZ2', 'C', 75, 0),
    ('ET1', 'N', 6, 0), ('ET2', 'N', 6, 0), ('ET3', 'N', 6, 0), ('ET4', 'N', 6, 0),
    ('ET5', 'N', 6, 0), ('ET6', 'N', 6, 0), ('ET7', 'N', 6, 0),
    ('OCHR2', 'C', 75, 0), ('OCHR3', 'C', 75, 0), ('OCHR4', 'C', 75, 0),
    ('P_OCH', 'N', 12, 4),
    ('ZDR', 'C', 75, 0), ('ZDR1', 'C', 75, 0), ('ZDR2', 'C', 75, 0),
    ('ZG1', 'N', 12, 4), ('ZG2', 'N', 12, 4), ('ZG3', 'N', 12, 4),
    ('PRZY', 'C', 75, 0), ('PRZY1', 'C', 75, 0), ('PRZY2', 'C', 75, 0),
    ('SANITAR', 'C', 75, 0), ('SANITAR1', 'C', 75, 0), ('SANITAR2', 'C', 75, 0),
    ('US1', 'C', 75, 0), ('US2', 'C', 75, 0), ('US3', 'C', 75, 0), ('US4', 'C', 75, 0),
    ('EG1', 'C', 50, 0), ('EG2', 'C', 50, 0), ('EG3', 'C', 50, 0),
    ('EG4', 'C', 50, 0), ('EG5', 'C', 50, 0),
    ('POWIAT', 'C', 30, 0),
]

def zrob_zestawienie_zbiorcze(folder):
    """Zestawienie zbiorcze rozliczeń całego obrębu.

    Skanuje folder z plikami <WIEŚ>_Rozliczone.xlsx i tworzy jeden plik
    ZESTAWIENIE_ZBIORCZE.xlsx z trzema typami arkuszy:
    - 'Zestawienie'  — wiersz na wieś (pow. rozliczona, przybyło, ubyło, saldo)
                      + wiersz RAZEM z sumami wszystkich wsi,
    - 'Przybyło'     — rozpiska WSZYSTKICH działek z przybyłem, per wieś
                      (J. rej., nr działki, właściciel, powierzchnie),
    - 'Ubyło'        — analogiczna rozpiska działek z ubyłem.
    Właściciel jest dobierany z arkusza Tabela_Glowna po J. rej. + nr działki.
    Zwraca słownik z podsumowaniem albo None (brak plików rozliczeń).
    """
    from openpyxl.styles import Font as _Font, Alignment as _Alignment, PatternFill as _Fill
    from openpyxl.utils import get_column_letter as _gcl

    folder = Path(folder)
    pliki = sorted(folder.glob("*_Rozliczone.xlsx"))
    if not pliki:
        return None

    def _norm(v):
        """Ujednolica wartość (liczba/tekst) do klucza porównawczego."""
        try:
            if v is not None and pd.notna(v) and str(v).strip() != "":
                return ("n", round(float(str(v).replace(",", ".")), 4))
        except (ValueError, TypeError):
            pass
        return ("s", str(v).strip() if v is not None else "")

    wiersze = []
    szczegoly = {"PRZYBYLO": [], "UBYLO": []}   # lista (wieś, DataFrame z właścicielem)

    for sciezka in pliki:
        wies = sciezka.stem
        if wies.lower().endswith("_rozliczone"):
            wies = wies[: -len("_rozliczone")] if wies.endswith("_rozliczone") else wies[:-len("_Rozliczone")]

        # Tabela_Glowna -> suma ROZLICZONE + mapa właścicieli
        tg = None
        wlasciciele = {}
        try:
            tg = pd.read_excel(sciezka, sheet_name="Tabela_Glowna")
        except Exception:
            tg = None
        if tg is not None and "właściciel" in getattr(tg, "columns", []):
            for _, row in tg.iterrows():
                klucz = (_norm(row.get("J. rej.")), _norm(row.get("nr_dz")))
                wl = str(row.get("właściciel") or "").strip()
                if wl and klucz not in wlasciciele:
                    wlasciciele[klucz] = wl
        rozliczona = 0.0
        if tg is not None and "ROZLICZONE" in getattr(tg, "columns", []):
            rozliczona = float(pd.to_numeric(tg["ROZLICZONE"], errors="coerce").fillna(0).sum())

        # PRZYBYLO / UBYLO — pełna treść (nagłówki w wierszu 2, startrow=1)
        for nazwa, kolumna_ile in (("PRZYBYLO", "ile przybyło"), ("UBYLO", "ile ubyło")):
            try:
                df = pd.read_excel(sciezka, sheet_name=nazwa, header=1)
            except Exception:
                df = None
            if df is None or df.empty or kolumna_ile not in getattr(df, "columns", []):
                continue
            ma_wlasciciela = "właściciel" in getattr(df, "columns", [])
            wl = []
            for _, row in df.iterrows():
                klucz = (_norm(row.get("J. rej.")), _norm(row.get("nr działki")))
                z_tg = wlasciciele.get(klucz, "")
                if ma_wlasciciela:
                    # kolumna już istnieje (plik zredagowany ręcznie):
                    # wartość z pliku ma pierwszeństwo, braki uzupełniamy z Tabela_Glowna
                    z_pliku = row.get("właściciel")
                    z_pliku = "" if pd.isna(z_pliku) else str(z_pliku).strip()
                    wl.append(z_pliku if z_pliku else z_tg)
                else:
                    wl.append(z_tg)
            df = df.copy()
            df["właściciel"] = wl
            # właściciel ZAWSZE jako ostatnia kolumna — wąska, tekst w jednej linii
            # wychodzi poza jej krawędź (nic go nie zasłania), jak po wyłączeniu
            # zawijania w Excelu
            kolej = [c for c in df.columns if c != "właściciel"] + ["właściciel"]
            df = df[kolej]
            szczegoly[nazwa].append((wies, df, kolumna_ile))

        # sumy do arkusza 'Zestawienie'
        def _sumuj(nazwa, kolumna):
            for w, df, kol in szczegoly[nazwa]:
                if w == wies and kol == kolumna:
                    return (float(pd.to_numeric(df[kolumna], errors="coerce").fillna(0).sum()),
                            int(df[kolumna].notna().sum()))
            return 0.0, 0

        p_ha, p_n = _sumuj("PRZYBYLO", "ile przybyło")
        u_ha, u_n = _sumuj("UBYLO", "ile ubyło")
        wiersze.append({
            "Wieś": wies,
            "Pow. rozliczona [ha]": round(rozliczona, 4),
            "Przybyło [ha]": round(p_ha, 4),
            "Przybyło działek": p_n,
            "Ubyło [ha]": round(u_ha, 4),
            "Ubyło działek": u_n,
        })

    if not wiersze:
        return None

    df = pd.DataFrame(wiersze)
    razem = {"Wieś": "RAZEM (wszystkie wsie)"}
    razem["Pow. rozliczona [ha]"] = round(df["Pow. rozliczona [ha]"].sum(), 4)
    razem["Przybyło [ha]"] = round(df["Przybyło [ha]"].sum(), 4)
    razem["Przybyło działek"] = int(df["Przybyło działek"].sum())
    razem["Ubyło [ha]"] = round(df["Ubyło [ha]"].sum(), 4)
    razem["Ubyło działek"] = int(df["Ubyło działek"].sum())
    df = pd.concat([df, pd.DataFrame([razem])], ignore_index=True)

    sciezka_out = folder / "ZESTAWIENIE_ZBIORCZE.xlsx"
    with pd.ExcelWriter(str(sciezka_out), engine="openpyxl") as writer:
        # ---------- arkusz zbiorczy ----------
        df.to_excel(writer, sheet_name="Zestawienie", index=False)
        ws = writer.sheets["Zestawienie"]
        for i, c in enumerate(df.columns, start=1):
            ws.column_dimensions[_gcl(i)].width = max(14, min(28, len(c) + 3))
            kom = ws.cell(row=1, column=i)
            kom.font = _Font(bold=True)
            kom.alignment = _Alignment(horizontal="center")
        for i in range(1, len(df.columns) + 1):
            ws.cell(row=len(df), column=i).font = _Font(bold=True)
        for r in range(2, len(df) + 1):
            for i, c in enumerate(df.columns, start=1):
                if "[ha]" in str(c):
                    ws.cell(row=r, column=i).number_format = "0.0000"

        # ---------- arkusze 'Przybyło' / 'Ubyło' (rozpiska działek) ----------
        for nazwa, nazwa_ark in (("PRZYBYLO", "Przybyło"), ("UBYLO", "Ubyło")):
            sekcje = szczegoly[nazwa]
            if not sekcje:
                continue
            kolumna_ile = "ile przybyło" if nazwa == "PRZYBYLO" else "ile ubyło"
            szer = [10, 13, 40, 15, 15, 13, 12]
            wypel = _Fill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") \
                if nazwa == "PRZYBYLO" else _Fill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

            ws = writer.book.create_sheet(nazwa_ark)
            wiersz = 1
            for wies, df_w, _ in sekcje:
                suma = float(pd.to_numeric(df_w[kolumna_ile], errors="coerce").fillna(0).sum())
                nagl = f"{wies} — {kolumna_ile}: {suma:.4f} ha, działek: {len(df_w)}"
                ws.merge_cells(start_row=wiersz, start_column=1, end_row=wiersz, end_column=7)
                kom = ws.cell(row=wiersz, column=1, value=nagl)
                kom.font = _Font(bold=True, size=12)
                kom.fill = wypel
                kom.alignment = _Alignment(horizontal="left", vertical="center")
                wiersz += 1
                for i, c in enumerate(df_w.columns, start=1):
                    kom = ws.cell(row=wiersz, column=i, value=str(c))
                    kom.font = _Font(bold=True)
                    kom.alignment = _Alignment(horizontal="center")
                wiersz += 1
                for _, row in df_w.iterrows():
                    for i, c in enumerate(df_w.columns, start=1):
                        v = row[c]
                        if pd.isna(v):
                            v = ""
                        kom = ws.cell(row=wiersz, column=i, value=v)
                        if c == "właściciel":
                            # bez zawijania: wąska kolumna na końcu, tekst w jednej linii
                            # swobodnie wychodzi poza jej prawą krawędź
                            kom.alignment = _Alignment(horizontal="left", vertical="top")
                        else:
                            kom.alignment = _Alignment(vertical="top")
                            if ("pow" in str(c)) or ("ls" in str(c)) or (c == kolumna_ile):
                                kom.number_format = "0.0000"
                    wiersz += 1
                wiersz += 1  # pusty wiersz między wsiami

            # wiersz RAZEM na końcu rozpiski
            suma_cala = sum(
                float(pd.to_numeric(df_w[kolumna_ile], errors="coerce").fillna(0).sum())
                for _, df_w, _ in sekcje)
            n_dz = sum(len(df_w) for _, df_w, _ in sekcje)
            ws.merge_cells(start_row=wiersz, start_column=1, end_row=wiersz, end_column=7)
            kom = ws.cell(row=wiersz, column=1,
                          value=f"RAZEM — {kolumna_ile}: {suma_cala:.4f} ha, działek: {n_dz}")
            kom.font = _Font(bold=True, size=12)
            for i, wdt in enumerate(szer, start=1):
                ws.column_dimensions[_gcl(i)].width = wdt
            ws.freeze_panes = "A2"

    return {"plik": str(sciezka_out), "wsie": len(wiersze), "razem": razem}

def _read_dbf_records(sciezka):
    """Odczyt DBF (dBase III) jak MIETEK — odporny na 'śmieci' w polach liczbowych.

    Zwraca listę słowników {nazwa_pola: wartość_tekstowa}; pól N nie konwertujemy
    od razu na liczby (bywają uszkodzone) — parsowanie odbywa się przy użyciu.
    """
    import struct
    with open(sciezka, "rb") as f:
        header = f.read(32)
        if len(header) < 32:
            raise Exception(f"Za krótki nagłówek DBF: {sciezka}")
        num_records = struct.unpack("<I", header[4:8])[0]
        header_length = struct.unpack("<H", header[8:10])[0]
        record_length = struct.unpack("<H", header[10:12])[0]
        pola = []
        while True:
            fld = f.read(32)
            if len(fld) < 32 or fld[0] == 0x0D:
                break
            pola.append((fld[0:11].split(b"\x00", 1)[0].decode("ascii", "replace"),
                         chr(fld[11]), fld[16]))
        f.seek(header_length)
        rekordy = []
        for _ in range(num_records):
            rec_raw = f.read(record_length)
            if len(rec_raw) < record_length:
                break
            rec = {}
            off = 1  # bajt flagi usunięcia
            for (nazwa, typ, dl) in pola:
                raw = rec_raw[off:off + dl]
                off += dl
                if typ in ("C", "M", "G"):
                    rec[nazwa] = raw.decode("cp852", "replace").strip()
                elif typ in ("N", "F"):
                    rec[nazwa] = raw.decode("ascii", "replace").replace("\x00", "").strip()
                else:
                    rec[nazwa] = raw.decode("ascii", "replace").strip()
            rekordy.append(rec)
    return rekordy


def _znajdz_dbf(folder_obrebu, litera):
    """Pierwszy <litera>*.DBF w obrębie (bez wielkości liter, pomijając WSIE.DBF)."""
    wyniki, widziane = [], set()
    for wzor in (f"{litera}*.DBF", f"{litera}*.dbf",
                 f"{litera.lower()}*.DBF", f"{litera.lower()}*.dbf"):
        for p in folder_obrebu.rglob(wzor):
            if p.stem.upper() == "WSIE":
                continue
            klucz = str(p).upper()
            if klucz not in widziane:
                widziane.add(klucz)
                wyniki.append(p)
    return wyniki[0] if wyniki else None


def _liczba_dbf(v, domyslna=0.0):
    """Bezpieczne parsowanie liczby z DBF/Excel — śmieci dają wartość domyślną."""
    try:
        if v is None or str(v).strip() == "":
            return domyslna
        return float(str(v).replace(",", "."))
    except (ValueError, TypeError):
        return domyslna


def zrob_zestawienie_z_mietkow(mietki_dir):
    """Zestawienie zbiorcze z mietków z wpisanymi krzyżówkami (D*.DBF).

    Dla każdego folderu-obrębu w folderze mietków sumuje powierzchnie rozliczone
    z krzyżówek i zapisuje ZESTAWIENIE_Z_MIETKOW.xlsx: wiersz na każdą wieś
    + wiersz RAZEM (suma wszystkich wsi).
    Zwraca słownik {plik, wsie, razem} albo None (brak krzyżówek w folderze).
    """
    from openpyxl.styles import Font as _Font, Alignment as _Alignment
    from openpyxl.utils import get_column_letter as _gcl

    mietki_dir = Path(mietki_dir)
    if not mietki_dir.is_dir():
        return None

    wiersze = []
    for obreb in sorted(f for f in mietki_dir.iterdir() if f.is_dir()):
        d_path = _znajdz_dbf(obreb, "D")
        if d_path is None:
            continue
        try:
            rekordy = _read_dbf_records(d_path)
        except Exception:
            continue
        if not rekordy:
            continue
        suma = round(sum(_liczba_dbf(r.get("POW")) for r in rekordy), 4)
        wiersze.append({"Wieś": obreb.name, "Pow. rozliczona [ha]": suma})

    if not wiersze:
        return None

    df = pd.DataFrame(wiersze)
    razem_suma = round(df["Pow. rozliczona [ha]"].sum(), 4)
    df = pd.concat(
        [df, pd.DataFrame([{"Wieś": "RAZEM (wszystkie wsie)",
                            "Pow. rozliczona [ha]": razem_suma}])],
        ignore_index=True)

    sciezka_out = mietki_dir / "ZESTAWIENIE_Z_MIETKOW.xlsx"
    with pd.ExcelWriter(str(sciezka_out), engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Zestawienie", index=False)
        ws = writer.sheets["Zestawienie"]
        for i, c in enumerate(df.columns, start=1):
            ws.column_dimensions[_gcl(i)].width = max(14, min(28, len(str(c)) + 3))
            kom = ws.cell(row=1, column=i)
            kom.font = _Font(bold=True)
            kom.alignment = _Alignment(horizontal="center")
        for i in range(1, len(df.columns) + 1):
            ws.cell(row=len(df), column=i).font = _Font(bold=True)
        for r in range(2, len(df) + 1):
            ws.cell(row=r, column=2).number_format = "0.0000"

    return {"plik": str(sciezka_out), "wsie": len(wiersze), "razem": razem_suma}
