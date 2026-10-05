# -*- coding: utf-8 -*-
"""
Forestly — schemat interfejsu webowego (PyWebView).

Jedno źródło prawdy: frontend (index.html/app.js) renderuje zakładki i
kontrolki na podstawie tego schematu, a backend (web_backend.py) tworzy
na jego podstawie "sztuczne widgety" (FakeEntry/FakeVar), które czyta
istniejąca logika mixinów. Dzięki temu każda funkcja z GUI CustomTkinter
ma swoje odwzorowanie 1:1.

Rodzaje kontrolek (kind):
  path     — pole na ścieżkę (folder=folder, file=plik, save=zapis pliku)
  text     — pole tekstowe
  check    — checkbox (bool)
  checks   — grupa checkboxów z opcją "Wszystkie" (wzajemne wykluczanie)
  select   — lista rozwijana
  margins  — tabela marginesów (10 typów plików × 4 strony)
  fonts    — tabela rozmiarów czcionek arkuszy Excel
  dashboard— kafelki postępu (tylko Pełny Automat)
  info     — statyczny tekst informacyjny
"""

from app.config import (CURRENT_VERSION, EXCEL_SHEET_DEFAULTS,
                        PDF_ORDER_TEMPLATES, load_margins)

# ---------------------------------------------------------------- pomocnicze

def _path(cid, attr, label, ph="", save=True, kind="folder"):
    return {"id": cid, "kind": "path", "attr": attr, "label": label,
            "ph": ph, "browse": kind, "save": save}


def _text(cid, attr, label, default="", ph=""):
    return {"id": cid, "kind": "text", "attr": attr, "label": label,
            "default": default, "ph": ph, "save": True}


def _textarea(cid, attr, label, default="", ph="", rows=4):
    return {"id": cid, "kind": "textarea", "attr": attr, "label": label,
            "default": default, "ph": ph, "save": True, "rows": rows}


def _check(cid, attr, label, default=False, tooltip=""):
    return {"id": cid, "kind": "check", "attr": attr, "label": label,
            "default": default, "tooltip": tooltip, "save": True}


def _checks(cid, base, label, choices, tooltip=""):
    return {"id": cid, "kind": "checks", "attr_base": base, "label": label,
            "choices": choices, "tooltip": tooltip, "save": True}


def _group(label, controls, tooltip="", collapsed=True):
    return {"kind": "group", "label": label, "controls": controls,
            "tooltip": tooltip, "collapsed": collapsed}


def _select(cid, attr, label, values, default, free=False):
    return {"id": cid, "kind": "select", "attr": attr, "label": label,
            "values": values, "default": default, "free": free, "save": True}


def _strtyt(cid, attr, label, values, default, imgs):
    """v2.0.111: wybór szablonu strony tytułowej z podglądami (B/D)."""
    return {"id": cid, "kind": "strtyt", "attr": attr, "label": label,
            "values": values, "default": default, "save": True, "imgs": imgs}


def _info(text_):
    return {"kind": "info", "text": text_}


def _margins(mode):
    return {"id": f"margins_{mode}", "kind": "margins", "mode": mode,
            "label": "Ustawienia marginesów (w cm):"}

# typy raportów nowego wyglądu + lista czcionek do wyboru
CZCIONKI_TYPY = ["REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1", "WSK_ZB",
                 "ZEST1", "HALIZNY", "WYK_NEG", "SKROTY"]
CZCIONKI_LISTA = ["", "Arial", "Times New Roman", "Calibri", "Verdana",
                  "Tahoma", "Georgia", "Trebuchet MS", "Courier New"]

def _czcionki(mode):
    return {"id": f"czcionki_{mode}", "kind": "czcionki", "mode": mode,
            "label": "Ustawienia czcionek (tytuł dokumentu / tekst w tabeli):",
            "types": CZCIONKI_TYPY, "fonts": CZCIONKI_LISTA,
            "tooltip": ("Rozmiar (pt) i rodzaj czcionki — osobno dla tytułu "
                        "dokumentu i osobno dla tekstu w tabelach, dla każdego "
                        "typu raportu oraz wykazu skrótów i symboli "
                        "(SKROTY — tam tabela ma 8,4 pt). Nazwa obiektu, "
                        "\"Stan na\" i AGENCJA zostają bez zmian. "
                        "Domyślne wartości zachowują dzisiejszy wygląd.")}


def _button(bid, label, task, style="primary", tooltip="", group=None):
    b = {"id": bid, "label": label, "task": task, "style": style,
         "tooltip": tooltip}
    if group:
        b["group"] = group
    return b


def _info_grupa(text_, grupa):
    """Nagłówek opisujący tryb przełącznika (należy do jego grupy)."""
    return {"kind": "info", "text": text_, "group": grupa}


def _segment(cid, label, options, default, tooltip="", attr=None,
             sel_class=None, groups=None):
    """Przełącznik segmentowy. Dwa tryby działania:
      * z 'groups' — przełącza widoczność grup kontrolek (np. tryb
        'Jedna wieś / Wiele wsi' pokazuje grupę 'jedna' albo 'wiele'),
      * z 'attr' — zwykły przełącznik wartości (np. PDF/Word), którego wybór
        jest zapamiętywany i wysyłany do backendu jak każde inne pole.
    'sel_class' pozwala pokolorować wybraną opcję (np. PDF na czerwono).
    WAŻNE: 'groups' trzeba podać jawnie — zwykły przełącznik wartości
    NIE może przełączać widoku zakładki (inaczej chowa kontrolki)."""
    c = {"id": cid, "kind": "segment", "label": label, "options": options,
         "default": default, "tooltip": tooltip}
    if groups:
        c["groups"] = groups
    if attr:
        c["attr"] = attr
        c["save"] = True
    if sel_class:
        c["sel_class"] = sel_class
    return c


def _grp(ctrl, grupa):
    """Kontrolka przypisana do trybu przełącznika segmentu."""
    c = dict(ctrl)
    c["group"] = grupa
    return c


WYDRUKI_CHOICES = ["Wszystkie", "OPTAX", "TAB_KLW3", "ZEST1", "REJESTR1",
                   "WSKAZ1", "WYK_NEG", "WSK_ZB", "HALIZNY"]
WORD_CHOICES = ["Wszystkie", "REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1",
                "HALIZNY", "WYK_NEG", "OPIS", "ZEST1", "WK_ZM1"]

DASHBOARD_STEPS = ["Halizny + TXT", "Czyszczenie", "Word", "PDF", "Scalanie"]


# ------------------------------------------------------------------- budowa

def build_schema():
    from app.config import TERRITORY_DATA
    woj_list = sorted(TERRITORY_DATA.keys()) if TERRITORY_DATA else ["BRAK DANYCH"]
    default_woj = "KUJAWSKO-POMORSKIE" if "KUJAWSKO-POMORSKIE" in TERRITORY_DATA else (woj_list[0] if woj_list else "")
    powiat_list = sorted(TERRITORY_DATA.get(default_woj, {}).keys()) or ["BRAK DANYCH"]
    default_powiat = "TUCHOLSKI" if "TUCHOLSKI" in TERRITORY_DATA.get(default_woj, {}) else (powiat_list[0] if powiat_list else "")
    gmina_list = list(TERRITORY_DATA.get(default_woj, {}).get(default_powiat, []))

    fonts = [
        {"sheet": s, "start_row": r, "default": f}
        for (s, r, f) in EXCEL_SHEET_DEFAULTS if s != "Sheet4"
    ]

    tabs = [
        # ================================================================ MIETEK
        {
            "key": "MIETEK|Mietek v2.0 — edytor danych",
            "tooltip": "Przeglądaj i edytuj dane mietka (pliki DBF: W/R/O/D/Z/WSIE) "
                       "w wygodnej siatce — zapis wprost do DBF z backupem .BAK. "
                       "Dokumenty (OPTAX, REJESTR1, WSKAZ1…) generują się nowymi "
                       "szablonami FORESTLY bezpośrednio z danych — bez Worda "
                       "(STR_TYT — z ustawień kreatora).",
            "controls": [
                {"id": "mietki_v2", "kind": "mietek_v2",
                 "label": "Edytor danych mietka"},
            ],
            "buttons": [],
        },

        {
            "key": "MIETEK|Pełny Automat (1-Click)",
            "tooltip": "Kompleksowy proces: halizny, TXT z DBF, Word, PDF i scalanie w gotowy dokument.",
            "controls": [
                _path("all_src", "entries.ALL.src", "Folder źródłowy:", "Wskaż folder..."),
                _path("all_dst", "entries.ALL.dst", "Folder docelowy:", "Wskaż lokalizację..."),
                _group(
                    "Kreator strony tytułowej (strona tytułowa i daty)",
                    [
                        _strtyt(
                            "all_tpl_szablon", "all_tpl_szablon_var",
                            "Szablon strony tytułowej:",
                            ["Wersja 1", "Wersja 2", "Wersja 3"],
                            "Wersja 1",
                            [{"src": "podglad_STR_TYT_1.png",
                              "label": "Wersja 1 — wbudowana",
                              "value": "Wersja 1"},
                             {"src": "podglad_STR_TYT_2.png",
                              "label": "Wersja 2 — klasyczna",
                              "value": "Wersja 2"},
                             {"src": "podglad_STR_TYT_3.png",
                              "label": "Wersja 3 — minimalna",
                              "value": "Wersja 3"}]),
                        _select("all_tpl_doc", "all_tpl_doc_var", "Typ dokumentu:",
                                ["UPUL", "ISL"], "UPUL"),
                        _select("all_tpl_prefix", "all_tpl_prefix_var", "Prefiks obrębu:",
                                ["położonych na terenie obrębu", "Obręb:"],
                                "położonych na terenie obrębu"),
                        _select("all_tpl_woj", "all_tpl_woj_var", "Województwo:",
                                woj_list, default_woj, free=True),
                        _select("all_tpl_powiat", "all_tpl_powiat_var", "Powiat (można wpisać własny):",
                                powiat_list, default_powiat, free=True),
                        _select("all_tpl_gmina", "all_tpl_gmina_var", "Gmina (można wpisać własną):",
                                gmina_list, gmina_list[0] if gmina_list else "", free=True),
                        _text("all_tpl_stan", "all_tpl_stan_na_entry",
                              "Stan na (także data we wszystkich Wordach):",
                              "30.06.2026 r.", "np. 30.06.2026 r."),
                        _text("all_tpl_okres", "all_tpl_okres_entry",
                              "Na okres (strona tytułowa):",
                              "01.01.2027 – 31.12.2036 r.", "np. 01.01.2027 – 31.12.2036 r."),
                        _text("all_wsk_od", "all_wsk_od_entry",
                              "WSK_ZB — 10-lecie od:", "01-01-2027", "np. 01-01-2027"),
                        _text("all_wsk_do", "all_wsk_do_entry",
                              "WSK_ZB — 10-lecie do:", "31-12-2036", "np. 31-12-2036"),
                    ],
                    tooltip=('Trzy wersje strony tytułowej. „Wersja 1” buduje '
                             'się z pól poniżej; „Wersja 2” i „Wersja 3” to '
                             'Twoje pliki (STR_TYT_wersja_2.docx / '
                             'STR_TYT_wersja_3.docx w folderze programu) — '
                             'nazwa wsi i powierzchnia podstawia się z OPTAX, '
                             'a gmina/powiat/województwo i daty z pól poniżej.\n'
                             'Klik na miniaturkę wybiera wersję, lupa (🔍) '
                             'powiększa podgląd.\n'
                             '„Stan na” zastępuje daty we wszystkich generowanych '
                             'dokumentach Word, a pola 10-lecia — okres w WSK_ZB.'), collapsed=False),
                _check("all_pelny_opis_og", "all_pelny_opis_og_var",
                       "Pełne opisy ogólne (z powiązaniami z gospodarką leśną "
                       "i opisami form)", True,
                       "Pełny opis ogólny (opis og_<wieś>.docx) — z powiązaniami "
                       "obszarów Natura 2000 z gospodarką leśną i opisami "
                       "pozostałych form ochrony przyrody.\n"
                       "Dane liczbowe pochodzą z WSK_ZB.doc wsi."),
                _check("all_krotki_opis_og", "all_krotki_opis_og_var",
                       "Skrócone opisy ogólne (sama lista form ochrony przyrody)", False,
                       "Opis ogólny w wersji skróconej — same zdania typu "
                       "„Obszar Natura 2000 SOO … PLH300032 w pododdziałach …”, "
                       "bez powiązań z gospodarką leśną i bez opisów form.\n"
                       "Zaznacz ALBO to, ALBO pełne (wzajemnie się wykluczają); "
                       "oba odznaczone = opisy ogólne nie powstają."),
                _path("all_gdos", "all_gdos_entry",
                      "Folder z wynikami GDOŚ (opcjonalny):",
                      "Formy ochrony przyrody (np. NN_WIEŚ_wynik.xlsx) do opisów ogólnych"),
                _textarea(
                    "all_og_nadzor", "all_og_nadzor_entry",
                    "1. NADZÓR — treść (można edytować):",
                    "Nadzór nad gospodarką leśną lasów nie stanowiących własności "
                    "Skarbu Państwa sprawuje Starosta Wołomiński w zakresie "
                    "zadań własnych.", rows=2),
                _textarea(
                    "all_og_warunki", "all_og_warunki_entry",
                    "2. WARUNKI PRZYRODNICZE — treść (można edytować):",
                    "Lasy objęte uproszczonym planem urządzenia lasów położone są w:\n"
                    "IV Mazowiecko-Podlaskiej krainie przyrodniczo-leśnej\n"
                    "Mezoregion Doliny Dolnego Bugu", rows=4),
                _select("all_og_kategoria", "all_og_kategoria_var",
                        "Kategoria zagrożenia pożarowego:",
                        ["I", "II", "III"], "I"),
                _select("all_og_tabela", "all_og_tabela_var",
                        "Tabela siedliskowa w opisie ogólnym:",
                        ["Mazowiecka", "Wielkopolska"], "Mazowiecka"),
                _check("all_custom_skroty", "all_custom_skroty_var",
                       "Użyj własnego pliku 'Skróty i symbole' (zamiast domyślnego z programu)", False),
                _path("all_skroty", "all_skroty_entry", "Własny plik:",
                      "Wskaż własny plik ze skrótami...", kind="file"),
                _check("all_kontrola", "all_kontrola_var",
                       "Kontrola powierzchni REJESTR ↔ OPTAX (dodatkowy PDF)", False,
                       "Obok scalonego pakietu powstanie KONTROLA_<wieś>.pdf: "
                       "sumy powierzchni Rejestru i Opisu taksacyjnego oraz wykaz "
                       "wydzieleń i działek z rozbieżnościami. Mietek pozostaje "
                       "bez zmian — to wyłącznie kontrola."),
                _check("all_obie_wersje", "all_obie_wersje_var",
                       "Obie wersje REJESTRU — dwa foldery wynikowe", False,
                       "Uruchamia Pełny Automat dwukrotnie: raz z pełnymi "
                       "nazwiskami (folder 'Z nazwiskami'), raz bez nich "
                       "(folder 'Bez nazwisk').\nIgnoruje powyższy przełącznik "
                       "usuwania nazwisk."),
                _path("all_mapa", "all_mapa_entry",
                      "Folder z mapami (jpg/png/tiff, opcjonalnie):",
                      "Wszystkie mapy z folderu — dopasuję do wsi po nazwie pliku",
                      kind="folder"),
                _check("remove_names", "remove_names_var",
                       "Usuwaj nazwiska z REJESTRU (oraz 1. stronę)", True,
                       "Włączenie tej opcji uruchamia makra 'ZamienLF' oraz 'UsunNazwiskaRej', "
                       "a także kasuje pierwszą stronę z rejestru."),
                _margins("ALL"),
                _czcionki("ALL"),
                {"kind": "dashboard", "id": "dashboard", "steps": DASHBOARD_STEPS},
            ],
            "buttons": [
                _button("run", "▶  Rozpocznij proces", "start_pipeline:ALL"),
                _button("order", "Skonfiguruj układ PDF", "open_order:ALL", "secondary"),
            ],
        },
        {
            "key": "MIETEK|Nowe Szablony",
            "tooltip": ("Nowe szablony wydruków (HTML → PDF, bez Worda). Dwa źródła "
                        "danych: pliki TXT / DBF mietka albo stare pliki Word "
                        "z Pełnego Automatu."),
            "controls": [
                _segment("ns_zrodlo", "Źródło danych:",
                         ["Pliki TXT / DBF mietka", "Stare pliki Word (Pełny Automat)"],
                         "Pliki TXT / DBF mietka",
                         tooltip="Z mietka (DBF → TXT → PDF) albo ze starych plików Word.",
                         groups={"Pliki TXT / DBF mietka": "txt",
                                 "Stare pliki Word (Pełny Automat)": "word"}),
                # ---------- tryb: pliki TXT / DBF mietka ----------
                _grp(_path("ns_src", "entries.NS.src", "Folder z Mietkami (obręby):",
                           "Folder, w którym leżą foldery obrębów (np. CHORZEWO\\WOL.001\\...DBF)"), "txt"),
                _grp(_path("ns_dst", "entries.NS.dst", "Folder docelowy (PDF):",
                           "Wskaż lokalizację..."), "txt"),
                _grp(_select("ns_typ", "ns_typ_var", "Raport do wygenerowania:",
                             ["WYK_NEG", "REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1",
                              "WSK_ZB", "ZEST1", "HALIZNY"], "WYK_NEG"), "txt"),
                _grp(_check("ns_bez_nazwisk", "ns_bez_nazwisk_var",
                            "Bez nazwisk (dotyczy REJESTR1 i WSKAZ1)", True,
                            "Działa tak samo jak opcja 'Usuwaj nazwiska' w Pełnym Automacie."), "txt"),
                _grp(_margins("NS"), "txt"),
                _grp(_czcionki("NS"), "txt"),
                _grp(_info("Program sam wygeneruje pliki TXT wybranego raportu z DBF-ów "
                           "każdego obrębu (jak zakładka 'MIETEK -> TXT'; pliki TXT powstaną "
                           "obok DBF-ów) i zamieni je na PDF-y nowym wyglądem."), "txt"),
                # ---------- tryb: stare pliki Word ----------
                _info_grupa("Przerabia WSZYSTKIE stare pliki Word z Pełnego Automatu "
                            "(OPTAX, REJESTR1, TAB_KLW3, WSK_ZB, WSKAZ1, WYK_NEG, ZEST1, "
                            "HALIZNY, „opis og_<wieś>”, STR_TYT) na nowe szablony. "
                            "Oryginały zostają nietknięte; puste raporty są pomijane; "
                            "pliki .doc konwertuje tymczasowo Word.", "word"),
                _grp(_path("stare_root", "stare_opisy_root_entry",
                           "1. Folder ze starymi plikami:",
                           "np. folder „ułożone” z podfolderami wsi albo folder jednej wsi"), "word"),
                _grp(_path("stare_out", "stare_opisy_out_entry",
                           "2. Folder docelowy zapisu:",
                           "Gdzie zapisać nowe pliki? (puste = obok oryginałów)"), "word"),
                _grp(_strtyt(
                    "stare_strtyt_szablon", "stare_strtyt_szablon_var",
                    "3. Szablon strony tytułowej:",
                    ["Wersja 1", "Wersja 2", "Wersja 3"],
                    "Wersja 1",
                    [{"src": "podglad_STR_TYT_1.png",
                      "label": "Wersja 1 — wbudowana", "value": "Wersja 1"},
                     {"src": "podglad_STR_TYT_2.png",
                      "label": "Wersja 2 — klasyczna", "value": "Wersja 2"},
                     {"src": "podglad_STR_TYT_3.png",
                      "label": "Wersja 3 — minimalna", "value": "Wersja 3"}]), "word"),
            ],
            "buttons": [
                _button("run", "▶  Generuj PDF", "start_nowe_szablony", group="txt"),
                _button("run_word", "Przerób na nowy szablon", "start_stare_opisy", group="word"),
            ],
        },
        {
            "key": "MIETEK|Generowanie: MIETEK -> TXT",
            "tooltip": "Generuje pliki wydrukowe MIETEKA bezpośrednio z plików DBF.",
            "controls": [
                _path("wydruki_src", "wydruki_mietki_entry", "Folder z Mietkami (obręby):",
                      "Folder, w którym leżą foldery obrębów (np. CHORZEWO\\WOL.001\\...DBF)"),
                _checks("wydruki_filters", "wydruki_filter_vars", "Generuj tylko:",
                        WYDRUKI_CHOICES,
                        "Możesz zaznaczyć wiele plików naraz. Opcja 'Wszystkie' wyklucza pozostałe.\n"
                        "HALIZNY.TXT zaznaczaj tylko PRZED przeniesieniem halizn (zakładka 'Halizny')."),
                _info("Pliki TXT trafiają tam, gdzie pliki DBF obrębu. HALIZNY.TXT generuj "
                      "PRZED przeniesieniem halizn, pozostałe wydruki PO (zakładka 'Halizny'). "
                      "Dalej przetwarzaj je zakładką 'Konwersja: MIETEK -> Word'."),
            ],
            "buttons": [_button("run", "Generuj wydruki TXT — zaznaczone w 'Generuj tylko:'",
                               "start_wydruki_all")],
        },
        {
            "key": "MIETEK|Konwersja: MIETEK -> Word",
            "tooltip": "Tylko etap 1: Oczyszcza surowe pliki z systemu MIETEK i układa pliki Word.",
            "controls": [
                _path("word_src", "entries.WORD.src", "Folder źródłowy (TXT):", "Wskaż folder..."),
                _path("word_dst", "entries.WORD.dst", "Folder docelowy (Word):", "Wskaż lokalizację..."),
                _check("word_remove_names", "remove_names_var",
                       "Usuwaj nazwiska z REJESTRU (oraz 1. stronę)", True,
                       "Włączenie tej opcji uruchamia makra 'ZamienLF' oraz 'UsunNazwiskaRej', "
                       "a także kasuje pierwszą stronę z rejestru."),
                _checks("word_filters", "word_filter_vars", "Konwertuj tylko:", WORD_CHOICES,
                        "Możesz zaznaczyć wiele typów plików naraz. Opcja 'Wszystkie' wyklucza pozostałe."),
                _margins("WORD"),
            ],
            "buttons": [_button("run", "▶  Rozpocznij proces", "start_pipeline:WORD")],
        },
        {
            "key": "MIETEK|Konwersja: Word -> PDF",
            "tooltip": "Tylko etap 2: Zamienia gotowe pliki Word na PDF i łączy w jeden plik.",
            "controls": [
                _path("pdf_src", "entries.PDF.src", "Folder źródłowy (Word):", "Wskaż folder..."),
                _path("pdf_dst", "entries.PDF.dst", "Folder docelowy (PDF):", "Wskaż lokalizację..."),
                _check("pdf_merge", "pdf_merge_var",
                       "Po konwersji scal pliki w jeden dokument PDF", True),
                _check("pdf_skroty", "pdf_skroty_var",
                       "Dołącz 'Skróty i symbole' na końcu scalonego PDF", True,
                       "Dokleja Skróty i symbole (skroty.pdf) na końcu każdego "
                       "scalonego PDF wsi. Wyłącz, jeśli chcesz same dokumenty."),
            ],
            "buttons": [
                _button("run", "▶  Rozpocznij proces", "start_pipeline:PDF"),
                _button("order", "Skonfiguruj układ PDF", "open_order:PDF", "secondary"),
            ],
        },
        {
            "key": "MIETEK|Kreator Stron tytułowych",
            "tooltip": "Strony tytułowe: jedna wieś z formularza (Word + PDF) "
                       "albo masowo dla wielu wsi (dane z plików Word OPTAX).",
            "controls": [
                _segment("kreator_tryb", "Tryb tworzenia:",
                         ["Jedna wieś", "Wiele wsi (masowo)"], "Jedna wieś",
                         groups={"Jedna wieś": "jedna",
                                 "Wiele wsi (masowo)": "wiele"}),
                _grp(_segment("tpl_MIETEK_format", "Format wyjścia:",
                              ["PDF", "Word"], "PDF",
                              attr="tpl_data.MIETEK.format_var",
                              sel_class={"PDF": "sel-pdf", "Word": "sel-word"}),
                     "jedna"),
                *[_grp(c, "jedna") for c in _tpl_controls("MIETEK")],
                _info_grupa("Wiele wsi (masowo) — dane z plików Word (OPTAX):", "wiele"),
                _grp(_path("mt_template", "mietek_title_template_entry", "Szablon STR_TYT:",
                           "Wskaż plik bazowy", kind="file"), "wiele"),
                _grp(_path("mt_word", "mietek_title_word_entry", "Fold. z plikami Word (OPTAX):",
                           "Wskaż folder, w którym znajdują się pliki OPTAX"), "wiele"),
                _grp(_text("mt_village_ph", "mietek_title_village_placeholder_entry",
                           "Placeholder nazwy wsi:", "NAZWA WSI"), "wiele"),
                _grp(_text("mt_area_ph", "mietek_title_area_placeholder_entry",
                           "Placeholder powierzchni:", "wielkość"), "wiele"),
            ],
            "buttons": [
                _button("run", "Wygeneruj stronę tytułową", "generate_str_tyt:MIETEK",
                        group="jedna"),
                _button("run", "Wygeneruj strony tytułowe", "start_mietek_title_pages",
                        group="wiele"),
            ],
        },
        {
            "key": "MIETEK|Opisy ogólne",
            "tooltip": "Tworzy \u201eopis og_<wie\u015b>.docx\u201d w folderach wsi \u2014 liczby czyta z WSK_ZB.doc.",
            "controls": [
                _info("Generator tworzy plik \u201eopis og_<nazwa wsi>.docx\u201d w każdym folderze wsi. "
                      "Jeśli w folderze wsi jest plik WSK_ZB.doc \u2014 liczby czyta z niego. "
                      "Jeśli są tam dane MIETEKA (O*.DBF) \u2014 program robi w folderze "
                      "tymczasowym MIETEK -> TXT -> Word, tworzy opis ogólny i usuwa pliki "
                      "tymczasowe (oryginalne dane zostają nietknięte). Formy ochrony "
                      "przyrody (Natura 2000, parki krajobrazowe) wpisujesz ręcznie "
                      "\u2014 chyba że wskażesz folder z wynikami GDOŚ: wtedy zostaną "
                      "wstawione automatycznie (kody obszarów i publikacje PZO "
                      "z pliku gdos_obszary.json; nieznane obszary będą oznaczone "
                      "do uzupełnienia)."),
                _path("opis_og_root", "opis_og_root_entry",
                      "Folder główny (z folderami wsi) albo folder pojedynczej wsi:",
                      "np. folder \u201eu\u0142o\u017cone\u201d albo folder jednej wsi"),
                _check("opis_og_pelny", "opis_og_pelny_var",
                       "Pełne opisy ogólne (z powiązaniami z gospodarką leśną "
                       "i opisami form)", True,
                       "Pełny opis ogólny (opis og_<wieś>.docx) — z powiązaniami "
                       "obszarów Natura 2000 z gospodarką leśną i opisami "
                       "pozostałych form ochrony przyrody.\n"
                       "Dane liczbowe pochodzą z WSK_ZB.doc wsi."),
                _check("opis_og_krotki", "opis_og_krotki_var",
                       "Skrócone opisy ogólne (sama lista form ochrony przyrody)", False,
                       "Opis ogólny w wersji skróconej — same zdania typu "
                       "„Obszar Natura 2000 SOO … w pododdziałach …”, "
                       "bez powiązań z gospodarką leśną, bez opisów form "
                       "i bez kodów obszarów.\n"
                       "Zaznacz ALBO to, ALBO pełne (wzajemnie się wykluczają); "
                       "oba odznaczone = opisy ogólne nie powstają."),
                _path("opis_og_gdos", "opis_og_gdos_entry",
                      "Folder z wynikami GDOŚ (opcjonalnie):",
                      "np. folder z plikami *_wynik.xlsx — formy ochrony przyrody "
                      "zostaną wstawione automatycznie"),
            ],
            "buttons": [_button("run", "Generuj opisy ogólne", "start_opis_og")],
        },
        {
            "key": "MIETEK|Wykaz Rozbieżności",
            "tooltip": "Porównuje powierzchnie z mietków (D*.DBF) z ewidencją XLS.",
            "controls": [
                _path("rozb_mietki", "mietek_rozb_mietki_entry", "1. Folder z Mietkami (D*.DBF):",
                      "Wskaż folder zawierający pliki D*.DBF (nawet bezpośrednio)"),
                _path("rozb_excel", "mietek_rozb_excel_entry", "2. Folder z plikami XLS (Ewidencja):",
                      "Wskaż folder z ewidencją (.xls/.xlsx)"),
                _path("rozb_out", "mietek_rozb_out_entry", "3. Folder docelowy zapisu:",
                      "Gdzie zapisać gotowe raporty Wykazu Rozbieżności?"),
                _info("Powierzchnie w plikach Excel są odczytywane jako m² i przeliczane na ha (÷10000)."),
            ],
            "buttons": [
                _button("run", "Generuj Wykaz Rozbieżności", "start_rozbieznosci"),
                _button("bez", "Bez Nazwisk", "start_rozbieznosci_bez", "secondary"),
            ],
        },
        {
            "key": "MIETEK|Kontrola powierzchni Rejestr–OPTAX",
            "tooltip": "Raport rozbieżności powierzchni między Rejestrem (D*.DBF) "
                       "a opisem taksacyjnym (OPTAX, O*.DBF) — bez zmieniania "
                       "danych: osobny PDF dla każdej wsi + raport zbiorczy.",
            "controls": [
                _path("kontrola_pow_mietki", "kontrola_pow_mietki_entry",
                      "1. Folder z Mietkami (obręby z D*.DBF / O*.DBF):",
                      "Program przejrzy podfoldery i znajdzie wszystkie obręby"),
                _path("kontrola_pow_out", "kontrola_pow_out_entry",
                      "2. Folder docelowy raportów:",
                      "Gdzie zapisać KONTROLA_<wieś>.pdf i raport zbiorczy?"),
                _info("Raport pokazuje sumy powierzchni obu źródeł oraz wydzielenia "
                      "z różnicą (albo obecne tylko w jednym z nich), wraz z wypisem "
                      "działek: działka, pozycja rejestru, właściciel, powierzchnia — "
                      "żeby błąd dało się namierzyć w mietku. Dane Mietka NIE są "
                      "zmieniane."),
            ],
            "buttons": [
                _button("run", "Generuj raport kontroli", "start_kontrola_pow"),
            ],
        },
        {
            "key": "MIETEK|Czyszczenie rejestru",
            "tooltip": "Usuwa z mietka działki bez przypisanej litery (pododdziału) "
                       "oraz właścicieli bez rozliczonej działki — i odświeża "
                       "wygenerowane raporty (HTML + PDF). Tworzy raport usuniętych.",
            "controls": [
                _info("Wskaż folder z mietkiem — pliki DBF (W/R/O/D/WSIE) mogą "
                      "leżeć w podfolderach, także głębiej (np. CHORZEWO\\WOL.001). "
                      "Program dla każdego obrębu: usuwa działki (D*.DBF) bez "
                      "przypisanego pododdziału (litery wydzielenia, np. „1a” ma "
                      "literę „a”), a następnie usuwa właścicieli / pozycje "
                      "rejestrowe (W*.DBF) bez żadnej rozliczonej działki. "
                      "Przed zapisem powstaje kopia .BAK (pierwsza kopia zostaje "
                      "nietknięta). Pliki O*.DBF i R*.DBF nie są zmieniane."),
                _path("rozl_mietki_root", "rozl_mietki_root_entry",
                      "1. Folder z Mietkiem (pliki DBF w podfolderach):",
                      "np. folder z podfolderami wsi/obrębów, np. CHORZEWO\\WOL.001"),
                _path("rozl_mietki_out", "rozl_mietki_out_entry",
                      "2. Folder z raportami do odświeżenia (opcjonalnie):",
                      "Folder z wygenerowanymi „nowymi szablonami” (HTML + pdf) — "
                      "tu program odświeży REJESTR1 i ZEST1"),
                _check("rozl_mietki_odswiez", "rozl_mietki_odswiez_var",
                       "Odśwież też wygenerowane raporty (HTML + PDF)", True),
                _check("rozl_mietki_dry", "rozl_mietki_dry_var",
                       "Tylko raport — NIE usuwaj danych (podgląd)", False),
            ],
            "buttons": [
                _button("run", "Wyczyść rejestr", "start_czyszczenie_rejestru"),
            ],
        },
        {
            "key": "MIETEK|Opisy na mapę",
            "tooltip": "Wpisywanie opisów do poligonów mapy GEO-MAP z trzech "
                       "baz: Forestly GO (arkusz .xlsx), MIETKA (pliki DBF) "
                       "albo TAKSATORA (.mdb). Oryginał zostaje bez zmian — "
                       "powstaje <NAZWA>_z_opisami.MAP obok mapy źródłowej.",
            "controls": [
                _segment("mapa_zrodlo", "Baza (źródło opisów):",
                         ["Baza Forestly GO", "Baza MIETEK", "Baza TAKSATOR"],
                         "Baza MIETEK",
                         attr="mapa_zrodlo_var",
                         tooltip="Z arkusza Forestly GO, z mietka (DBF) "
                                 "albo z bazy taksatora (.mdb).",
                         groups={"Baza Forestly GO": "mapa_zrodlo:go",
                                 "Baza MIETEK": "mapa_zrodlo:mietek",
                                 "Baza TAKSATOR": "mapa_zrodlo:taksator"}),
                # ---------- Forestly GO ----------
                _grp(_path("mapa_excel", "mapa_excel_entry",
                           "Folder z arkuszami .xlsx (Forestly GO):",
                           "Eksport „Taksacja” — jeden arkusz na wieś/obręb"),
                     "mapa_zrodlo:go"),
                _grp(_text("mapa_kol_nr", "mapa_kol_nr_entry",
                           "Kolumna z numerem porządkowym (Forestly GO):",
                           "N"), "mapa_zrodlo:go"),
                # ---------- MIETEK ----------
                _grp(_path("mapa_mietki", "mapa_mietki_entry",
                           "Folder z Mietkiem (pliki DBF w podfolderach):",
                           "np. folder z podfolderami obrębów (O*.DBF, R*.DBF)"),
                     "mapa_zrodlo:mietek"),
                # ---------- TAKSATOR ----------
                _grp(_path("mapa_mdb", "mapa_mdb_entry",
                           "Baza taksatora (.mdb):",
                           "np. BYSŁAW.mdb", kind="file"),
                     "mapa_zrodlo:taksator"),
                # ---------- wspólne ----------
                _path("mapa_src", "mapa_src_entry",
                      "Folder z mapami (.MAP) albo plik mapy:",
                      "Wskaż folder z mapami — albo pojedynczy plik .MAP",
                      kind="open"),
                _text("mapa_font_mm", "mapa_font_mm_entry",
                      "Wysokość pisma podpisu [mm] (układanie):", "2.5"),
                _text("mapa_skala", "mapa_skala_entry",
                      "Skala mapy (układanie), np. 5000 dla 1:5000:", "5000"),
                _text("mapa_p3", "mapa_p3_entry",
                      "Obrót opisów (P3 w GEO-MAP, w gradach):", "-0.25"),
                _check("mapa_tylko_srodek", "mapa_tylko_srodek_var",
                       "Tylko wyśrodkuj i obróć opisy (BEZ rozsuwania — opisy "
                       "mogą wtedy zasłaniać litery)", False),
                _info("Program wpisuje do mapy opisy z wybranej bazy — "
                      "zależnie od źródła: A2 „Oznaczenie” i A5 „Opis taks.” "
                      "(Forestly GO / MIETEK) albo A1 i A2 „Oznaczenie” "
                      "(TAKSATOR). Wynik zapisuje się obok mapy źródłowej jako "
                      "<NAZWA>_z_opisami.MAP. „Sprawdź braki przed wpisaniem” "
                      "pokazuje tabelę różnic między mapą a regułą i zapisuje je "
                      "do pliku „Opisy na mapę - braki.csv”. „Ułóż opisy "
                      "automatycznie” rozsuwa podpisy na mapie tak, aby się nie "
                      "nakładały (wynik: <NAZWA>_ulozone.MAP)."),
            ],
            "buttons": [
                _button("run", "Wpisz opisy do map", "start_opisy_na_mape"),
                _button("run_braki", "Sprawdź braki przed wpisaniem",
                        "start_sprawdz_opisy_na_mape", "secondary"),
                _button("run_uloz", "Ułóż opisy automatycznie",
                        "start_uloz_opisy", "secondary"),
                _button("run_edytor", "Otwórz edytor opisów (okno)",
                        "open_edytor_opisow", "secondary"),
            ],
        },
        {
            "key": "MIETEK|Mietki +10 lat",
            "tooltip": "Przesuwa wiek w opisach taksacyjnych (zastępuje makro VBA).",
            "controls": [
                _path("plus10_folder", "plus10_folder_entry",
                      "1. Folder z mietkiem (np. z podfolderem WOL.001):",
                      "Program sam znajdzie pliki O*.DBF w podfolderach .001"),
                _text("plus10_lata", "plus10_lata_entry",
                      "2. Przesunięcie wieku (lata):", default="10"),
                _text("plus10_wys", "plus10_wys_entry",
                      "3. Wysokość WYS w R*.DBF (+):", default="1"),
                _text("plus10_piers", "plus10_piers_entry",
                      "4. Pierśnica PIERS w R*.DBF (+):", default="2"),
                _info("O*.DBF: każde /65-75/80 (lub /65-75/80l) w opisach taksacyjnych (OP_TAX) "
                      "dostaje +N lat do wszystkich trzech liczb, np. /65-75/80l → /75-85/90l. "
                      "R*.DBF: pole WIEK +N lat, klasa wieku o pół klasy (IIa→IIb), WYS i PIERS o podane wartości (0 = bez zmian). "
                      "Zawsze: kopia zapasowa .BAK przed zapisem, pola OP_TAX i OP_TAX1. "
                      "Najpierw podgląd zmian, dopiero potem zapis."),
            ],
            "buttons": [
                _button("plus10_podglad", "🔍 Podgląd zmian", "plus10_podglad", "secondary"),
                _button("plus10_zastosuj", "✅ Zastosuj przesunięcie wieku", "plus10_zastosuj"),
            ],
        },
        {
            "key": "MIETEK|NAZWISKA -> MIETEK",
            "tooltip": "Klonuje strukturę MS-DOS i generuje W*.DBF na podstawie Ewidencji XLS.",
            "controls": [
                _path("nm_base", "nazwiska_bazowy_entry", "1. Folder z plikami XLS (Ewidencja):",
                      "Stąd program pobierze nazwy wsi i wszystkich właścicieli..."),
                _path("nm_out", "nazwiska_out_entry", "2. Folder docelowy zapisu:",
                      "Gdzie wygenerować struktury MIETEK (W*.DBF)?"),
                _text("nm_wojew", "nm_wsie_wojew_entry", "Województwo (kod):", "10", "np. 10"),
                _text("nm_powiat", "nm_wsie_powiat_entry", "Powiat:", "", "np. WYSZKOWSKI"),
                _text("nm_stan", "nm_wsie_stan_entry", "Stan na:", "01.01.2023", "DD.MM.RRRR"),
                _text("nm_obod", "nm_wsie_obod_entry", "Obowiązuje od:", "01.01.2023", "DD.MM.RRRR"),
                _text("nm_obdo", "nm_wsie_obdo_entry", "Obowiązuje do:", "31.12.2032", "DD.MM.RRRR"),
                _text("nm_nrws", "nm_wsie_nrws_entry", "Nr wsi:", "1", "np. 1"),
                _text("nm_rokz", "nm_wsie_rokz_entry", "Rok zal.:", "19", "np. 19"),
                _info("Dane nagłówka WSIE.DBF (stałe dla całego uruchomienia). "
                      "NAZWA i GMINA = nazwa obrębu, wpisywane automatycznie."),
            ],
            "buttons": [_button("run", "Generuj struktury (tylko Ewidencja)", "start_nazwiska_mietek")],
        },

        {
            "key": "MIETEK|Baza obszarów GDOŚ",
            "tooltip": "Edytor bazy obszarów ochrony przyrody używanej "
                       "przy generowaniu opisów ogólnych (z wyników GDOŚ).",
            "controls": [
                _info("Baza obszarów ochrony przyrody dla zakładki „Opisy "
                      "ogólne” — uzupełnia kody obszarów, publikacje PZO "
                      "i teksty powiązań, których nie ma w plikach wynikowych "
                      "GDOŚ. Dopisz tu kolejne obszary: nazwa musi być taka sama jak "
                      "w pliku GDOŚ (np. „Ostoja Międzychodzko-Sierakowska”). "
                      "Zmiany obowiązują od następnego generowania opisów. "
                      "W pustych polach edytora pokazują się przykłady."),
                {"id": "gdos_table", "kind": "gdos_table",
                 "label": "Obszary ochrony przyrody:"},
                _path("gdos_xlsx", "gdos_xlsx_entry",
                      "Plik Excel do importu bazy:",
                      "np. gdos_obszary.xlsx (kolumny: Nazwa, Typ, Kod, Publikacja PZO, Powiązanie..., Opis)",
                      kind="file"),
            ],
            "buttons": [
                _button("export", "Eksportuj bazę do Excela", "gdos_export"),
                _button("import", "Importuj bazę z Excela", "gdos_import"),
            ],
        },
        {
            "key": "MIETEK|Ręczne scalanie PDF",
            "tooltip": "Moduł ręczny: wczytaj luźne PDF-y, ustaw kolejność i połącz.",
            "controls": [
                _path("mm_src", "manual_pdf_src", "Wybierz folder PDF:",
                      "Wybierz lokalizację z plikami PDF..."),
                _path("mm_dst", "manual_pdf_dst", "Wybierz folder docelowy:",
                      "Gdzie zapisać plik wynikowy?"),
            ],
            "buttons": [_button("run", "Zarządzaj układem i scal pliki", "web_manual_merge")],
        },

        # ============================================================== PDF
        {
            "key": "MIETEK|Edycja PDF",
            "tooltip": "Popraw dowolny plik PDF: dopisz brakujący tekst, "
                       "zakryj błędny fragment i wpisz nowy, usuń / obróć / "
                       "przesuń strony — i zapisz jako nowy plik (oryginał "
                       "zostaje nietknięty).",
            "controls": [
                {"id": "pdf_edycja", "kind": "pdf_edycja",
                 "label": "Edytor PDF"},
            ],
            "buttons": [],
        },

        # ============================================================== TAKSATOR
        {
            "key": "TAKSATOR|Układanie Exceli do druku",
            "tooltip": "Optymalizuje pliki Excel: ukrywa zbędne arkusze, sortuje i dostosowuje czcionki.",
            "controls": [
                _path("xl_src", "excel_folder_entry", "Folder z plikami Excel:",
                      "Wskaż folder z plikami .xls / .xlsx"),
                _path("xl_dst", "excel_output_entry", "Folder docelowy:",
                      "Wskaż folder zapisu dla ułożonych plików Excel"),
                _check("xl_subfolders", "include_subfolders_var", "Przetwarzaj także podfoldery", True),
                _check("xl_global_font", "global_font_var",
                       "Zastosuj ten sam rozmiar do WSZYSTKICH arkuszy:", False),
                _text("xl_global_size", "global_font_entry", "Rozmiar globalny:", "10"),
                {"id": "xl_fonts", "kind": "fonts", "label": "Dostosowanie rozmiaru czcionek w arkuszach:",
                 "fonts": fonts},
                _check("xl_remove_owners", "remove_owners_var",
                       "Usuń Właścicieli (Nazwisko, Adres, Współwłaśc., Udział)", False),
                _check("xl_remove_ls", "remove_ls_var", "Usuń Użytek ew.", False),
            ],
            "buttons": [_button("run", "Uruchom układanie Exceli", "start_excel")],
        },
        {
            "key": "TAKSATOR|Wyłożenie Exceli",
            "tooltip": "Scala strony tytułowe, opisy i raporty w gotowe paczki PDF dla każdej wsi.",
            "controls": [
                _path("le_title", "layout_title_folder_entry", "Folder STR_TYT:",
                      "Folder ze stronami tytułowymi STRTYT*.docx"),
                _path("le_opisy", "layout_opisy_folder_entry", "Folder z opisami:",
                      "Folder z plikami opisów"),
                _path("le_raporty", "layout_raporty_folder_entry", "Folder z raportami:",
                      "Folder z raportami Excel"),
                _path("le_out", "layout_output_folder_entry", "Folder docelowy PDF:",
                      "Folder wyjściowy na gotowe PDF"),
            ],
            "buttons": [_button("run", "Twórz gotowe PDF", "start_layout_excel")],
        },
        {
            "key": "TAKSATOR|Kreator Stron tytułowych",
            "tooltip": "Strony tytułowe: jedna wieś z formularza (Word + PDF) "
                       "albo masowo dla wielu wsi (dane z zestawień Excel).",
            "controls": [
                _segment("kreator_tryb_taks", "Tryb tworzenia:",
                         ["Jedna wieś", "Wiele wsi (masowo)"], "Jedna wieś",
                         groups={"Jedna wieś": "jedna",
                                 "Wiele wsi (masowo)": "wiele"}),
                _grp(_segment("tpl_TAKSATOR_format", "Format wyjścia:",
                              ["PDF", "Word"], "PDF",
                              attr="tpl_data.TAKSATOR.format_var",
                              sel_class={"PDF": "sel-pdf", "Word": "sel-word"}),
                     "jedna"),
                *[_grp(c, "jedna") for c in _tpl_controls("TAKSATOR")],
                _info_grupa("Wiele wsi (masowo) — dane z zestawień Excel:", "wiele"),
                _grp(_path("tt_template", "title_template_entry", "Szablon STR_TYT:",
                           "Wskaż plik bazowy (np. wygenerowany w Kreatorze Szablonów)",
                           kind="file"), "wiele"),
                _grp(_path("tt_excel", "title_excel_entry", "Fold. z rejestrami Excel:",
                           "Wskaż folder z plikami .xls / .xlsx"), "wiele"),
                _grp(_path("tt_out", "title_output_entry", "Folder zapisu STR_TYT:",
                           "Wskaż folder docelowy dla nowych stron"), "wiele"),
                _grp(_text("tt_village_ph", "title_village_placeholder_entry",
                           "Placeholder nazwy wsi:", "NAZWA WSI"), "wiele"),
                _grp(_text("tt_area_ph", "title_area_placeholder_entry",
                           "Placeholder powierzchni:", "wielkość"), "wiele"),
            ],
            "buttons": [
                _button("run", "Wygeneruj stronę tytułową", "generate_str_tyt:TAKSATOR",
                        group="jedna"),
                _button("run", "Wygeneruj strony tytułowe", "start_title_pages",
                        group="wiele"),
            ],
        },
        {
            "key": "TAKSATOR|Opisy ogólne",
            "tooltip": "Tworzy opisy ogólne (opis og_<wieś>.docx) na podstawie "
                       "raportów Excel do druku (zadania gospodarcze z arkusza Zestawienie).",
            "controls": [
                _path("ogt_root", "opis_og_taksator_entry", "Folder z raportami Excel:",
                      "Wskaż folder z ułożonymi raportami (np. 042-0001-Bysław.xls)"),
                _path("ogt_gdos", "opis_og_taksator_gdos_entry",
                      "Folder z wynikami GDOŚ (opcjonalnie):",
                      "Formy ochrony przyrody zostaną wstawione automatycznie"),
            ],
            "buttons": [_button("run", "Generuj opisy ogólne", "start_opis_og_taksator")],
        },
        {
            "key": "TAKSATOR|Excel -> PDF",
            "tooltip": "Konwertuje raporty i opisy jako osobne PDF podzielone na foldery wsi.",
            "controls": [
                _path("sp_title", "split_title_folder_entry", "Folder STR_TYT:",
                      "Folder ze stronami tytułowymi STRTYT*.docx"),
                _path("sp_opisy", "split_opisy_folder_entry", "Folder z opisami:",
                      "Folder z plikami opisów"),
                _path("sp_raporty", "split_raporty_folder_entry", "Folder z raportami:",
                      "Folder z raportami Excel"),
                _path("sp_out", "split_output_folder_entry", "Folder docelowy:",
                      "Folder wyjściowy dla rozdzielonych PDF"),
                _check("split_scalony", "split_scalony_var",
                       "Utwórz dodatkowo scalony PDF dla każdej wsi (w kolejności powstawania plików)",
                       False),
            ],
            "buttons": [_button("run", "Rozdziel na osobne PDF", "start_split_pdf")],
        },
        {
            "key": "TAKSATOR|Usuwanie 0 w MDB",
            "tooltip": "Kopiuje bazy Access (.mdb) do nowego folderu i modyfikuje adresy leśne.",
            "controls": [
                _path("mdb_src", "mdb_source_entry", "Folder źródłowy z .mdb:",
                      "Wskaż folder z oryginalnymi bazami"),
                _path("mdb_dst", "mdb_output_entry", "Folder docelowy zapisu:",
                      "Gdzie zapisać poprawione bazy?"),
            ],
            "buttons": [_button("run", "Usuń 0 w bazach (MDB)", "start_mdb_update")],
        },
        # ============================================================ ROZLICZANIE
        {
            "key": "ROZLICZANIE|Rozliczanie powierzchni",
            "tooltip": "Rozlicza powierzchnie ewidencji względem geodezji (.val).",
            "controls": [
                _path("rozl_xls", "rozl_xls_entry", "Folder z plikami XLS:",
                      "Wskaż folder z ewidencją (.xls/.xlsx)"),
                _path("rozl_val", "rozl_val_entry", "Folder z plikami VAL:",
                      "Wskaż folder z plikami z geodezji (.val)"),
                _path("rozl_out", "rozl_out_entry", "Folder docelowy:",
                      "Gdzie zapisać rozliczone tabele?"),
                _info("Dopasowanie: nazwa pliku XLS → plik *.val o tej samej nazwie "
                      "(ignorując spacje i _ )."),
                _check("rozl_wyrownywanie", "rozl_tylko_wyrownywanie_var",
                       "Wymuś proporcjonalne wyrównanie do ewidencji Ls (nie twórz arkuszy PRZYBYŁO/UBYŁO)", False),
                _check("rozl_puste_jrej", "rozl_usun_puste_jrej_var",
                       "Usuń wiersze jeśli brakuje wartości w J. rej. (usuwanie wydzieleń bez właścicieli)", False),
            ],
            "buttons": [_button("run", "Uruchom rozliczanie obrębów", "start_rozliczanie")],
        },
        {
            "key": "ROZLICZANIE|Tworzenie i wpisywanie mietków",
            "tooltip": "Generuje struktury MS-DOS (mietki) z bazą DBF z ewidencji.",
            "controls": [
                _path("tm_base", "mietki_bazowy_entry", "1. Folder Główny XLS (Ewidencja):",
                      "Stąd program pobierze nazwy wsi i właścicieli..."),
                _path("tm_rozlicz", "mietki_rozlicz_entry", "2. Folder XLSX (Rozliczone):",
                      "Stąd program pobierze numery J.rej i krzyżówki..."),
                _path("tm_out", "mietki_out_entry", "3. Folder docelowy zapisu:",
                      "Gdzie zapisać gotowe struktury MS-DOS z bazą DBF?"),
                _text("tm_wojew", "wsie_wojew_entry", "Województwo (kod):", "10", "np. 10"),
                _text("tm_powiat", "wsie_powiat_entry", "Powiat:", "", "np. WYSZKOWSKI"),
                _text("tm_stan", "wsie_stan_entry", "Stan na:", "01.01.2023", "DD.MM.RRRR"),
                _text("tm_obod", "wsie_obod_entry", "Obowiązuje od:", "01.01.2023", "DD.MM.RRRR"),
                _text("tm_obdo", "wsie_obdo_entry", "Obowiązuje do:", "31.12.2032", "DD.MM.RRRR"),
                _text("tm_nrws", "wsie_nrws_entry", "Nr wsi:", "1", "np. 1"),
                _text("tm_rokz", "wsie_rokz_entry", "Rok zal.:", "19", "np. 19"),
                _info("Dane nagłówka WSIE.DBF (stałe dla całego uruchomienia). "
                      "NAZWA i GMINA = nazwa obrębu, wpisywane automatycznie."),
                _check("tm_usun_puste", "krzyz_usun_puste_jrej_var",
                       "Usuń wydzielenia bez właścicieli (pomiń wiersze bez J. rej. przy wpisywaniu krzyżówek)", False),
            ],
            "buttons": [
                _button("run", "Generuj same mietki", "start_tworzenie_mietkow"),
                _button("krzyz", "Generuj mietki i wpisz krzyżówki", "start_mietki_krzyzowki", "secondary"),
            ],
        },
        {
            "key": "ROZLICZANIE|Halizny",
            "tooltip": "Przenosi halizny (wydzielenia niezalesione) w plikach D*.DBF.",
            "controls": [
                _path("halizny_src", "halizny_mietki_entry", "Gdzie leżą foldery obrębów:",
                      "Gdzie leżą foldery obrębów (np. BIAŁCZ\\WOL.001\\HALIZNY.TXT)?"),
            ],
            "buttons": [_button("run", "Przenieś halizny w D*.DBF", "start_halizny")],
        },
        {
            "key": "ROZLICZANIE|Zestawienie zbiorcze",
            "tooltip": "Składa wszystkie pliki <WIEŚ>_Rozliczone.xlsx w jeden arkusz: "
                       "sumy per wieś + wiersz RAZEM + rozpiska działek przybyło/ubyło.",
            "controls": [
                _path("zestaw", "zestaw_entry", "Folder z plikami rozliczeń (krzyżówki):",
                      "Folder docelowy rozliczeń (pliki <WIEŚ>_Rozliczone.xlsx)"),
                _path("zestaw_mietki", "zestaw_mietki_entry", "Folder z mietkami:",
                      "Foldery obrębów z wpisanymi krzyżówkami (D*.DBF), np. BIAŁCZ\\WOL.001"),
            ],
            "buttons": [
                _button("run", "Zestawienie z rozliczonych Exceli", "start_zestawienie"),
                _button("mietki", "Zestawienie z mietków (sumy z DBF)",
                        "start_zestawienie_mietki", "secondary"),
            ],
        },
        {
            "key": "ROZLICZANIE|Excel z MDB",
            "tooltip": "Wyciąga dane z bazy Access (.mdb) do pliku Excel.",
            "controls": [
                _path("zm_src", "excel_z_mdb_src_entry", "Plik źródłowy (.mdb):",
                      "Wskaż plik bazy danych MDB...", kind="file"),
                _path("zm_out", "excel_z_mdb_out_entry", "Zapisz Excel jako:",
                      "Gdzie zapisać gotowy plik .xlsx?", kind="save"),
            ],
            "buttons": [_button("run", "Wyciągnij dane z MDB", "start_excel_z_mdb")],
        },
        # ========================================================= KONWERTER PDF
        {
            "key": "KONWERTER PDF|Konwerter PDF",
            "tooltip": "Konwertuje dokumenty Office / obrazy do PDF.",
            "controls": [
                _path("pc_src", "pdfconv_source_entry",
                      "Folder źródłowy (opcjonalnie, gdy przeciągasz pliki):",
                      "Folder z plikami Office / PDF / obrazami..."),
                _path("pc_dst", "pdfconv_output_entry", "Folder docelowy PDF:",
                      "Miejsce zapisu przekonwertowanych dokumentów..."),
                {"id": "pc_drop", "kind": "dropfiles",
                 "label": "Przeciągnij pliki do konwersji (opcjonalnie):",
                 "ph": "Office / obrazy / PDF — można wiele naraz",
                 "exts": [".doc", ".docx", ".rtf", ".txt", ".xls", ".xlsx",
                          ".csv", ".jpg", ".jpeg", ".png", ".bmp", ".tif",
                          ".tiff", ".gif", ".webp", ".pdf"]},
            ],
            "buttons": [_button("run", "Konwertuj wszystko do PDF", "start_pdf_converter")],
        },
    ]

    # v2.0.140: do każdej zakładki dokładamy krótką, prostą instrukcję —
    # frontend pokazuje ją pod przyciskiem „Instrukcja” w nagłówku zakładki.
    for _t in tabs:
        _t["instrukcja"] = INSTRUKCJE.get(_t["key"])

    return {
        "app_name": "Forestly",
        "version": CURRENT_VERSION,
        "pdf_order_templates": [
            {"key": t["key"], "label": t["label"], "aliases": t["aliases"]}
            for t in PDF_ORDER_TEMPLATES],
        # pełna mapa terytorium: {WOJEWÓDZTWO: {POWIAT: [gminy]}} — frontend
        # odświeża listy powiatów/gmin po zmianie województwa/powiatu
        "tpl_territory": dict(TERRITORY_DATA) if TERRITORY_DATA else {},
        # ostatnio dodane / usprawnione — kafelki na ekranie Start
        "nowosci": NOWOSCI,
        "usprawnione": USPRAWNIONE,
        "tabs": tabs,
    }


def _tpl_controls(mode):
    """Kontrolki Kreatora Szablonu STR_TYT (dla MIETEK i TAKSATOR)."""
    from app.config import TERRITORY_DATA
    woj_list = sorted(TERRITORY_DATA.keys()) if TERRITORY_DATA else ["BRAK DANYCH"]
    default_woj = "KUJAWSKO-POMORSKIE" if "KUJAWSKO-POMORSKIE" in TERRITORY_DATA else (woj_list[0] if woj_list else "")
    powiat_list = sorted(TERRITORY_DATA.get(default_woj, {}).keys()) or ["BRAK DANYCH"]
    default_powiat = "TUCHOLSKI" if "TUCHOLSKI" in TERRITORY_DATA.get(default_woj, {}) else (powiat_list[0] if powiat_list else "")
    gmina_list = list(TERRITORY_DATA.get(default_woj, {}).get(default_powiat, []))
    b = f"tpl_data.{mode}."
    return [
        _select(f"tpl_{mode}_doc", b + "doc_type_var", "Typ dokumentu:", ["UPUL", "ISL"], "UPUL"),
        _select(f"tpl_{mode}_prefix", b + "prefix_var", "Prefiks obrębu:",
                ["położonych na terenie obrębu", "Obręb:"], "położonych na terenie obrębu"),
        _select(f"tpl_{mode}_woj", b + "woj_var", "Województwo:", woj_list, default_woj,
                free=True),
        _select(f"tpl_{mode}_powiat", b + "powiat_var", "Powiat:", powiat_list, default_powiat, free=True),
        _select(f"tpl_{mode}_gmina", b + "gmina_var", "Gmina:", gmina_list,
                gmina_list[0] if gmina_list else "", free=True),
        _text(f"tpl_{mode}_stan", b + "stan_na_entry", "Stan na:", "30.06.2026 r."),
        _text(f"tpl_{mode}_okres", b + "okres_entry", "Na okres:", "01.01.2027 – 31.12.2036 r."),
        _check(f"tpl_{mode}_single", b + "single_village_var", "Stwórz stronę dla konkretnej wsi", False),
        _text(f"tpl_{mode}_village", b + "village_entry", "Nazwa wsi:", "NAZWA WSI"),
        _check(f"tpl_{mode}_area", b + "area_var", "Dodaj wiersz z powierzchnią (ha)", False),
        _text(f"tpl_{mode}_area_v", b + "area_entry", "Powierzchnia:", "wielkość"),
        _path(f"tpl_{mode}_out", b + "output_entry", "Miejsce zapisu szablonu:",
              "Folder, w którym zapisać plik", kind="folder"),
    ]


# ---------------------------------------------------------------------------
# INSTRUKCJE ZAKŁADEK (v2.0.140)
# ---------------------------------------------------------------------------
# Proste, „ludzkie” wyjaśnienie dla każdej zakładki: co się w niej robi („co”)
# i jak się nią posługiwać („jak” — kroki). Frontend pokazuje to pod
# przyciskiem „Instrukcja” w nagłówku zakładki. Klucz = pełny klucz zakładki
# z „key” („SEKCJA|Nazwa zakładki”).
#
# Uzupełnienie treści NIE wymaga zmian w app.js — wystarczy dopisać/zmienić
# wpis w tym słowniku.

INSTRUKCJE = {
    # ============================================================ MIETEK
    "MIETEK|Mietek v2.0 — edytor danych": {
        "co": "Przeglądanie i ręczna poprawa danych mietka (pliki DBF) w wygodnej "
              "tabeli — bez wychodzenia z programu.",
        "jak": [
            "Wskaż folder z mietkiem.",
            "Wybierz obręb, a potem plik (W / R / O / D / Z / WSIE).",
            "Kliknij komórkę, popraw wartość i naciśnij „Zapisz”.",
            "Z tego samego miejsca możesz od razu wygenerować dokument (PDF) z danych.",
        ],
        "wskazowka": "Zapis robi kopię zapasową .BAK — pierwszą kopię program "
                     "zostawia nietkniętą, więc zawsze można wrócić do stanu sprzed zmian.",
    },
    "MIETEK|Pełny Automat (1-Click)": {
        "co": "Cała droga od mietków do gotowego, scalonego PDF — jednym przebiegiem.",
        "jak": [
            "W kreatorze wskaż folder z mietkami i folder, gdzie mają trafić wyniki.",
            "Ustaw stronę tytułową i daty (stan na / okres).",
            "Wybierz, czy dodać opis ogólny i formy ochrony przyrody (GDOŚ).",
            "Uruchom — program sam generuje wydruki z DBF, robi Word/PDF, dokłada "
            "skróty i mapy, scala i numeruje dokumenty.",
        ],
        "wskazowka": "Możesz włączyć „Obie wersje” — powstaną dwa gotowe zestawy: "
                     "z nazwiskami i bez nazwisk.",
    },
    "MIETEK|Nowe Szablony": {
        "co": "Nowe szablony wydruków (HTML → PDF, bez Worda). Jedna zakładka, "
              "dwa źródła danych — z mietka albo ze starych plików Word.",
        "jak": [
            "Wybierz źródło przełącznikiem „Źródło danych”.",
            "Pliki TXT / DBF mietka: wskaż folder z obrębami i folder docelowy, "
            "wybierz typ raportu i kliknij „Generuj PDF”.",
            "Stare pliki Word: wskaż folder ze starymi plikami (i opcjonalnie folder "
            "docelowy oraz wersję strony tytułowej), potem „Przerób na nowy szablon”.",
        ],
        "wskazowka": "W trybie „Stare pliki Word” program przerabia wszystkie raporty "
                     "oraz opis ogólny na nowy wygląd — oryginały zostają nietknięte.",
    },
    "MIETEK|Generowanie: MIETEK -> TXT": {
        "co": "Zamiana danych z plików DBF mietka na pliki wydrukowe (TXT) — "
              "te same, które dawniej tworzył program MS-DOS.",
        "jak": [
            "Wskaż folder z mietkami.",
            "Zaznacz, które wydruki mają powstać (albo „Wszystkie”).",
            "Uruchom — pliki zapisują się obok plików DBF.",
        ],
    },
    "MIETEK|Konwersja: MIETEK -> Word": {
        "co": "Etap 1: oczyszczenie surowych plików TXT z systemu MIETEK "
              "i ułożenie ich w pliki Word.",
        "jak": [
            "Wskaż folder z plikami TXT i folder docelowy.",
            "Zaznacz typy dokumentów do przetworzenia.",
            "Uruchom.",
        ],
    },
    "MIETEK|Konwersja: Word -> PDF": {
        "co": "Etap 2: zamiana gotowych plików Word na PDF i scalenie ich "
              "w jeden dokument.",
        "jak": [
            "Wskaż folder z plikami Word i folder docelowy.",
            "Uruchom — program przekonwertuje dokumenty i połączy je w jeden PDF.",
        ],
    },
    "MIETEK|Kreator Stron tytułowych": {
        "co": "Tworzenie strony tytułowej planu (Word i PDF) — dla jednej wsi "
              "albo masowo dla wielu.",
        "jak": [
            "Uzupełnij dane: typ dokumentu, obręb, województwo/powiat/gmina, daty.",
            "Wskaż szablon i folder zapisu.",
            "Kliknij, aby wygenerować. Tryb „Wiele wsi” tworzy strony dla "
            "wszystkich wsi z danych.",
        ],
    },
    "MIETEK|Opisy ogólne": {
        "co": "Automatyczne tworzenie pliku „opis og_<wieś>.docx” — opis ogólny "
              "uproszczonego planu.",
        "jak": [
            "Wskaż folder z wsiami (z plikami WSK_ZB).",
            "Opcjonalnie wskaż folder z wynikami GDOŚ (formy ochrony przyrody).",
            "Uruchom.",
        ],
    },
    "MIETEK|Wykaz Rozbieżności": {
        "co": "Porównanie powierzchni z mietków (DBF) z ewidencją w Excelu — "
              "wykaz różnic.",
        "jak": [
            "Wskaż folder z mietkami.",
            "Wskaż plik (albo folder) z ewidencją XLS.",
            "Uruchom.",
        ],
    },
    "MIETEK|Kontrola powierzchni Rejestr–OPTAX": {
        "co": "Raport rozbieżności powierzchni między Rejestrem a opisem "
              "taksacyjnym (OPTAX) — bez zmieniania danych mietka.",
        "jak": [
            "Wskaż folder mietków i folder wynikowy.",
            "Uruchom — powstaje osobny PDF dla każdej wsi oraz raport zbiorczy.",
        ],
    },
    "MIETEK|Czyszczenie rejestru": {
        "co": "Usuwa z mietka pozycje nierozliczone: działki bez przypisanej "
              "litery (pododdziału) oraz właścicieli bez rozliczonej działki. "
              "Potrafi też odświeżyć wygenerowane raporty.",
        "jak": [
            "Wskaż folder z mietkiem — pliki DBF mogą leżeć w podfolderach, "
            "także głębiej (np. CHORZEWO\\WOL.001).",
            "Opcjonalnie wskaż folder z wygenerowanymi raportami (nowe szablony) "
            "i zaznacz „Odśwież też wygenerowane raporty”.",
            "Uruchom. Program usuwa działki bez litery oraz właścicieli/pozycje "
            "rejestrowe bez rozliczonej działki, a na koniec odświeża REJESTR1 "
            "i ZEST1 (HTML + PDF) z oczyszczonego mietka.",
        ],
        "wskazowka": "To zmienia dane mietka — przed zapisem powstaje kopia .BAK "
                     "obok pliku DBF (pierwsza kopia zostaje nietknięta). Chcesz "
                     "najpierw tylko zobaczyć, co zostałoby usunięte? Zaznacz "
                     "„Tylko raport — NIE usuwaj danych”.",
    },
    "MIETEK|Opisy na mapę": {
        "co": "Wpisuje opisy z jednej z trzech baz wprost do poligonów mapy "
              "GEO-MAP: Forestly GO (arkusz), MIETKA (DBF) albo TAKSATORA (.mdb).",
        "jak": [
            "Wybierz bazę: „Baza Forestly GO”, „Baza MIETEK” albo „Baza TAKSATOR”.",
            "Wskaż odpowiednie źródło (folder z arkuszami / folder z mietkiem / plik .mdb).",
            "Wskaż folder z mapami — albo pojedynczy plik .MAP.",
            "Kliknij „Wpisz opisy do map” (albo najpierw „Sprawdź braki”).",
            "W tabeli braków popraw kolumnę „Co da reguła”, zaznacz "
            "„Zapamiętać?” i kliknij „Zapamiętaj zaznaczone” — program zapamięta "
            "zamianę i będzie ją sam stosował przy kolejnych uruchomieniach.",
            "Na koniec „Ułóż opisy automatycznie” — domyślnie wyśrodkuje i obróci "
            "opisy w wydzieleniach (wynik: <NAZWA>_ulozone.MAP). Odznacz "
            "„Tylko wyśrodkuj i obróć”, żeby program dodatkowo je rozsuwał.",
        ],
        "wskazowka": "Wynik zapisuje się obok mapy źródłowej jako "
                     "<NAZWA>_z_opisami.MAP, a nagłówek mapy jest przeliczany. "
                     "„Sprawdź braki przed wpisaniem” zapisuje różnice do "
                     "pliku „Opisy na mapę - braki.csv”. Zapamiętane zamiany "
                     "trzymane są w pliku config/zamiany_opisow.json.",
    },
    "MIETEK|Mietki +10 lat": {
        "co": "Przesunięcie wieku w opisach taksacyjnych o wybraną liczbę lat "
              "(zastępuje dawne makro VBA).",
        "jak": [
            "Wskaż folder mietka.",
            "Podaj liczbę lat i wybierz pola do zmiany.",
            "Najpierw podejrzyj wynik, a gdy jest poprawny — zastosuj zmiany.",
        ],
    },
    "MIETEK|NAZWISKA -> MIETEK": {
        "co": "Zbudowanie plików W*.DBF (właściciele) na podstawie ewidencji "
              "z Excela, ze strukturą MS-DOS.",
        "jak": [
            "Wskaż plik ewidencji XLS.",
            "Wskaż folder wzorcowy i folder wynikowy.",
            "Uruchom.",
        ],
    },
    "MIETEK|Baza obszarów GDOŚ": {
        "co": "Edytor bazy obszarów ochrony przyrody (Natura 2000, parki) "
              "używanej przy tworzeniu opisów ogólnych.",
        "jak": [
            "Przeglądaj i popraw wpisy w tabeli.",
            "Bazę możesz też zaimportować lub wyeksportować do Excela.",
        ],
        "wskazówka": "Nowe obszary dopisują się automatycznie z wyników GDOŚ — "
                     "kody i powiązania z gospodarką leśną możesz uzupełnić później.",
    },
    "MIETEK|Ręczne scalanie PDF": {
        "co": "Ręczne wczytanie luźnych plików PDF, ustawienie kolejności "
              "i połączenie ich w jeden dokument.",
        "jak": [
            "Wskaż folder z plikami PDF.",
            "Ustaw kolejność (przeciąganiem albo strzałkami).",
            "Kliknij, aby scalić.",
        ],
    },
    "MIETEK|Edycja PDF": {
        "co": "Poprawa dowolnego pliku PDF: dopisanie tekstu, zakrycie błędnego "
              "fragmentu, usuwanie / obrót / przesuwanie stron.",
        "jak": [
            "Otwórz plik PDF.",
            "Na podglądzie strony dopisz tekst albo załóż zakrycie (możesz też "
            "przestawić strony).",
            "Zapisz wynik jako nowy plik.",
        ],
        "wskazówka": "Oryginalny plik zostaje nietknięty — program zapisuje "
                     "zmiany pod nową nazwą.",
    },

    # ========================================================== TAKSATOR
    "TAKSATOR|Układanie Exceli do druku": {
        "co": "Przygotowanie arkuszy Excel do druku: ukrycie zbędnych arkuszy, "
              "sortowanie i dobranie czcionek.",
        "jak": [
            "Wskaż folder z plikami Excel i folder wynikowy.",
            "Uruchom.",
        ],
    },
    "TAKSATOR|Wyłożenie Exceli": {
        "co": "Złożenie stron tytułowych, opisów i raportów w gotowe paczki PDF "
              "dla każdej wsi.",
        "jak": [
            "Wskaż foldery ze stronami tytułowymi, opisami i raportami.",
            "Wskaż folder wynikowy.",
            "Uruchom.",
        ],
    },
    "TAKSATOR|Kreator Stron tytułowych": {
        "co": "Tworzenie stron tytułowych na podstawie danych z zestawień Excel "
              "(jedna wieś albo masowo).",
        "jak": [
            "Wskaż szablon oraz plik Excel z danymi.",
            "Wskaż folder zapisu i wybierz tryb.",
            "Kliknij, aby wygenerować.",
        ],
    },
    "TAKSATOR|Opisy ogólne": {
        "co": "Tworzenie opisów ogólnych (opis og_<wieś>.docx) na podstawie "
              "raportów Excel do druku.",
        "jak": [
            "Wskaż folder z raportami Excel (arkusz „Zestawienie”).",
            "Opcjonalnie wskaż folder z wynikami GDOŚ.",
            "Uruchom.",
        ],
    },
    "TAKSATOR|Excel -> PDF": {
        "co": "Konwersja raportów i opisów na osobne pliki PDF, rozdzielone do "
              "folderów poszczególnych wsi.",
        "jak": [
            "Wskaż folder źródłowy i folder docelowy.",
            "Uruchom.",
        ],
    },
    "TAKSATOR|Usuwanie 0 w MDB": {
        "co": "Kopia bazy Access (.mdb) do nowego folderu i poprawa adresów "
              "leśnych (usuwanie zbędnych zer).",
        "jak": [
            "Wskaż plik .mdb i folder docelowy.",
            "Uruchom.",
        ],
    },

    # ======================================================= ROZLICZANIE
    "ROZLICZANIE|Rozliczanie powierzchni": {
        "co": "Rozliczenie powierzchni z ewidencji względem geodezji "
              "(plik .val).",
        "jak": [
            "Wskaż plik ewidencji XLS.",
            "Wskaż plik .val (geodezja).",
            "Wskaż folder wynikowy i uruchom.",
        ],
    },
    "ROZLICZANIE|Tworzenie i wpisywanie mietków": {
        "co": "Wygenerowanie struktury MS-DOS (mietków) z bazą DBF z ewidencji "
              "oraz wpisanie krzyżówek.",
        "jak": [
            "Wskaż bazę z ewidencji i folder wynikowy.",
            "Podaj listę nazw (wsi).",
            "Uruchom.",
        ],
    },
    "ROZLICZANIE|Halizny": {
        "co": "Przeniesienie halizn (wydzieleń niezalesionych) w plikach D*.DBF.",
        "jak": [
            "Wskaż folder z mietkami.",
            "Uruchom.",
        ],
    },
    "ROZLICZANIE|Zestawienie zbiorcze": {
        "co": "Zebranie wszystkich plików „<WIEŚ>_Rozliczone.xlsx” w jeden "
              "arkusz: sumy dla każdej wsi, wiersz RAZEM oraz rozpiska "
              "działek przybyło/ubyło.",
        "jak": [
            "Wskaż folder z plikami rozliczeń.",
            "Uruchom.",
        ],
    },
    "ROZLICZANIE|Excel z MDB": {
        "co": "Wyciągnięcie danych z bazy Access (.mdb) do pliku Excel.",
        "jak": [
            "Wskaż plik .mdb i folder docelowy.",
            "Uruchom.",
        ],
    },

    # ======================================================= KONWERTER PDF
    "KONWERTER PDF|Konwerter PDF": {
        "co": "Konwersja dokumentów Office i obrazów do PDF.",
        "jak": [
            "Wskaż folder źródłowy i folder docelowy (albo przeciągnij pliki "
            "na listę).",
            "Uruchom.",
        ],
    },
}


# ---------------------------------------------------------------------------
# OSTATNIO DODANE / USPRAWNIONE (ekran Start)
# ---------------------------------------------------------------------------
# Dwie listy pokazywane na dole ekranu Start:
#   * NOWOSCI       — nowe zakładki/funkcje (zielone kafelki),
#   * USPRAWNIONE   — usprawnienia w zakładkach, które JUŻ były (niebieskie).
# Dopisuj nowy wpis NA POCZĄTKU po każdej aktualizacji (najnowszy pierwszy).
# Na Starcie widać 3 pierwsze wpisy, których zakładka faktycznie istnieje.
# "wersja" to tylko etykietka (plakietka) — możesz ją dowolnie ustawić.
NOWOSCI = [
    {"key": "MIETEK|Opisy na mapę", "wersja": "v2.0.143",
     "label": "Opisy na mapę — trzy bazy",
     "opis": "Jedna zakładka dla baz Forestly GO, MIETKA i TAKSATORA. "
             "Tabela braków ma edytowalną kolumnę reguły i znacznik "
             "„Zapamiętać?” — zapamiętane zamiany stosują się same."},
    {"key": "MIETEK|Czyszczenie rejestru", "wersja": "v2.0.140",
     "label": "Czyszczenie rejestru",
     "opis": "Usuwa działki bez litery i właścicieli bez rozliczonej działki."},
    {"key": "MIETEK|Kontrola powierzchni Rejestr–OPTAX", "wersja": "v2.0.139",
     "label": "Kontrola powierzchni Rejestr–OPTAX",
     "opis": "Raport rozbieżności powierzchni: Rejestr vs opis taksacyjny."},
    {"key": "MIETEK|Mietek v2.0 — edytor danych", "wersja": "v2.0.115",
     "label": "Mietek v2.0 — edytor danych",
     "opis": "Przeglądanie i edycja danych mietka (DBF) w tabeli."},
]

# Usprawnienia w zakładkach, które już istniały (niebieskie kafelki).
USPRAWNIONE = [
    {"key": "MIETEK|Nowe Szablony", "wersja": "v2.0.140",
     "label": "Nowe Szablony — jedno okno",
     "opis": "Scalone z „Stare → nowe szablony” + 3 miniatury strony tytułowej."},
    {"key": "MIETEK|Opisy ogólne", "wersja": "v2.0.140",
     "label": "Opisy ogólne — formy z GDOŚ",
     "opis": "Czyta formy ochrony także z nazw arkuszy GDOŚ (Rezerwaty, Parki Narodowe…)."},
    {"key": "MIETEK|Pełny Automat (1-Click)", "wersja": "v2.0.139",
     "label": "Pełny Automat — PDF w .exe",
     "opis": "Stabilne generowanie PDF (HTML → PDF) także z Forestly.exe."},
]
