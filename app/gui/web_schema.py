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
  margins  — tabela marginesów (9 typów plików × 4 strony)
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


def _check(cid, attr, label, default=False, tooltip=""):
    return {"id": cid, "kind": "check", "attr": attr, "label": label,
            "default": default, "tooltip": tooltip, "save": True}


def _checks(cid, base, label, choices, tooltip=""):
    return {"id": cid, "kind": "checks", "attr_base": base, "label": label,
            "choices": choices, "tooltip": tooltip}


def _select(cid, attr, label, values, default, free=False):
    return {"id": cid, "kind": "select", "attr": attr, "label": label,
            "values": values, "default": default, "free": free}


def _info(text_):
    return {"kind": "info", "text": text_}


def _margins(mode):
    return {"id": f"margins_{mode}", "kind": "margins", "mode": mode,
            "label": "Ustawienia marginesów (w cm):"}


def _button(bid, label, task, style="primary", tooltip=""):
    return {"id": bid, "label": label, "task": task, "style": style,
            "tooltip": tooltip}


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
            "key": "MIETEK|Pełny Automat (1-Click)",
            "tooltip": "Kompleksowy proces: halizny, TXT z DBF, Word, PDF i scalanie w gotowy dokument.",
            "controls": [
                _path("all_src", "entries.ALL.src", "Folder źródłowy:", "Wskaż folder..."),
                _path("all_dst", "entries.ALL.dst", "Folder docelowy:", "Wskaż lokalizację..."),
                _check("all_gen_txt", "all_gen_txt_var",
                       "Generuj pliki TXT z DBF mietka (na starcie)", True,
                       "Pierwszy etap procesu, dla każdego obrębu źródłowego:\n"
                       "  1. generuje HALIZNY.TXT (zestawienie pow. niezalesionych),\n"
                       "  2. przenosi halizny w D*.DBF (jak przycisk w zakładce 'Halizny'),\n"
                       "  3. generuje OPTAX, TAB_KLW3, ZEST1, REJESTR1, WSKAZ1, WYK_NEG\n"
                       "     i WSK_ZB bezpośrednio z DBF.\n"
                       "Powtórne uruchomienie jest bezpieczne."),
                _check("all_str_tyt", "all_gen_str_tyt_var",
                       "Generuj strony tytułowe (STR_TYT) na podstawie OPTAX", False),
                _path("all_template", "all_template_entry", "Plik szablonu (.docx):",
                      "Wskaż plik bazowy STR_TYT...", kind="file"),
                _check("all_custom_skroty", "all_custom_skroty_var",
                       "Użyj własnego pliku 'Skróty i symbole' (zamiast domyślnego z programu)", False),
                _path("all_skroty", "all_skroty_entry", "Własny plik:",
                      "Wskaż własny plik ze skrótami...", kind="file"),
                _check("remove_names", "remove_names_var",
                       "Usuwaj nazwiska z REJESTRU (oraz 1. stronę)", True,
                       "Włączenie tej opcji uruchamia makra 'ZamienLF' oraz 'UsunNazwiskaRej', "
                       "a także kasuje pierwszą stronę z rejestru."),
                _margins("ALL"),
                {"kind": "dashboard", "id": "dashboard", "steps": DASHBOARD_STEPS},
            ],
            "buttons": [
                _button("run", "▶  Rozpocznij proces", "start_pipeline:ALL"),
                _button("order", "Skonfiguruj układ PDF", "open_order:ALL", "secondary"),
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
            "key": "MIETEK|Kreator Szablonu STR_TYT",
            "tooltip": "Generuje jeden bazowy dokument Word ze stroną tytułową na podstawie danych.",
            "controls": _tpl_controls("MIETEK"),
            "buttons": [_button("run", "Wygeneruj Szablon STR_TYT", "generate_template:MIETEK")],
        },
        {
            "key": "MIETEK|Zaczytywanie danych STR_TYT",
            "tooltip": "Masowo tworzy strony tytułowe dla każdej wsi (MIETEK), wciągając dane z OPTAX.",
            "controls": [
                _path("mt_template", "mietek_title_template_entry", "Szablon STR_TYT:",
                      "Wskaż plik bazowy", kind="file"),
                _path("mt_word", "mietek_title_word_entry", "Fold. z plikami Word (OPTAX):",
                      "Wskaż folder, w którym znajdują się pliki OPTAX"),
                _path("mt_out", "mietek_title_output_entry", "Folder zapisu STR_TYT:",
                      "Wskaż folder docelowy dla nowych stron"),
                _text("mt_village_ph", "mietek_title_village_placeholder_entry",
                      "Placeholder nazwy wsi:", "NAZWA WSI"),
                _text("mt_area_ph", "mietek_title_area_placeholder_entry",
                      "Placeholder powierzchni:", "wielkość"),
            ],
            "buttons": [_button("run", "Masowo twórz strony STR_TYT", "start_mietek_title_pages")],
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
                _path("opis_og_gdos", "opis_og_gdos_entry",
                      "Folder z wynikami GDOŚ (opcjonalnie):",
                      "np. folder z plikami *_wynik.xlsx — formy ochrony przyrody "
                      "zostaną wstawione automatycznie"),
            ],
            "buttons": [_button("run", "Generuj opisy ogólne", "start_opis_og")],
        },

        # ============================================================== TAKSATOR
        {
            "key": "TAKSATOR|Kreator Szablonu STR_TYT",
            "tooltip": "Generuje jeden bazowy dokument Word ze stroną tytułową na podstawie danych.",
            "controls": _tpl_controls("TAKSATOR"),
            "buttons": [_button("run", "Wygeneruj Szablon STR_TYT", "generate_template:TAKSATOR")],
        },
        {
            "key": "TAKSATOR|Zaczytywanie danych STR_TYT",
            "tooltip": "Masowo tworzy strony tytułowe dla każdej wsi, wciągając dane z zestawień Excel.",
            "controls": [
                _path("tt_template", "title_template_entry", "Szablon STR_TYT:",
                      "Wskaż plik bazowy (np. wygenerowany w Kreatorze Szablonów)", kind="file"),
                _path("tt_excel", "title_excel_entry", "Fold. z rejestrami Excel:",
                      "Wskaż folder z plikami .xls / .xlsx"),
                _path("tt_out", "title_output_entry", "Folder zapisu STR_TYT:",
                      "Wskaż folder docelowy dla nowych stron"),
                _text("tt_village_ph", "title_village_placeholder_entry",
                      "Placeholder nazwy wsi:", "NAZWA WSI"),
                _text("tt_area_ph", "title_area_placeholder_entry",
                      "Placeholder powierzchni:", "wielkość"),
            ],
            "buttons": [_button("run", "Masowo twórz strony STR_TYT", "start_title_pages")],
        },
        {
            "key": "TAKSATOR|Układanie Exceli",
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
            "key": "TAKSATOR|Wyłożenie Excel",
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
            "key": "TAKSATOR|PDF + segregowanie wsi",
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
                _path("pc_src", "pdfconv_source_entry", "Folder źródłowy:",
                      "Folder z plikami Office / PDF / obrazami..."),
                _path("pc_dst", "pdfconv_output_entry", "Folder docelowy PDF:",
                      "Miejsce zapisu przekonwertowanych dokumentów..."),
            ],
            "buttons": [_button("run", "Konwertuj wszystko do PDF", "start_pdf_converter")],
        },
    ]

    return {
        "app_name": "Forestly",
        "version": CURRENT_VERSION,
        "pdf_order_templates": [
            {"key": t["key"], "label": t["label"], "aliases": t["aliases"]}
            for t in PDF_ORDER_TEMPLATES],
        "tpl_territory": {
            "woj": woj_list, "powiat": powiat_list, "gmina": gmina_list,
        },
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
        _select(f"tpl_{mode}_woj", b + "woj_var", "Województwo:", woj_list, default_woj),
        _select(f"tpl_{mode}_powiat", b + "powiat_var", "Powiat:", powiat_list, default_powiat, free=True),
        _select(f"tpl_{mode}_gmina", b + "gmina_var", "Gmina:", gmina_list,
                gmina_list[0] if gmina_list else "", free=True),
        _text(f"tpl_{mode}_stan", b + "stan_na_entry", "Stan na:", "30.06.2026 r."),
        _text(f"tpl_{mode}_okres", b + "okres_entry", "Na okres:", "01.01.2027 – 31.12.2036 r."),
        _check(f"tpl_{mode}_single", b + "single_village_var", "Stwórz stronę dla konkretnej wsi", False),
        _text(f"tpl_{mode}_village", b + "village_entry", "Nazwa wsi:", "NAZWA WSI"),
        _check(f"tpl_{mode}_area", b + "area_var", "Dodaj wiersz z powierzchnią (ha)", False),
        _text(f"tpl_{mode}_area_v", b + "area_entry", "Powierzchnia:", "wielkość"),
        _path(f"tpl_{mode}_out", b + "output_entry", "Zapisz szablon jako:",
              "Ścieżka do pliku np. Szablon.docx", kind="save"),
    ]
