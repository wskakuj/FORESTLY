# Forestly

System automatyzacji dokumentacji leśnej *(dawniej: Kombajn Leśny PRO)*.

Program ma **dwa interfejsy ze wspólną logiką**:

- **Nowe GUI (web)** — `python main_web.py` — nowoczesny interfejs HTML/CSS/JS
  w oknie aplikacji (PyWebView + WebView2, preinstalowany na Windows 10/11).
  Ekran Start z kartami kategorii, strony przeglądowe kategorii, zwijane grupy
  w menu, motyw ciemny/jasny, dziennik zdarzeń, dashboard kroków 1-Click.
- **Klasyczne GUI (CustomTkinter)** — `python main.py` — pełny fallback
  z całą funkcjonalnością.

Obie wersje **współdzielą ustawienia** (`settings.json`, `margins_config.json`,
`folder_history.json`) — można ich używać zamiennie.

> **Ważne:** `main.py` **i** `main_web.py` obsługują flagę `--word-worker`
> (proces pomocniczy Word COM, uruchamiany wewnętrznie przez oba GUI — także
> z wnętrza EXE, gdzie workerem jest sam plik `Forestly.exe`).
> Nie usuwać tego trybu — odpowiada za etap Word w Pełnym Automacie.

## Struktura projektu

```
kombajn_lesny/
├── main.py                      # Klasyczne GUI (CTk) + tryb --word-worker
├── main_web.py                  # Nowe GUI (PyWebView) — punkt wejścia
├── webapp/                      # Interfejs nowego GUI
│   ├── index.html               # Szkielet strony
│   ├── style.css                # Wygląd (motywy ciemny/jasny)
│   └── app.js                   # Logika frontendu (most pywebview)
├── app/
│   ├── config.py                # Stałe, wersja, sekwencje PCL, PDF_ORDER_TEMPLATES
│   ├── models.py                # Dataclasses (OrderStore, TerritoryEntry, etc.)
│   ├── updater.py               # Aktualizacja z GitHub (UpdaterMixin)
│   ├── gui/
│   │   ├── main_window.py       # ModernApp — shell klasycznego GUI + mixins
│   │   ├── web_schema.py        # Opis zakładek/kontrolek nowego GUI (1 źródło prawdy)
│   │   ├── web_backend.py      # WebBackend — te same mixiny, most pywebview
│   │   ├── tabs/                # Każda zakładka jako osobny mixin (wspólna logika)
│   │   │   ├── tab_all.py            # Pełny Automat (1-Click)
│   │   │   ├── tab_wydruki.py        # Generowanie: MIETEK -> TXT
│   │   │   ├── tab_word.py           # Konwersja: MIETEK -> Word
│   │   │   ├── tab_pdf.py            # Konwersja: Word -> PDF + scalanie
│   │   │   ├── tab_manual_merge.py   # Ręczne scalanie PDF
│   │   │   ├── tab_template_generator.py  # Kreator STR_TYT
│   │   │   ├── tab_title_pages.py    # Zaczytywanie danych STR_TYT
│   │   │   ├── tab_excel.py          # Układanie Exceli
│   │   │   ├── tab_layout_excel.py   # Wyłożenie Excel
│   │   │   ├── tab_split_pdf.py      # PDF + segregowanie wsi
│   │   │   ├── tab_mdb_update.py     # Usuwanie 0 w MDB
│   │   │   ├── tab_pdf_converter.py  # Konwerter PDF
│   │   │   ├── tab_rozliczanie.py    # Rozliczanie powierzchni
│   │   │   ├── tab_krzyzowki.py     # Wpisanie krzyżówek
│   │   │   ├── tab_halizny.py        # Halizny
│   │   │   ├── tab_excel_z_mdb.py   # Excel z MDB
│   │   │   ├── tab_tworzenie_mietkow.py  # Tworzenie Mietków
│   │   │   ├── tab_nazwiska_mietek.py     # NAZWISKA → MIETEK
│   │   │   └── tab_mietek_rozbieznosci.py # Wykaz Rozbieżności
│   │   └── widgets/             # Okna modalne i pomocnicze (CTk)
│   │       ├── pdf_order_window.py       # Kolejność PDF
│   │       ├── manual_pdf_merge_window.py # Ręczne scalanie
│   │       ├── changelog_window.py       # Okno changelogu
│   │       └── validation_window.py      # Okno walidacji
│   └── core/                    # Logika biznesowa (bez UI)
│       ├── word_worker.py       # Proces Word COM
│       ├── wydruki.py           # Wydruki MIETEK z DBF (8 typów plików)
│       └── excel_tasks.py       # Zadania Excel
├── config/
│   └── territory.json           # Dane województw/powiatów/gmin
├── tests/
│   ├── test_config.py           # Testy funkcji pomocniczych
│   └── test_mietek.py           # Testy parsowania DBF (placeholder)
├── requirements.txt
└── README.md
```

Pliki `settings.json`, `margins_config.json`, `folder_history.json` powstają
w czasie działania (zapamiętane ścieżki, marginesy, historia folderów).

## Uruchomienie

```bash
pip install -r requirements.txt
python main_web.py    # nowe GUI (zalecane)
python main.py        # klasyczne GUI (CustomTkinter)
```

## Nowe GUI (web) — co warto wiedzieć

- `main_web.py` uruchamia okno PyWebView z `webapp/`; cała logika Pythona
  siedzi w `app/gui/web_backend.py` (WebBackend dziedziczy **te same mixiny
  zakładek** co ModernApp — funkcjonalność 1:1 z wersją CTk).
- `app/gui/web_schema.py` opisuje zakładki i kontrolki (jedno źródło prawdy
  dla obu GUI web).
- Komunikaty Python → JS: kolejka zdarzeń odpytywana przez `poll()`
  (dziennik zdarzeń, pasek postępu, dashboard kroków 1-Click).
- Ekran **Start** otwiera się na starcie: baner „Pełny Automat (1-Click)"
  oraz karty kategorii (MIETEK, TAKSATOR, ROZLICZANIE, KONWERTER PDF);
  kliknięcie karty otwiera stronę kategorii ze wszystkimi jej zakładkami.
- Menu boczne: zwijane grupy sekcji (stan zapamiętywany), liczniki zakładek.
- Motyw ciemny/jasny — przełącznik ☀ w prawym górnym rogu.

## Wydruki MIETEKA bez MS-DOS (zakładka „Generowanie: MIETEK -> TXT")

Generuje pliki wydrukowe MIETEKA wprost z plików DBF mietka (format 1:1 z
oryginałem: cp852, CRLF, ramki, sekwencje PCL, paginacja):

  * **HALIZNY.TXT** — zestawienie powierzchni niezalesionych (uruchom
    **przed** przeniesieniem halizn w zakładce „Halizny");
  * **OPTAX.TXT** — opis lasów i gruntów przeznaczonych do zalesienia;
  * **TAB_KLW3.TXT** — zestawienie wg klas i podklas wieku + siedliska,
    ochronność, przebudowa (numeracja stron kontynuuje OPTAX);
  * **ZEST1.TXT** — skorowidz działek;
  * **REJESTR1.TXT** — rejestr działek wg właścicieli;
  * **WSKAZ1.TXT** — wskazówki gospodarki leśnej (wykaz wskaźników);
  * **WYK_NEG.TXT** — drzewostany negatywne i źle produkujące;
  * **WSK_ZB.TXT** — zestawienie czynności gospodarczych (wskaźniki zbiorcze).

Pozostałe wydruki generuj **po** przeniesieniu halizn. Wzorce odtworzone i
zweryfikowane bajt-po-bajcie na oryginalnych wydrukach MIETEKA (m.in. obręby
CHORZEWO i WOL.001); znane rozbieżności: ZEST1/REJESTR1 w niektórych obrębach
(JAŹWIE) to przybliżenia — kolejność wierszy odtwarzana heurystycznie, bo
sortowanie wg indeksów NTX nie jest w pełni odtwarzalne bez silnika MS-DOS.

## Układ scalanego PDF

- Domyślna kolejność szablonów: `PDF_ORDER_TEMPLATES` w `app/config.py`
  (Strona tytułowa, Opis ogólny, Tabela klas wieku, Opis taksacyjny,
  Wskazówki zbiorcze, Wykaz negatywny, Halizny, Wskazówki gospodarki leśnej,
  Rejestr, Skorowidz działek, Wykaz zmian, Skróty i symbole).
- „Skonfiguruj układ PDF" (Pełny Automat i Konwersja Word → PDF): zmiana
  kolejności strzałkami oraz **wykluczanie** pozycji (kosz/przywróć).
  Ustawienia zapisywane per folder docelowy (`<docelowy>\PDF`), wspólne
  dla obu GUI i respektowane przy scalaniu.
- Pliki niepasujące do szablonów trafiają na koniec scalonego dokumentu.

## Architektura

### Mixiny
Każda zakładka jest osobnym plikiem zawierającym klasę mixin (np. `TabAllMixin`).
Klasa `ModernApp` (klasyczne GUI) dziedziczy po wszystkich mixinach i `ctk.CTk`,
a `WebBackend` (nowe GUI) po tych samych mixinach — dzięki temu oba interfejsy
korzystają z identycznej logiki:

```python
class ModernApp(TabAllMixin, TabWordMixin, TabPdfMixin, ..., UpdaterMixin, ctk.CTk):
    ...

class WebBackend(TabAllMixin, TabWordMixin, TabPdfMixin, ...):
    ...
```

Dzięki temu:
- Każdy plik tab_*.py zawiera tylko metody dla jednej zakładki
- `main_window.py` zawiera tylko logikę wspólną (init, UI, dashboard, progress)
- Można testować logikę biznesową bez GUI

### Konfiguracja
Wszystkie stałe (m.in. `SEQUENCES_TO_REMOVE` — czyszczone sekwencje PCL,
`EXCEL_SHEET_DEFAULTS`, `PDF_ORDER_TEMPLATES`) są w `app/config.py`.

### Logika biznesowa
Funkcje niezależne od UI (Word COM, Excel, wydruki DBF) są w `app/core/`.
Funkcje zależne od UI (threads, pipelines) są w `app/gui/tabs/`.

## Jak pracować z AI nad tym kodem

1. **Pomoc z konkretną zakładką** — wyślij tylko odpowiedni `tab_*.py`
2. **Pomoc z logiką Word/Excel/wydrukami** — wyślij `core/word_worker.py`,
   `core/excel_tasks.py` lub `core/wydruki.py`
3. **Pomoc z UI web** — wyślij `webapp/app.js` + `webapp/style.css`
   (+ `web_schema.py`, jeśli zmiana dotyczy układu kontrolek)
4. **Pomoc z UI klasycznym** — wyślij `main_window.py` + odpowiedni `tab_*.py`
5. **Pomoc z konfiguracją** — wyślij `config.py`
6. **Pełny projekt** — spakuj cały katalog do zip

## GitHub — automatyczne budowanie EXE

### Struktura repozytorium
Wgraj CAŁY katalog `kombajn_lesny/` do repozytorium GitHub (zachowując
strukturę katalogów — w tym `webapp/`).

### Automatyczna budowa (GitHub Actions)
Plik `.github/workflows/build.yml` uruchamia się automatycznie, gdy
tworzysz tag `v*` (np. `v2.4.1`):

```bash
git add .
git commit -m "Forestly: nowe GUI web"
git tag v2.5.0
git push origin v2.4.1
```

GitHub Actions budują **dwa pliki EXE**:
1. **`Forestly.exe`** — nowe GUI web (`main_web.py`), z zasobami `webapp/`,
   `config/` i szablonami DOCX,
2. **`Forestly_OLD.exe`** — klasyczne GUI CustomTkinter (`main.py`),
   z zasobami `config/`, `pusty/` i szablonami DOCX.

Oba trafiają jako załączniki do Release. Aktualizator pobiera zawsze
`Forestly.exe` (nie pomylą go warianty).

### Ręczna budowa (lokalnie)
```bash
build.bat
```
Buduje `dist/Forestly.exe` i `dist/Forestly_OLD.exe` (oba GUI) — PyInstaller
pakuję­e do środka m.in. `webapp/`, `config/` i szablony DOCX.

### Aktualizacje
Program sprawdza GitHub Releases przy starcie (i po kliknięciu „↻ Aktualizacje").
Pobiera najnowszy `.exe` i podmienia go przez skrypt PowerShell.

Wersja jest ustawiona w `app/config.py`:
```python
CURRENT_VERSION = "v2.5.0"
GITHUB_USER = "wskakuj"
GITHUB_REPO = "FORESTLY"
```
Zmień ją przed każdym release'm.

### Ważne — plik `kombajn.ico`
Ikona `kombajn.ico` musi być w głównym katalogu projektu. Nie jest dołączona
do repo (dodaj ją ręcznie lub wstaw do `.gitignore` wyjątek).
