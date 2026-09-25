# -*- coding: utf-8 -*-
"""
Forestly — generator „opisu ogólnego" (opis og) dla wsi.
======================================================================
Zależności: python-docx, szablon opis_og_szablon.docx (repo / EXE)

Co robi:
  Dla każdego folderu wsi (folder z plikiem WSK_ZB.doc) tworzy plik
  „opis og_<nazwa wsi>.docx" — opis ogólny uproszczonego planu.
  Jeśli w wskazanym folderze są dane MIETEKA (O*.DBF) zamiast Worda,
  program robi w folderze TYMCZASOWYM MIETEK -> TXT -> Word (tylko
  WSK_ZB), na tej podstawie tworzy opis ogólny, a pliki tymczasowe
  usuwa — oryginalne dane MIETEKA pozostają nietknięte.

Skąd bierze dane:
  Liczby (etaty, użytkowanie przedrębne) czyta BEZ WORDA bezpośrednio
  z pliku WSK_ZB.doc w folderze wsi (Word 97 zapisuje tekst w UTF-16LE):
    Etat wg potrzeb hodowlanych  = Użytki rębne właściwe (m3)
    Etat przyjęty                = Użytki rębne właściwe (m3)
    Pozostałe użytki rębne       = Pozostałe użytki rębne (m3)
    Maksymalna miąższość         = Ogółem użytki rębne (m3)
    Użytkowanie przedrębne       = Razem użytki przedrębne (m3)
  Formy ochrony przyrody (Natura 2000 itp.) nie występują w żadnym
  pliku folderu — w wygenerowanym dokumencie zostaje wyraźnie
  oznaczone miejsce do ręcznego wpisania.
"""

import json
import re
import shutil
import sys
import tempfile
import threading
import traceback
from pathlib import Path

TPL_FILENAME = "opis_og_szablon.docx"            # wariant MIETEK (WSK_ZB) — tabela wielkopolska
TPL_MAZ_FILENAME = "opis_og_szablon_mazowiecka.docx"  # wariant MIETEK — tabela mazowiecka
TPL_TAKSATOR_FILENAME = "opis_og_szablon_taksator.docx"  # wariant TAKSATOR (raporty)

# domyślne teksty edytowalnych sekcji opisu ogólnego (kreator 1-Click);
# użytkownik może je nadpisać w polach '1. NADZÓR' i '2. WARUNKI PRZYRODNICZE'
OG_NADZOR_DOMYSLNY = (
    "Nadzór nad gospodarką leśną lasów nie stanowiących własności Skarbu "
    "Państwa sprawuje Starosta Wołomiński w zakresie zadań własnych.")
OG_WARUNKI_DOMYSLNE = (
    "Lasy objęte uproszczonym planem urządzenia lasów położone są w:\n"
    "IV Mazowiecko-Podlaskiej krainie przyrodniczo-leśnej\n"
    "Mezoregion Doliny Dolnego Bugu")
# kategoria zagrożenia pożarowego -> opis w zdaniu
OG_KATEGORIE = {
    "I": "dużego",
    "II": "średniego",
    "III": "małego",
}
OCHRONA_MARKER = "[TU WPISZ formy ochrony przyrody — po wpisaniu usuń tę linię]"


class TabOpisOgMixin:
    """Generator opisu ogólnego (zakładka TAKSATOR | Opis ogólny)."""

    # ------------------------------------------------ Parsing WSK_ZB (bez Worda)
    @staticmethod
    def _parse_wsk_zb_text(text):
        """Wyciąga 4 liczby (m3) z treści WSK_ZB. Zwraca dict lub None."""

        def m3(pattern):
            m = re.search(pattern + r"\s+([\d.,]+) ha\s+([\d.,]+) m3", text)
            if not m:
                return None
            return m.group(2).replace(",", ".")

        wl = m3(r"1\. Użytki rębne właściwe")
        if wl is None:
            return None
        po = m3(r"2\. Pozostałe użytki rębne") or "0"
        og = m3(r"Ogółem użytki rębne") or wl
        pr = m3(r"Razem użytki przedrębne") or "0"
        return {
            "ETAT_POTRZEBY": wl,
            "ETAT_PRZYJETY": wl,
            "POZOSTALE_REBNE": po,
            "MAKS_MIAZSZOSC": og,
            "UZYTK_PRZEDRZEBNE": pr,
        }

    @classmethod
    def _read_wsk_zb_doc(cls, path):
        """Czyta liczby z WSK_ZB.doc (binarnie) albo WSK_ZB.TXT (cp852) — bez Worda."""
        try:
            data = Path(path).read_bytes()
        except OSError as e:
            return None, f"nie można odczytać pliku ({e})"
        if str(path).lower().endswith(".txt"):
            encs = ("cp852", "cp1250", "utf-8")
        else:
            encs = ("utf-16-le", "cp1250", "utf-8")
        for enc in encs:
            vals = cls._parse_wsk_zb_text(data.decode(enc, errors="ignore"))
            if vals:
                return vals, None
        return None, "nie znaleziono danych o użytkowaniu lasu w pliku"

    # ------------------------------------------------ Wypełnianie szablonu
    @staticmethod
    def _podmien_akapit(p, tekst):
        """Ustawia tekst akapitu (zachowując formatowanie pierwszego runa)."""
        if p.runs:
            p.runs[0].text = tekst
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.add_run(tekst)

    @classmethod
    def _wpisz_linie(cls, p, linie):
        """Wpisuje wieloliniowy tekst: pierwsza linia w 'p', kolejne jako
        kopie akapitu wstawione za nim (to samo wcięcie/formatowanie)."""
        import copy as _copy
        from docx.text.paragraph import Paragraph
        if not linie:
            return p
        cls._podmien_akapit(p, linie[0])
        kotwica = p._p
        for ln in linie[1:]:
            np_ = _copy.deepcopy(p._p)
            # w kopii zostaje jeden run z tekstem
            kotwica.addnext(np_)
            nowy = Paragraph(np_, p._parent)
            cls._podmien_akapit(nowy, ln)
            kotwica = np_
        return Paragraph(kotwica, p._parent)

    @staticmethod
    def _podmien_sekcje(doc, nadzor=None, warunki=None, kategoria=None):
        """Podmienia edytowalne sekcje opisu ogólnego (kreator 1-Click).

        nadzor — tekst pod nagłówkiem '1. NADZÓR' (wieloliniowy);
        warunki — tekst pod '2. WARUNKI PRZYRODNICZE' (do nagłówka tabeli
                  siedliskowej, wieloliniowy);
        kategoria — 'I', 'II' albo 'III' (kategoria zagrożenia pożarowego).
        None = zostaje treść szablonu.
        """
        from docx.oxml.ns import qn as _qn
        paras = list(doc.paragraphs)

        def _czysty(el):
            return "".join(t.text or "" for t in el.findall(".//" + _qn("w:t"))).strip()

        if nadzor:
            for i, p in enumerate(paras):
                if p.text.strip().startswith("Nadzór nad gospodark"):
                    TabOpisOgMixin._wpisz_linie(p, [l for l in nadzor.split("\n")])
                    break

        if warunki:
            linie = [l for l in warunki.split("\n")]
            for i, p in enumerate(paras):
                if "położone są w" in p.text:
                    ostatni = TabOpisOgMixin._wpisz_linie(p, linie)
                    # stare akapity sekcji (kraina, mezoregion, puste) — do
                    # pierwszego niepustego 'Poniżej przedstawiono' / tabeli;
                    # zaczynamy ZA świeżo wstawionymi liniami
                    nast = ostatni._p.getnext()
                    while nast is not None:
                        if nast.tag != _qn("w:p"):
                            break
                        txt = _czysty(nast)
                        if txt.startswith("Poniżej przedstawiono"):
                            break
                        po = nast.getnext()
                        nast.getparent().remove(nast)
                        nast = po
                    break

        if kategoria and str(kategoria).upper() in OG_KATEGORIE:
            kat = str(kategoria).upper()
            for p in paras:
                if "należą do" in p.text and "kategorii" in p.text:
                    TabOpisOgMixin._podmien_akapit(
                        p, f"Lasy objęte opracowaniem, należą do {kat} kategorii - "
                           f"{OG_KATEGORIE[kat]} zagrożenia pożarowego.")
                    break

    @staticmethod
    def _fill_template(tpl_path, values, out_path, formy=None,
                       nadzor=None, warunki=None, kategoria=None):
        """Wypełnia opis_og_szablon.docx liczbami i zapisuje jako .docx.

        formy — lista akapitów o formach ochrony przyrody (z wyników GDOŚ);
        None => w tekście zostaje zdanie o braku form ochrony przyrody.
        nadzor/warunki/kategoria — edytowalne sekcje z kreatora 1-Click.
        """
        from docx import Document
        import copy

        doc = Document(str(tpl_path))
        TabOpisOgMixin._podmien_sekcje(doc, nadzor=nadzor, warunki=warunki,
                                       kategoria=kategoria)
        vals = dict(values)
        vals["FORMY_OCHRONY"] = (formy[0] if formy
                                 else "Nie zlokalizowano form ochrony przyrody.")

        # tabele (etaty, przedrębne): komórka zawierająca sam placeholder
        # dostaje wartość; klucza brak w danych (np. MIETEK bez arkusza Etaty)
        # => pusta komórka zamiast dosłownego {{...}}
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    ph = cell.text.strip()
                    if ph.startswith("{{") and ph.endswith("}}"):
                        if cell.paragraphs and cell.paragraphs[0].runs:
                            cell.paragraphs[0].runs[0].text = vals.get(ph[2:-2], "")

        # akapity + marker do ręcznego wpisania form ochrony przyrody
        marker_done = False
        for p in list(doc.paragraphs):
            txt = p.text
            m = re.search(r"\{\{(\w+)\}\}", txt)
            if m and m.group(1) not in vals:
                # brak danych (np. {{OBREB}} w wariancie MIETEK) — usuń akapit
                p._p.getparent().remove(p._p)
                continue
            if not m:
                continue
            new_txt = re.sub(r"\{\{(\w+)\}\}", lambda k: vals[k.group(1)], txt)
            if p.runs:
                p.runs[0].text = new_txt
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.add_run(new_txt)
            if m.group(1) == "FORMY_OCHRONY":
                from docx.text.paragraph import Paragraph
                PODKRESL = "Powiązanie z gospodarką leśną"

                def _ustaw_akapit(par, tekst):
                    """Formatuje akapit formy: zwykła czcionka; dla 'Powiązanie...'
                    fraza wiodąca podkreślona (reszta zwykła) — jak w opisach wzorcowych."""
                    for r in par.runs:
                        r.text = ""
                    r0 = par.runs[0] if par.runs else par.add_run("")
                    if tekst.startswith(PODKRESL):
                        if len(par.runs) == 1:
                            par._p.append(copy.deepcopy(r0._r))
                        r0.text = PODKRESL
                        r0.bold = None
                        r0.underline = True
                        r1 = par.runs[1]
                        r1.text = tekst[len(PODKRESL):]
                        r1.bold = None
                        r1.underline = None
                    else:
                        r0.text = tekst
                        r0.bold = None
                        r0.underline = None

                from docx.oxml.ns import qn as _qn

                def _usun_keepnext(par):
                    """Bez keepNext — akapity form nie mogą tworzyć niełamliwego
                    bloku, który przy braku miejsca skacze w całości na nową stronę
                    (zostawiając pustkę pod nagłówkiem sekcji)."""
                    par.paragraph_format.keep_with_next = None

                def _usun_puste_za(el):
                    """Usuwa puste akapity pozostawione w szablonie za sekcją
                    (miejsce na ręczne wpisanie form) — do najbliższego tekstu."""
                    nast = el.getnext()
                    while nast is not None and nast.tag == _qn("w:p"):
                        tekst = "".join(t.text or "" for t in nast.findall(".//" + _qn("w:t"))).strip()
                        if tekst:
                            break
                        po = nast.getnext()
                        nast.getparent().remove(nast)
                        nast = po

                if formy:
                    # nagłówek "Zlokalizowano..." zostaje pogrubiony (jak we wzorcu),
                    # pojedynczy akapit (np. brak form) — czcionka zwykła
                    if not formy[0].startswith("Zlokalizowano"):
                        for r in p.runs:
                            if r.bold:
                                r.bold = None
                        _usun_keepnext(p)
                        _usun_puste_za(p._p)
                    if len(formy) > 1:
                        # wstaw kolejne akapity form ochrony (z GDOŚ) pod placeholderem
                        anchor = p._p
                        for extra in formy[1:]:
                            new_p = copy.deepcopy(p._p)
                            anchor.addnext(new_p)
                            np_ = Paragraph(new_p, p._parent)
                            _ustaw_akapit(np_, extra)
                            _usun_keepnext(np_)
                            anchor = new_p
                        _usun_puste_za(anchor)
                elif not formy:
                    # brak form ochrony przyrody — zwykła czcionka (nie jak
                    # nagłówek sekcji), bez keepNext i pustych akapitów za sekcją
                    for r in p.runs:
                        if r.bold:
                            r.bold = None
                    _usun_keepnext(p)
                    _usun_puste_za(p._p)

        doc.save(str(out_path))

    # ------------------------------------------------ Właściwy przebieg zadania
    def start_opis_og_pipeline(self):
        """Zadanie z mapy zadań (web) / przycisk (stare GUI)."""
        self._disable_ui_for_process()

        def _run():
            try:
                self._opis_og_run()
            except Exception:
                self.log("[OPIS OG BŁĄD] " + traceback.format_exc())
                self.update_status("Błąd", "#D83B01", animate=False)
            finally:
                self.restore_all_buttons()

        threading.Thread(target=_run, daemon=True).start()

    def _opis_og_run(self):
        """Wariant z zakładki 'Opisy ogólne': folder i GDOŚ czyta z pól GUI."""
        entry = getattr(self, "opis_og_root_entry", None)
        raw = entry.get().strip() if entry else ""
        if not raw:
            self.log("[OPIS OG] Wskaż najpierw folder (główny z folderami wsi albo pojedynczą wieś).")
            self.update_status("Brak folderu", "#D83B01", animate=False)
            return
        gdos_entry = getattr(self, "opis_og_gdos_entry", None)
        gdos_raw = gdos_entry.get().strip() if gdos_entry is not None else ""
        # przełączniki pełne/skrócone (wzajemnie się wykluczają; oba odznaczone
        # = opisy pomijane) — tak samo jak w kreatorze Pełnego Automatu
        pelny_v = getattr(self, "opis_og_pelny_var", None)
        krotki_v = getattr(self, "opis_og_krotki_var", None)
        pelny = bool(pelny_v.get()) if pelny_v is not None else True
        krotki = bool(krotki_v.get()) if krotki_v is not None else False
        if pelny and krotki:            # obrona przed starym zapisem obu
            krotki = False
        if not pelny and not krotki:
            self.log("[OPIS OG] Oba przełączniki (pełne / skrócone) są odznaczone "
                     "— opisy ogólne nie powstaną.")
            self.update_status("Opisy ogólne pominięte", "#D83B01", animate=False)
            return
        self._opis_og_generuj(Path(raw), Path(gdos_raw) if gdos_raw else None,
                              skrocony=krotki)

    def _opis_og_generuj(self, root, gdos_folder=None, tylko_istniejace=False,
                         skrocony=False, nadzor=None, warunki=None,
                         kategoria=None, tabela=None):
        """Generuje 'opis og_<wieś>.docx' we wszystkich wsiach pod 'root'.

        root — folder główny z folderami wsi (albo pojedyncza wieś z WSK_ZB.doc);
        gdos_folder — opcjonalny folder z wynikami GDOŚ (formy ochrony przyrody);
        tylko_istniejace — bez tworzenia tymczasowych Wordów z DBF (Pełny
        Automat: korzystamy wyłącznie z WSK_ZB.doc utworzonych przez pipeline);
        skrocony — opis ogólny bez 'Powiązanie z gospodarką leśną' i bez opisów
        pozostałych form ochrony przyrody (sama lista form)."""
        from app.core.word_worker import get_resource_path

        root = Path(root)
        self.last_output_dir = root
        if not root.exists():
            self.log(f"[OPIS OG] Folder nie istnieje: {root}")
            self.update_status("Brak folderu", "#D83B01", animate=False)
            return

        # --- jakie mamy źródła? ---
        # 1) wsie z gotowym WSK_ZB (.doc z Worda albo .TXT z nowych szablonów)
        def _wsk_zb(d):
            for n in ("WSK_ZB.doc", "WSK_ZB.TXT", "WSK_ZB.txt"):
                if (d / n).exists():
                    return d / n
            return None

        if _wsk_zb(root):
            word_villages = [root]
        else:
            word_villages = sorted(
                d for d in root.iterdir()
                if d.is_dir() and _wsk_zb(d)
            )
        # 2) wsie z danymi MIETEKA (O*.DBF) — dla nich robimy tymczasowo
        #    MIETEK -> TXT -> Word i czytamy WSK_ZB z folderu tymczasowego
        #    (w Pełnym Automacie wyłączone — WSK_ZB.doc już istnieją z pipeline)
        if tylko_istniejace:
            mietek_sources = []
            mietek_todo = []
        else:
            mietek_sources = [(v, dbf) for v, dbf in self._find_mietek_sources(root)
                              if v not in word_villages]
            mietek_todo = [v for v, _dbf in mietek_sources]

        if not word_villages and not mietek_todo:
            self.log("[OPIS OG] Nie znaleziono folderów wsi z plikiem WSK_ZB.doc"
                     + ("" if tylko_istniejace else
                        " ani z danymi MIETEKA (O*.DBF)")
                     + f" w: {root}")
            self.update_status("Brak wsi", "#D83B01", animate=False)
            return

        # wariant tabeli siedliskowej: mazowiecka (domyślnie) albo wielkopolska
        if str(tabela or "").strip().lower().startswith("wielk"):
            tpl = get_resource_path(TPL_FILENAME)
        else:
            tpl = get_resource_path(TPL_MAZ_FILENAME)
            if not Path(tpl).exists():
                tpl = get_resource_path(TPL_FILENAME)
        if not Path(tpl).exists():
            self.log(f"[OPIS OG BŁĄD] Brak szablonu: {tpl}")
            self.update_status("Brak szablonu", "#D83B01", animate=False)
            return

        # opcjonalny folder z wynikami GDOŚ -> automatyczne formy ochrony przyrody
        gdos_map = {}
        if gdos_folder is not None:
            if Path(gdos_folder).exists():
                gdos_map = self._gdos_files_map(gdos_folder)
                if gdos_map:
                    self.log(f"[OPIS OG] Wyniki GDOŚ: {len(gdos_map)} plik(ów) — formy "
                             "ochrony przyrody zostaną wstawione automatycznie.")
            else:
                self.log(f"[OPIS OG] Folder GDOŚ nie istnieje: {gdos_folder} — pomijam.")

        temp_dir = None
        try:
            # --- lista zadań: (folder docelowy, nazwa wsi, skąd czytać WSK_ZB.doc) ---
            # dla wsi z Wordem: folder wsi; dla MIETEKA: folder z plikami DBF
            jobs = [(d, d.name, _wsk_zb(d) or (d / "WSK_ZB.doc"))
                    for d in word_villages]
            if mietek_todo:
                temp_dir, wsk_map = self._mietek_to_wsk_docs(root, mietek_sources)
                # folder z DBF -> nazwa wsi (z pary wyznaczonej przy detekcji)
                dbf_to_village = {dbf: v for v, dbf in mietek_sources}
                for dbf_dir, wsk_path in sorted(wsk_map.items()):
                    v = dbf_to_village.get(dbf_dir, dbf_dir)
                    jobs.append((dbf_dir, v.name, wsk_path))
                if not wsk_map:
                    self.log("[OPIS OG] Z danych MIETEKA nie udało się wygenerować "
                             "WSK_ZB.doc — te wsie zostaną pominięte.")

            if not jobs:
                self.log("[OPIS OG] Nic do wygenerowania.")
                self.update_status("Brak wsi", "#D83B01", animate=False)
                return

            self.log(f"[OPIS OG] Generator start — wsi: {len(jobs)}, "
                     f"szablon: {Path(tpl).name}"
                     + (f" (w tym z MIETEKA: {len(mietek_todo)})" if mietek_todo else ""))
            done = 0
            for i, (d, vname, wsk_path) in enumerate(jobs, 1):
                self.check_stop()
                self.update_status(f"Opis ogólny: {vname} ({i}/{len(jobs)})",
                                   "#0078D7")
                vals, err = self._read_wsk_zb_doc(wsk_path)
                if vals is None:
                    self.log(f"[OPIS OG] {vname}: POMINIĘTO — {err}")
                    continue

                # nazwa pliku: zachowaj pisownię z istniejącego "opis og_*", jeśli jest
                out_base = f"opis og_{vname}"
                for existing in d.glob("opis og_*"):
                    out_base = existing.name.rsplit(".", 1)[0]
                    break
                out_path = d / (out_base + ".docx")

                # formy ochrony przyrody z GDOŚ (jeśli wskazano folder)
                formy = None
                if gdos_map:
                    fx = self._gdos_find_for_village(gdos_map, vname)
                    if fx is not None:
                        try:
                            formy = self._gdos_formy_dla_wsi(fx, skrocony=skrocony)
                            n_for = (len(formy) if skrocony
                                     else sum(1 for f_ in formy if f_.startswith("- ")))
                            self.log(f"[OPIS OG] {vname}: formy ochrony z GDOŚ "
                                     f"({fx.name}; {n_for} form)"
                                     + (" — wersja skrócona (bez powiązań i opisów)"
                                        if skrocony else ""))
                        except Exception as e:
                            self.log(f"[OPIS OG] {vname}: błąd odczytu GDOŚ "
                                     f"({fx.name}) — {e}")
                    else:
                        self.log(f"[OPIS OG] {vname}: brak pliku GDOŚ — zostanie marker "
                                 "do ręcznego wpisania form ochrony.")

                try:
                    self._fill_template(tpl, vals, out_path, formy=formy,
                                        nadzor=nadzor, warunki=warunki,
                                        kategoria=kategoria)
                except Exception as e:
                    self.log(f"[OPIS OG] {vname}: BŁĄD zapisu — {e}")
                    continue

                done += 1
                self.log(f"[OPIS OG] {vname}: zapisano {out_path.name} "
                         f"(rębne {vals['MAKS_MIAZSZOSC']} m3, "
                         f"przedrębne {vals['UZYTK_PRZEDRZEBNE']} m3)")

            self.log(f"[OPIS OG] Zakończono — wygenerowano {done}/{len(jobs)}."
                     + ("" if gdos_map else
                        " Pamiętaj o ręcznym wpisaniu form ochrony przyrody "
                        "(Natura 2000 itp.) w wygenerowanych plikach."))
            self.update_status("Opisy ogólne gotowe", "#107C10", animate=False)
        finally:
            if temp_dir is not None:
                shutil.rmtree(temp_dir, ignore_errors=True)
                self.log("[OPIS OG] Usunięto pliki tymczasowe (MIETEK -> TXT -> Word).")

    # ------------------------------------------------ GDOŚ: formy ochrony przyrody

    GDOS_OBSZARY_PLIK = "gdos_obszary.json"
    # arkusz GDOŚ -> etykieta formy (None = park krajobrazowy bez etykiety)
    GDOS_SHEETY = [
        ("ObszarySpecjalnejOchronyPolygon", "OSO"),
        ("SpecjalneObszaryOchronyPolygon", "SOO"),
        ("RezerwatyPolygon", "REZERWAT"),
        ("ParkiNarodowePolygon", "PARK NARODOWY"),
        ("ParkiKrajobrazowePolygon", None),
        ("UzytkiEkologicznePolygon", "UŻYTK EKOLOGICZNY"),
        ("ZespolyPrzyrodniczoKrajobrazowe", "ZESPÓŁ PRZYRODNICZO-KRAJOBRAZOWY"),
        ("ObszaryChronionegoKrajobrazuPol", "OBSZAR CHRONIONEGO KRAJOBRAZU"),
        ("StanowiskaDokumentacyjnePolygon", "STANOWISKO DOKUMENTACYJNE"),
    ]

    @staticmethod
    def _fix_gdos_mojibake(s):
        """Naprawia nazwy z krzakami w plikach GDOŚ (np. 'ê' zamiast 'ę')."""
        if not s or all(ord(c) < 128 for c in s):
            return s
        try:
            return s.encode("cp1252").decode("cp1250")
        except (UnicodeEncodeError, UnicodeDecodeError):
            return s

    @classmethod
    def _gdos_kb(cls):
        """Baza wiedzy o obszarach (nazwa -> kod/PZO/powiązanie) z gdos_obszary.json."""
        kb = getattr(cls, "_gdos_kb_cache", None)
        if kb is not None:
            return kb
        from app.core.word_worker import get_resource_path
        kb = {}
        try:
            with open(get_resource_path(cls.GDOS_OBSZARY_PLIK), encoding="utf-8") as f:
                for o in json.load(f).get("obszary", []):
                    kb[str(o.get("nazwa", "")).strip().casefold()] = o
        except Exception:
            kb = {}
        cls._gdos_kb_cache = kb
        return kb

    @staticmethod
    def _gdos_ak(a):
        """Klucz naturalnego sortowania pododdziałów (1a < 1b < 2a < 10a)."""
        return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", a)]

    @staticmethod
    def _gdos_norm(s):
        """Nazwa bez cyfr/podkreśleń/myślników i ogonków, WIELKIMI literami (dopasowania)."""
        import unicodedata
        bez = unicodedata.normalize("NFKD", s.upper())
        return "".join(c for c in bez if not unicodedata.combining(c) and c.isalpha())

    @staticmethod
    def _gdos_files_map(folder):
        """Mapa: znormalizowana nazwa wsi -> plik *_wynik.xlsx z folderu GDOŚ."""
        folder = Path(folder)
        m = {}
        for pf in sorted(folder.rglob("*.xlsx")):
            if pf.name.startswith("~$"):
                continue
            norm = TabOpisOgMixin._gdos_norm(pf.stem)
            if norm.endswith("WYNIK"):
                norm = norm[:-5]
            if norm:
                m.setdefault(norm, pf)
        return m

    @classmethod
    def _gdos_find_for_village(cls, gdos_map, vname):
        """Dopasowuje wieś do pliku GDOŚ.

        Toleruje numerację plików, '0' zamiast 'O', brak ogonków
        oraz drobne różnice pisowni (np. CHUDOPCZYCE vs CHUDOBCZYCE).
        """
        norm = cls._gdos_norm(vname)
        if not norm:
            return None
        if norm in gdos_map:
            return gdos_map[norm]
        if len(norm) >= 5:
            # nazwa wsi zawiera się w nazwie pliku albo odwrotnie
            # (np. '0RZESZKOWO' bez zera vs wieś ORZESZKOWO)
            for k, v in gdos_map.items():
                if norm in k or k in norm:
                    return v
            # ostatnia deska: bardzo podobna pisownia (literówki)
            import difflib
            hit = difflib.get_close_matches(norm, list(gdos_map), n=1, cutoff=0.85)
            if hit:
                return gdos_map[hit[0]]
        return None

    def _gdos_formy_dla_wsi(self, xlsx_path, skrocony=False):
        """Parsuje wynik GDOŚ — zwraca listę akapitów o formach ochrony przyrody.

        skrocony=True -> tylko lista form (bez 'Powiązanie z gospodarką leśną'
        i bez opisów pozostałych form) — wersja skrócona opisu ogólnego.
        """
        import openpyxl
        wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)

        # komplet pododdziałów wsi (do stwierdzenia "w całym obszarze opracowania")
        wszystkie_a1 = set()
        if "Analiza_szczegółowa" in wb.sheetnames:
            for row in wb["Analiza_szczegółowa"].iter_rows(min_row=2, values_only=True):
                if row and row[0]:
                    a = str(row[0]).strip()
                    if a.upper() != "A1":
                        wszystkie_a1.add(a)

        kb = self._gdos_kb()
        obszary = []
        for sheet, typ in self.GDOS_SHEETY:
            if sheet not in wb.sheetnames:
                continue
            grupy = {}
            for row in wb[sheet].iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 7:
                    continue
                oddzial, a1, status, nazwa = row[0], row[1], row[2], row[6]
                pierwsza = str(oddzial or "").strip().upper()
                if pierwsza.startswith(("PODSUMOW", "WYDZIELENIA")):
                    break
                if not a1 or not nazwa:
                    continue
                a1s = str(a1).strip()
                if a1s.upper() == "A1":
                    continue
                # pododdział bez numeru oddziału -> doklej oddział, jeśli podany
                if not any(c.isdigit() for c in a1s) and oddzial:
                    a1s = f"{str(oddzial).strip()}{a1s}"
                nazwa = self._fix_gdos_mojibake(str(nazwa).strip())
                entry = kb.get(nazwa.casefold())
                if entry is None:
                    for k, v in kb.items():
                        if k[:12] == nazwa.casefold()[:12]:
                            entry = v
                            break
                if entry is None:
                    # samouzupełnianie bazy: nieznany obszar trafia od razu do
                    # gdos_obszary.json (kod/powiązanie uzupełniasz w edytorze)
                    if self._gdos_baza_dopisz(nazwa, typ or "PARK KRAJOBRAZOWY"):
                        entry = self._gdos_kb().get(nazwa.casefold())
                        self.log(f"[GDOŚ] Dopisano do bazy nowy obszar: {nazwa} "
                                 f"({typ or 'PARK KRAJOBRAZOWY'}) — uzupełnij kod "
                                 f"i powiązanie w edytorze bazy.")
                if entry is not None:
                    nazwa = entry["nazwa"]
                z = grupy.setdefault(nazwa, {"pelne": [], "czesciowe": [], "kb": entry,
                                              "typ": typ or "PARK KRAJOBRAZOWY"})
                if str(status or "").strip().upper().startswith("CZ"):
                    z["czesciowe"].append(a1s)
                else:
                    z["pelne"].append(a1s)
            for nazwa, z in grupy.items():
                obszary.append(z | {"nazwa": nazwa})

        if not obszary:
            return ["W obszarze objętym opracowaniem nie zlokalizowano "
                    "żadnych form ochrony przyrody."]

        # wersja skrócona: zwykłe zdania bez nagłówka, myślników, PZO i KODU
        # obszaru, np. „Obszar Natura 2000 SOO Ostoja Międzychodzko-Sierakowska
        # w pododdziałach 1a, 1b (w części: 1a, 1b)."
        par = [] if skrocony else ["Zlokalizowano następujące formy ochrony przyrody"]
        for z in obszary:
            typ, nazwa, kbe = z["typ"], z["nazwa"], z["kb"]
            pelne = sorted(set(z["pelne"]), key=self._gdos_ak)
            czesc = sorted(set(z["czesciowe"]), key=self._gdos_ak)
            if wszystkie_a1 and pelne and set(pelne) | set(czesc) >= wszystkie_a1:
                gdzie = "w całym obszarze opracowania"
            else:
                gdzie = ("w pododdziałach " + ", ".join(pelne + czesc)) if (pelne or czesc) else ""
                if czesc:
                    gdzie += f" (w części: {', '.join(czesc)})"
            if skrocony:
                if typ in ("OSO", "SOO"):
                    # bez kodu obszaru — samo „OSO/SOO + nazwa"
                    par.append((f"Obszar Natura 2000 {typ} {nazwa} {gdzie}"
                                if gdzie else f"Obszar Natura 2000 {typ} {nazwa}") + ".")
                else:
                    par.append((f"{nazwa} {gdzie}" if gdzie else nazwa) + ".")
                continue
            if typ in ("OSO", "SOO"):
                if kbe and kbe.get("kod"):
                    par.append(f"- Obszar Natura 2000 {typ} {nazwa} {kbe['kod']} "
                               f"(PZO: {kbe['pzo']}) {gdzie}.")
                else:
                    par.append(f"- Obszar Natura 2000 {typ} {nazwa} {gdzie}. "
                               "[TU UZUPEŁNIJ: kod obszaru i publikację PZO]")
            else:
                par.append(f"- {nazwa} {gdzie}.")
            if kbe and not skrocony:
                if typ in ("OSO", "SOO") and kbe.get("powiazanie"):
                    par.append(f"Powiązanie z gospodarką leśną - {nazwa}: {kbe['powiazanie']}")
                elif kbe.get("opis"):
                    par.append(f"{nazwa}: {kbe['opis']}")
        return par

    # ------------------------------------------------ MIETEK -> TXT -> Word (temp)

    @staticmethod
    def _find_mietek_sources(root):
        """Zwraca listę par (folder wsi, folder z danymi MIETEKA/O*.DBF).

        Obsługiwane układy (analogicznie do etapu MIETEK -> TXT):
          root/<WIEŚ>/WOL.001/O*.DBF  ->  (root/<WIEŚ>, root/<WIEŚ>/WOL.001)
          root/<WIEŚ>/O*.DBF          ->  (root/<WIEŚ>, root/<WIEŚ>)
          root/WOL.001/O*.DBF         ->  (root, root/WOL.001)   (pojedyncza wieś)
          root/O*.DBF                 ->  (root, root)             (pojedyncza wieś)
        """
        root = Path(root)
        found = {}
        for p in root.rglob("*.DBF"):
            if p.name[:1].upper() != "O":
                continue
            d = p.parent
            if d == root:
                village = root
            elif d.name.upper().startswith("WOL"):
                village = d.parent
            else:
                village = d
            if village == root or village.parent == root:
                found.setdefault(village, d)
        return sorted(found.items())

    def _mietek_to_wsk_docs(self, root, sources):
        """Robi MIETEK -> TXT -> Word w folderze tymczasowym (tylko WSK_ZB).

        sources — lista par (folder wsi, folder z danymi MIETEKA/O*.DBF).
        Źródło (pliki DBF) jest najpierw kopiowane do folderu tymczasowego,
        więc oryginalne dane MIETEKA pozostają nietknięte.
        Zwraca (temp_dir, {folder z DBF: ścieżka WSK_ZB.doc w tempie}) —
        wygenerowany opis og powinien trafić właśnie do folderu z DBF.
        Folder tymczasowy trzeba potem usunąć (shutil.rmtree).
        """
        root = Path(root)
        temp_dir = Path(tempfile.mkdtemp(prefix="forestly_opis_og_"))
        src = temp_dir / "MIETEK"
        txt_dir = temp_dir / "TXT"
        word_dir = temp_dir / "Word"
        try:
            src.mkdir(parents=True)
            for v, _dbf in sources:
                shutil.copytree(v, src / v.relative_to(root), dirs_exist_ok=True)
            self.log(f"[OPIS OG] Dane MIETEKA — wsi: {len(sources)}. Robię tymczasowo "
                     "MIETEK -> TXT -> Word (źródło pozostaje nietknięte)...")
            self.update_status("Opisy ogólne: MIETEK -> TXT (tymczasowo)...", "#0078D7")
            self.task_generuj_txt(src)
            self.check_stop()
            self.update_status("Opisy ogólne: MIETEK -> Word, plik WSK_ZB...", "#0078D7")
            self.task_clean_txt(src, txt_dir, ["WSK_ZB"])
            self._flatten_001_subfolders(txt_dir)
            self.task_word_processing_subprocess(
                txt_dir, word_dir, remove_names=False, file_filter=["WSK_ZB"],
                margins_dict=None,
            )
            village_to_dbf = dict(sources)
            wsk = {}
            for p in word_dir.rglob("WSK_ZB.doc"):
                orig_village = root / p.parent.relative_to(word_dir)
                wsk[village_to_dbf.get(orig_village, orig_village)] = p
            return temp_dir, wsk
        except Exception:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise

    # ------------------------------------------------ Zakładka w klasycznym GUI
    def setup_opis_og_tab(self, parent):
        """Zakładka „Opis ogólny" w starym GUI (Forestly_OLD)."""
        import customtkinter as ctk

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)
        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)
        card = ctk.CTkFrame(
            scroll_frame, fg_color="#252526", corner_radius=8,
            border_width=1, border_color="#333333",
        )
        card.grid(row=0, column=0, padx=20, pady=(15, 15), sticky="new")
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            card, text="Folder główny (z folderami wsi) albo folder pojedynczej wsi:",
            font=font_label, text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.opis_og_root_entry = ctk.CTkEntry(
            card, placeholder_text="np. folder \u201eu\u0142o\u017cone\u201d albo folder jednej wsi",
            height=36,
        )
        self.opis_og_root_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=lambda: self.select_dir(self.opis_og_root_entry),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))
        ctk.CTkLabel(
            card,
            text="Liczby (etaty, użytkowanie przedrębne) czytane są automatycznie z pliku "
                 "WSK_ZB.doc znajdującego się w folderze wsi. Formy ochrony przyrody "
                 "(Natura 2000 itp.) wpisujesz ręcznie — chyba że wskażesz poniżej folder "
                 "z wynikami GDOŚ, to wtedy zostaną wstawione automatycznie.",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#888888",
            wraplength=700, justify="left",
        ).grid(row=1, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="w")
        ctk.CTkLabel(
            card, text="Folder z wynikami GDOŚ (opcjonalnie):",
            font=font_label, text_color="#E0E0E0",
        ).grid(row=2, column=0, padx=15, pady=(0, 8), sticky="w")
        self.opis_og_gdos_entry = ctk.CTkEntry(
            card, placeholder_text="np. folder z plikami *_wynik.xlsx (GDOŚ)",
            height=36,
        )
        self.opis_og_gdos_entry.grid(row=2, column=1, padx=5, pady=(0, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=lambda: self.select_dir(self.opis_og_gdos_entry),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=2, column=2, padx=15, pady=(0, 8))
        ctk.CTkButton(
            scroll_frame, text="Generuj opisy ogólne", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0", hover_color="#005A9E", height=44, corner_radius=6,
            command=self.start_opis_og_pipeline,
        ).grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    # ------------------------------------------------ TAKSATOR: opisy z raportów Excel
    @staticmethod
    def _opis_og_fmt(x, nd=0):
        """Liczba m3/procent: nd=0 całkowita, nd=2 z przecinkiem (jak we wzorcach)."""
        return f"{float(x):.{nd}f}".replace(".", ",")

    @staticmethod
    def _sheet_row_po_num(df, n_kol):
        """Wiersz wartości następujący po wierszu z numeracją 1..n (pozycyjnie,
        brakujące komórki = None)."""
        for i in range(len(df) - 1):
            kom = [str(v).strip() for v in df.iloc[i].tolist()]
            nn = [v for v in kom if v and v != "nan"]
            if nn and all(re.fullmatch(r"\d+", v) for v in nn) \
                    and [int(v) for v in nn] == list(range(1, n_kol + 1)):
                out = []
                for v in df.iloc[i + 1].tolist():
                    try:
                        fv = float(str(v).replace(",", "."))
                        out.append(None if fv != fv else fv)  # NaN -> None
                    except (ValueError, TypeError):
                        out.append(None)
                while out and out[-1] is None:
                    out.pop()
                return out
        return None

    @staticmethod
    def _parse_raport_xls(path):
        """Wyciąga wartości opisu ogólnego z raportu Excel do druku.

        Arkusz 'Zestawienie', pierwsza tabela (zadania gospodarcze):
          * kody rzymskie (IB, IIIA, IIIB, IVD...) -> użytki rębne właściwe (etat),
          * PRZEST / PŁAZ -> pozostałe użytki rębne,
          * TW + TP -> użytkowanie przedrębne.
        Zwraca (values, wieś, error).
        """
        import pandas as pd
        RZYMSKI = re.compile(r"^(?:I{1,3}|IV|V|VI)[A-Z]*$")

        def _num(x):
            try:
                fv = float(str(x).replace(",", "."))
                return None if fv != fv else fv  # NaN -> None
            except (ValueError, TypeError):
                return None

        try:
            df = pd.read_excel(path, sheet_name="Zestawienie", header=None)
        except Exception as e:
            return None, None, f"nie można odczytać arkusza 'Zestawienie' ({e})"

        rebn = poz = przed = 0.0
        for _, row in df.iterrows():
            kom = [str(c) for c in row.tolist()]
            if any(k.startswith("Zestawienie zabieg") for k in kom if k and k != "nan"):
                break  # koniec pierwszej tabeli
            kod = kom[1].strip().upper() if len(kom) > 1 else ""
            if not kod or kod == "NAN":
                continue
            brutto = _num(kom[3]) if len(kom) > 3 else None
            if brutto is None:
                brutto = _num(kom[2]) if len(kom) > 2 else None
            if brutto is None:
                continue
            if RZYMSKI.match(kod):
                rebn += brutto
            elif kod in ("TW", "TP"):
                przed += brutto
            elif kod in ("PRZEST", "PŁAZ", "POZ"):
                poz += brutto

        # 2) nazwa wsi (TPM_FL)
        wieś = None
        try:
            tpm = pd.read_excel(path, sheet_name="TPM_FL", header=None)
            for _, row in tpm.iterrows():
                for v in row.tolist():
                    s = str(v)
                    if s.startswith("Obiekt (obr. ew.):") and " - " in s:
                        wieś = s.split(" - ", 1)[1].strip().rstrip(".").strip()
                        break
                if wieś:
                    break
        except Exception:
            pass

        # 3) arkusze Etaty / Przedrebne (raporty nietknięte) + linia obrębu
        etaty = przedr = None
        obreb = None
        try:
            et = pd.read_excel(path, sheet_name="Etaty", header=None)
            etaty = TabOpisOgMixin._sheet_row_po_num(et, 6)
            for _, row in et.iterrows():
                for v in row.tolist():
                    sx = str(v).strip()
                    if re.match(r"^\d+[-/]\d+[-/]\d+[-/]\d+\s*-\s*.+", sx):
                        obreb = sx
                        break
                if obreb:
                    break
        except Exception:
            pass
        try:
            pr = pd.read_excel(path, sheet_name="Przedrebne", header=None)
            przedr = TabOpisOgMixin._sheet_row_po_num(pr, 5)
        except Exception:
            pass
        if obreb is None:
            try:
                tpm = pd.read_excel(path, sheet_name="TPM_FL", header=None)
                for _, row in tpm.iterrows():
                    for v in row.tolist():
                        s = str(v)
                        if s.startswith("Obiekt (obr. ew.):") and " - " in s:
                            czesci = s.split(":", 1)[1].strip().split(" - ", 1)
                            obreb = f"{czesci[0].strip()} - " \
                                    f"{czesci[1].strip().rstrip('.').strip()}"
                            break
                    if obreb:
                        break
            except Exception:
                pass

        if etaty is None and przedr is None and rebn == 0 and poz == 0 and przed == 0:
            return None, wieś, "brak zadań gospodarczych w arkuszu 'Zestawienie'"

        _l = TabOpisOgMixin._opis_og_fmt
        vals = {}
        if obreb:
            vals["OBREB"] = obreb
        if etaty:
            for i, key in enumerate(("ETAT_OST_KL", "ETAT_2_OST", "ETAT_POTRZEBY",
                                     "ETAT_PRZYJETY", "POZOSTALE_REBNE")):
                if i < len(etaty) and etaty[i] is not None:
                    vals[key] = _l(etaty[i])
        vals.setdefault("ETAT_POTRZEBY", _l(rebn))
        vals.setdefault("ETAT_PRZYJETY", _l(rebn))
        vals.setdefault("POZOSTALE_REBNE", _l(poz))
        # maksymalna = etat przyjęty + pozostałe użytki rębne (wartość z arkusza
        # Etaty bywa błędna, gdy "Etat przyjęty" jest tam pusty)
        _f = lambda x: float(str(x).replace(",", "."))
        vals["MAKS_MIAZSZOSC"] = _l(_f(vals["ETAT_PRZYJETY"])
                                    + _f(vals["POZOSTALE_REBNE"]))
        if przedr:
            vals["PRZEDRZ_MIAZSZ"] = _l(przedr[0])
            if len(przedr) > 1 and przedr[1] is not None:
                vals["PRZEDRZ_PRZYROST"] = _l(przedr[1], nd=2)
            if len(przedr) > 2 and przedr[2] is not None:
                vals["UZYTK_PRZEDRZEBNE"] = _l(przedr[2])
            if len(przedr) > 3 and przedr[3] is not None:
                vals["PRZEDRZ_P1"] = _l(przedr[3], nd=2)
            if len(przedr) > 4 and przedr[4] is not None:
                vals["PRZEDRZ_P2"] = _l(przedr[4], nd=2)
        vals.setdefault("UZYTK_PRZEDRZEBNE", _l(przed))
        return vals, wieś, None

    def start_opis_og_taksator_pipeline(self):
        """Zadanie z mapy zadań (web) / przycisk (stare GUI) — wariant TAKSATOR."""
        self._disable_ui_for_process()

        def _run():
            try:
                self._opis_og_taksator_run()
            except Exception:
                self.log("[OPIS OG/TAKSATOR BŁĄD] " + traceback.format_exc())
                self.update_status("Błąd", "#D83B01", animate=False)
            finally:
                self.restore_all_buttons()

        threading.Thread(target=_run, daemon=True).start()

    def _opis_og_taksator_run(self):
        from app.core.word_worker import get_resource_path

        entry = getattr(self, "opis_og_taksator_entry", None)
        raw = entry.get().strip() if entry else ""
        if not raw or not Path(raw).exists():
            self.log("[OPIS OG/TAKSATOR] Wskaż najpierw folder z raportami Excel do druku.")
            self.update_status("Brak folderu", "#D83B01", animate=False)
            return
        root = Path(raw)
        self.last_output_dir = root

        xlsy = [p for p in sorted(root.rglob("*.xls")) + sorted(root.rglob("*.xlsx"))
                if not p.name.startswith("~$") and not p.name.lower().endswith(".dbf")]
        if not xlsy:
            self.log(f"[OPIS OG/TAKSATOR] Brak plików Excel (.xls/.xlsx) w: {root}")
            self.update_status("Brak raportów", "#D83B01", animate=False)
            return

        tpl = get_resource_path(TPL_TAKSATOR_FILENAME)
        if not Path(tpl).exists():
            self.log(f"[OPIS OG/TAKSATOR BŁĄD] Brak szablonu taksatora: {tpl} "
                     f"(plik opis_og_szablon_taksator.docx musi być w głównym folderze)")
            self.update_status("Brak szablonu", "#D83B01", animate=False)
            return
            self.update_status("Brak szablonu", "#D83B01", animate=False)
            return

        # opcjonalny folder z wynikami GDOŚ -> automatyczne formy ochrony przyrody
        gdos_entry = getattr(self, "opis_og_taksator_gdos_entry", None)
        gdos_raw = gdos_entry.get().strip() if gdos_entry is not None else ""
        gdos_map = {}
        if gdos_raw:
            if Path(gdos_raw).exists():
                gdos_map = self._gdos_files_map(gdos_raw)
                if gdos_map:
                    self.log(f"[OPIS OG/TAKSATOR] Wyniki GDOŚ: {len(gdos_map)} plik(ów) — "
                             "formy ochrony przyrody zostaną wstawione automatycznie.")
            else:
                self.log(f"[OPIS OG/TAKSATOR] Folder GDOŚ nie istnieje: {gdos_raw} — pomijam.")

        self.log(f"[OPIS OG/TAKSATOR] Generator start — raportów: {len(xlsy)}, "
                 f"szablon: {Path(tpl).name}")
        done = 0
        for i, xls in enumerate(xlsy, 1):
            self.check_stop()
            vals, wieś, err = self._parse_raport_xls(xls)
            if vals is None:
                self.log(f"[OPIS OG/TAKSATOR] {xls.name}: POMINIĘTO — {err}")
                continue
            if not wieś:
                wieś = xls.stem.split("-")[-1].strip()
            self.update_status(f"Opis ogólny: {wieś} ({i}/{len(xlsy)})", "#0078D7")

            # nazwa pliku: zachowaj pisownię istniejącego "opis og_*" dla tej wsi
            out_name = f"opis og_{wieś}.docx"
            for existing in xls.parent.glob("opis og_*"):
                if existing.stem.split("og_", 1)[-1].casefold() == wieś.casefold():
                    out_name = existing.name
                    break
            out_path = xls.parent / out_name

            formy = None
            if gdos_map:
                fx = self._gdos_find_for_village(gdos_map, wieś)
                if fx is not None:
                    try:
                        formy = self._gdos_formy_dla_wsi(fx)
                        n_for = sum(1 for f_ in formy if f_.startswith("- "))
                        self.log(f"[OPIS OG/TAKSATOR] {wieś}: formy ochrony z GDOŚ "
                                 f"({fx.name}; {n_for} form)")
                    except Exception as e:
                        self.log(f"[OPIS OG/TAKSATOR] {wieś}: błąd odczytu GDOŚ "
                                 f"({fx.name}) — {e}")

            try:
                self._fill_template(tpl, vals, out_path, formy=formy)
            except Exception as e:
                self.log(f"[OPIS OG/TAKSATOR] {wieś}: BŁĄD zapisu — {e}")
                continue

            done += 1
            self.log(f"[OPIS OG/TAKSATOR] {wieś}: zapisano {out_path.name} "
                     f"(rębne {vals['MAKS_MIAZSZOSC']} m3, "
                     f"przedrębne {vals['UZYTK_PRZEDRZEBNE']} m3)")

        self.log(f"[OPIS OG/TAKSATOR] Zakończono — wygenerowano {done}/{len(xlsy)}."
                 + ("" if gdos_map else
                    " Pamiętaj o ręcznym wpisaniu form ochrony przyrody "
                    "(Natura 2000 itp.) w wygenerowanych plikach."))
        self.update_status("Opisy ogólne gotowe", "#107C10", animate=False)

    def setup_opis_og_taksator_tab(self, parent):
        """Zakładka „Opisy ogólne" (TAKSATOR) w starym GUI (Forestly_OLD)."""
        import customtkinter as ctk

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)
        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)
        card = ctk.CTkFrame(
            scroll_frame, fg_color="#252526", corner_radius=8,
            border_width=1, border_color="#333333",
        )
        card.grid(row=0, column=0, padx=20, pady=(15, 15), sticky="new")
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            card, text="Folder z raportami Excel do druku:",
            font=font_label, text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.opis_og_taksator_entry = ctk.CTkEntry(
            card, placeholder_text="np. folder z plikami 042-0001-Bysław.xls",
            height=36,
        )
        self.opis_og_taksator_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=lambda: self.select_dir(self.opis_og_taksator_entry),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))
        ctk.CTkLabel(
            card,
            text="Liczby (etat przyjęty, pozostałe użytki rębne, użytkowanie przedrębne) "
                 "czytane są z arkusza 'Zestawienie' każdego raportu — na podstawie zadań "
                 "gospodarczych (rębne, PRZEST/PŁAZ, TW+TP).",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#888888",
            wraplength=700, justify="left",
        ).grid(row=1, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="w")
        ctk.CTkLabel(
            card, text="Folder z wynikami GDOŚ (opcjonalnie):",
            font=font_label, text_color="#E0E0E0",
        ).grid(row=2, column=0, padx=15, pady=(0, 8), sticky="w")
        self.opis_og_taksator_gdos_entry = ctk.CTkEntry(
            card, placeholder_text="np. folder z plikami *_wynik.xlsx (GDOŚ)",
            height=36,
        )
        self.opis_og_taksator_gdos_entry.grid(row=2, column=1, padx=5, pady=(0, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=lambda: self.select_dir(self.opis_og_taksator_gdos_entry),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=2, column=2, padx=15, pady=(0, 8))
        ctk.CTkButton(
            scroll_frame, text="Generuj opisy ogólne", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0", hover_color="#005A9E", height=44, corner_radius=6,
            command=self.start_opis_og_taksator_pipeline,
        ).grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")


    # ------------------------------------------------ Baza GDOŚ: zapis/dopisywanie

    @classmethod
    def _gdos_baza_sciezka(cls):
        """Plik zapisu gdos_obszary.json (obok EXE w wersji przenośnej)."""
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent / cls.GDOS_OBSZARY_PLIK
        return Path(__file__).resolve().parents[3] / cls.GDOS_OBSZARY_PLIK

    @classmethod
    def _gdos_baza_wczytaj(cls):
        """Lista obszarów z gdos_obszary.json (pusta lista przy braku pliku)."""
        from app.core.word_worker import get_resource_path
        try:
            with open(get_resource_path(cls.GDOS_OBSZARY_PLIK), encoding="utf-8") as f:
                return json.load(f).get("obszary", [])
        except Exception:
            return []

    @classmethod
    def _gdos_baza_zapisz(cls, obszary):
        """Zapis bazy obszarów (zachowuje _opis) + odświeżenie cache generatora."""
        from app.core.word_worker import get_resource_path
        opis = ("Baza obszarów ochrony przyrody dla generatora opisów "
                "ogólnych (folder z wynikami GDOŚ).")
        try:
            with open(get_resource_path(cls.GDOS_OBSZARY_PLIK), encoding="utf-8") as f:
                opis = json.load(f).get("_opis", opis)
        except Exception:
            pass
        czyste = []
        for r in obszary:
            nazwa = str(r.get("nazwa", "")).strip()
            if not nazwa:
                continue
            czyste.append({k: str(r.get(k, "") or "").strip()
                           for k in ("nazwa", "typ", "kod", "pzo",
                                     "powiazanie", "opis")})
        with open(cls._gdos_baza_sciezka(), "w", encoding="utf-8") as f:
            json.dump({"_opis": opis, "obszary": czyste},
                      f, ensure_ascii=False, indent=1)
        cls._gdos_kb_cache = None  # odśwież cache generatora opisów
        return len(czyste)

    @classmethod
    def _gdos_baza_dopisz(cls, nazwa, typ):
        """Dopisuje nieznany obszar do bazy (kod/powiązanie uzupełniasz potem).

        Zwraca True, gdy obszar był nowy i został dodany."""
        if not nazwa:
            return False
        klucz = nazwa.casefold()
        if klucz in cls._gdos_kb():
            return False
        if any(o.get("nazwa", "").casefold() == klucz
               for o in cls._gdos_baza_wczytaj()):
            return False
        cls._gdos_baza_zapisz(cls._gdos_baza_wczytaj() + [
            {"nazwa": nazwa, "typ": typ or "", "kod": "", "pzo": "",
             "powiazanie": "", "opis": ""}])
        return True

    # ------------------------------------------------ Baza GDOŚ: Excel import/eksport

    GDOS_XLSX_NAGLOWKI = [
        ("nazwa", "Nazwa"), ("typ", "Typ"), ("kod", "Kod"),
        ("pzo", "Publikacja PZO"),
        ("powiazanie", "Powiązanie z gospodarką leśną"),
        ("opis", "Opis"),
    ]

    def gdos_exportuj_excel(self, xlsx_path):
        """Eksport bazy obszarów do Excela (do masowej edycji)."""
        import openpyxl
        from openpyxl.utils import get_column_letter
        obszary = self._gdos_baza_wczytaj()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Baza GDOŚ"
        ws.append([n for _, n in self.GDOS_XLSX_NAGLOWKI])
        for o in obszary:
            ws.append([str(o.get(k, "") or "") for k, _ in self.GDOS_XLSX_NAGLOWKI])
        for i, w in enumerate((44, 10, 14, 34, 62, 46), 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = openpyxl.styles.Alignment(wrap_text=True,
                                                         vertical="top")
        wb.save(str(xlsx_path))
        return len(obszary)

    def gdos_importuj_excel(self, xlsx_path):
        """Import bazy z Excela (kolumny rozpoznawane po nagłówkach).

        Zwraca liczbę wczytanych obszarów; rzuca ValueError przy złym pliku."""
        import openpyxl
        wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
        ws = wb.active
        wiersze = list(ws.iter_rows(values_only=True))
        if not wiersze:
            raise ValueError("plik jest pusty")
        nagl = [str(c).strip().lower() if c is not None else "" for c in wiersze[0]]
        mapa = {}
        for k, n in self.GDOS_XLSX_NAGLOWKI:
            for i, h in enumerate(nagl):
                if h and (h == n.lower() or h == k):
                    mapa[k] = i
                    break
        if "nazwa" not in mapa:
            raise ValueError("w pierwszym wierszu brak kolumny 'Nazwa'")
        obszary = []
        for w in wiersze[1:]:
            if not w:
                continue
            d = {k: ("" if i >= len(w) or w[i] is None else str(w[i]).strip())
                 for k, i in mapa.items()}
            if d.get("nazwa"):
                obszary.append(d)
        if not obszary:
            raise ValueError("brak wierszy z nazwą obszaru")
        return self._gdos_baza_zapisz(obszary)

    # ------------------------------------------------ Baza GDOŚ: edytor w starym GUI

    def setup_gdos_editor_tab(self, parent):
        """Zakładka „Baza obszarów GDOŚ" w starym GUI (Forestly_OLD)."""
        import customtkinter as ctk

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            scroll, text="Baza obszarów ochrony przyrody (gdos_obszary.json)",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=20, pady=(15, 4), sticky="w")
        ctk.CTkLabel(
            scroll,
            text="Kody, publikacje PZO i powiązania z gospodarką leśną wstawiane "
                 "automatycznie do opisów ogólnych. Obszary nieobecne w bazie są "
                 "dopisywane samoczynnie przy generowaniu opisów z wyników GDOŚ "
                 "(zostaje im do uzupełnienia kod i powiązanie).",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#888888",
            wraplength=760, justify="left",
        ).grid(row=1, column=0, padx=20, sticky="w")

        bar = ctk.CTkFrame(scroll, fg_color="transparent")
        bar.grid(row=2, column=0, padx=20, pady=12, sticky="ew")
        ctk.CTkButton(
            bar, text="Dodaj obszar", height=32,
            fg_color="#333333", hover_color="#444444",
            command=self._gdos_edytor_dodaj,
        ).pack(side="left")
        ctk.CTkButton(
            bar, text="Zapisz zmiany", height=32,
            fg_color="#0067C0", hover_color="#005A9E",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=self._gdos_edytor_zapisz,
        ).pack(side="left", padx=8)
        ctk.CTkButton(
            bar, text="Eksportuj do Excela", height=32,
            fg_color="#333333", hover_color="#444444",
            command=self._gdos_edytor_eksport,
        ).pack(side="right")
        ctk.CTkButton(
            bar, text="Importuj z Excela", height=32,
            fg_color="#333333", hover_color="#444444",
            command=self._gdos_edytor_import,
        ).pack(side="right", padx=8)

        # --- pasek szukania (duże bazy: bez niego 40 tys. kart zamraża GUI) ---
        self._gdos_edytor_dane = []
        self._gdos_edytor_filtr = ""
        self._gdos_edytor_limit = 100
        szukaj = ctk.CTkFrame(scroll, fg_color="transparent")
        szukaj.grid(row=3, column=0, padx=20, pady=(0, 8), sticky="ew")
        ctk.CTkLabel(szukaj, text="Szukaj:", font=ctk.CTkFont(family="Segoe UI", size=12),
                     text_color="#888888").pack(side="left")
        self._gdos_szukaj_entry = ctk.CTkEntry(
            szukaj, width=280, height=30,
            placeholder_text="nazwa, kod lub typ obszaru…")
        self._gdos_szukaj_entry.pack(side="left", padx=8)
        self._gdos_szukaj_entry.bind(
            "<KeyRelease>", lambda _e: self._gdos_edytor_szukaj())
        self._gdos_szukaj_wiecej_btn = ctk.CTkButton(
            szukaj, text="Pokaż więcej (+100)", height=30,
            fg_color="#333333", hover_color="#444444",
            command=self._gdos_edytor_wiecej)
        self._gdos_szukaj_wiecej_btn.pack(side="left", padx=8)
        self._gdos_edytor_info = ctk.CTkLabel(
            szukaj, text="", font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#888888")
        self._gdos_edytor_info.pack(side="right")

        self._gdos_edytor_karty = []
        self._gdos_edytor_kontener = ctk.CTkFrame(scroll, fg_color="transparent")
        self._gdos_edytor_kontener.grid(row=4, column=0, padx=0, pady=(0, 20), sticky="ew")
        self._gdos_edytor_kontener.grid_columnconfigure(0, weight=1)
        self._gdos_edytor_wczytaj()

    def _gdos_edytor_karta(self, dane=None):
        """Jedna karta edycji obszaru w edytorze (OLD)."""
        import customtkinter as ctk
        dane = dane or {}
        fr = ctk.CTkFrame(self._gdos_edytor_kontener, fg_color="#252526",
                          corner_radius=8, border_width=1, border_color="#333333")
        fr.grid(row=len(self._gdos_edytor_karty), column=0, padx=20, pady=6, sticky="ew")
        fr.grid_columnconfigure((0, 2, 4, 6), weight=1)
        font_lbl = ctk.CTkFont(family="Segoe UI", size=12)

        karta = {"frame": fr}
        for i, (klucz, etykieta) in enumerate((
                ("nazwa", "Nazwa:"), ("typ", "Typ (OSO/SOO/PARK KRAJOBRAZOWY...):"),
                ("kod", "Kod (np. PLB300015):"), ("pzo", "Publikacja PZO:"))):
            kol = i * 2
            ctk.CTkLabel(fr, text=etykieta, font=font_lbl,
                         text_color="#888888").grid(
                row=0, column=kol, padx=(12, 4), pady=(8, 2), sticky="w")
            e = ctk.CTkEntry(fr, height=30)
            e.insert(0, str(dane.get(klucz, "") or ""))
            e.grid(row=1, column=kol, padx=(12, 4), pady=(0, 6), sticky="ew")
            karta[klucz] = e

        ctk.CTkLabel(fr, text="Powiązanie z gospodarką leśną (dla Natura 2000):",
                     font=font_lbl, text_color="#888888").grid(
            row=2, column=0, columnspan=8, padx=12, pady=(4, 2), sticky="w")
        tb1 = ctk.CTkTextbox(fr, height=88)
        tb1.insert("1.0", str(dane.get("powiazanie", "") or ""))
        tb1.grid(row=3, column=0, columnspan=7, padx=12, pady=(0, 6), sticky="ew")
        karta["powiazanie"] = tb1

        ctk.CTkLabel(fr, text="Opis (dla pozostałych form):",
                     font=font_lbl, text_color="#888888").grid(
            row=4, column=0, columnspan=7, padx=12, sticky="w")
        tb2 = ctk.CTkTextbox(fr, height=64)
        tb2.insert("1.0", str(dane.get("opis", "") or ""))
        tb2.grid(row=5, column=0, columnspan=7, padx=12, pady=(0, 8), sticky="ew")
        karta["opis"] = tb2
        ctk.CTkButton(
            fr, text="Usuń", width=70, height=28,
            fg_color="#5c1f1f", hover_color="#7a2a2a",
            command=lambda k=None: self._gdos_edytor_usun(
                k or [c for c in self._gdos_edytor_karty if c["frame"] is fr][0]),
        ).grid(row=5, column=7, padx=(4, 12), pady=(0, 8), sticky="e")
        # edycje kart od razu trafiają do danych w pamięci — dzięki temu
        # zapis zbiera CAŁĄ bazę, a nie tylko wyrenderowane karty
        karta["dane"] = dane
        for klucz in ("nazwa", "typ", "kod", "pzo"):
            w = karta[klucz]
            w.bind("<KeyRelease>",
                   lambda _e, k=klucz, ww=w: dane.__setitem__(k, ww.get()))
        for klucz in ("powiazanie", "opis"):
            w = karta[klucz]
            w.bind("<KeyRelease>",
                   lambda _e, k=klucz, ww=w: dane.__setitem__(
                       k, ww.get("1.0", "end-1c")))
        self._gdos_edytor_karty.append(karta)

    def _gdos_edytor_wczytaj(self):
        """Wczytuje całą bazę do pamięci i renderuje (z limitem) karty."""
        self._gdos_edytor_dane = self._gdos_baza_wczytaj()
        self._gdos_edytor_render()

    def _gdos_edytor_szukaj(self):
        self._gdos_edytor_filtr = self._gdos_szukaj_entry.get().strip().lower()
        self._gdos_edytor_limit = 100
        self._gdos_edytor_render()

    def _gdos_edytor_wiecej(self):
        self._gdos_edytor_limit += 100
        self._gdos_edytor_render()

    def _gdos_edytor_render(self):
        """Rysuje tylko przefiltrowane + pierwsze N kart (zabezpieczenie
        przed zamrożeniem GUI przy bazie z tysięcy obszarów)."""
        for karta in getattr(self, "_gdos_edytor_karty", []):
            karta["frame"].destroy()
        self._gdos_edytor_karty = []
        f = getattr(self, "_gdos_edytor_filtr", "")
        widok = [d for d in self._gdos_edytor_dane
                 if not f or f in " ".join(
                     str(d.get(k, "") or "").lower()
                     for k in ("nazwa", "kod", "typ"))]
        for d in widok[:getattr(self, "_gdos_edytor_limit", 100)]:
            self._gdos_edytor_karta(d)
        lim = getattr(self, "_gdos_edytor_limit", 100)
        txt = (f"Pokazano {min(len(widok), lim)} z {len(widok)} obszarów"
               + (f" (cała baza: {len(self._gdos_edytor_dane)})" if f else ""))
        if hasattr(self, "_gdos_edytor_info"):
            self._gdos_edytor_info.configure(text=txt)
        if hasattr(self, "_gdos_szukaj_wiecej_btn"):
            self._gdos_szukaj_wiecej_btn.configure(
                state="normal" if len(widok) > lim else "disabled")

    def _gdos_edytor_dodaj(self):
        dane = {}
        self._gdos_edytor_dane.append(dane)
        self._gdos_edytor_karta(dane)
        ostatnia = self._gdos_edytor_karty[-1]
        ostatnia["nazwa"].focus()

    def _gdos_edytor_usun(self, karta):
        dane = karta.get("dane")
        if dane is not None:
            try:
                self._gdos_edytor_dane.remove(dane)
            except ValueError:
                pass
        try:
            self._gdos_edytor_karty.remove(karta)
        except ValueError:
            pass
        karta["frame"].destroy()
        self._gdos_edytor_render()

    def _gdos_edytor_zapisz(self):
        # baza w pamięci jest źródłem prawdy (karty synchronizują ją na żywo),
        # więc zapis obejmuje także wiersze ukryte przez filtr/limit
        obszary = [{
            "nazwa": str(d.get("nazwa", "") or "").strip(),
            "typ": str(d.get("typ", "") or "").strip(),
            "kod": str(d.get("kod", "") or "").strip(),
            "pzo": str(d.get("pzo", "") or "").strip(),
            "powiazanie": str(d.get("powiazanie", "") or "").strip(),
            "opis": str(d.get("opis", "") or "").strip(),
        } for d in getattr(self, "_gdos_edytor_dane", [])]
        try:
            n = self._gdos_baza_zapisz(obszary)
            self.log(f"[GDOŚ] Zapisano bazę obszarów: {n} pozycji.")
            self.update_status("Baza zapisana", "#107C10", animate=False)
        except Exception as e:
            self.log(f"[GDOŚ] Błąd zapisu bazy: {e}")
            self.update_status("Błąd zapisu", "#D83B01", animate=False)

    def _gdos_edytor_eksport(self):
        from tkinter import filedialog
        sciezka = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arkusz Excel", "*.xlsx")],
            initialfile="gdos_obszary.xlsx",
            title="Eksport bazy GDOŚ do Excela",
        )
        if not sciezka:
            return
        try:
            n = self.gdos_exportuj_excel(Path(sciezka))
            self.log(f"[GDOŚ] Wyeksportowano {n} obszarów → {sciezka}")
            self.update_status("Baza wyeksportowana", "#107C10", animate=False)
        except Exception as e:
            self.log(f"[GDOŚ] Błąd eksportu: {e}")
            self.update_status("Błąd eksportu", "#D83B01", animate=False)

    def _gdos_edytor_import(self):
        from tkinter import filedialog
        sciezka = filedialog.askopenfilename(
            filetypes=[("Arkusz Excel", "*.xlsx")],
            title="Import bazy GDOŚ z Excela",
        )
        if not sciezka:
            return
        try:
            n = self.gdos_importuj_excel(Path(sciezka))
            self.log(f"[GDOŚ] Zaimportowano {n} obszarów z pliku: {sciezka}")
            self.update_status("Baza zaimportowana", "#107C10", animate=False)
            self._gdos_edytor_wczytaj()
        except Exception as e:
            self.log(f"[GDOŚ] Błąd importu: {e}")
            self.update_status("Błąd importu", "#D83B01", animate=False)
