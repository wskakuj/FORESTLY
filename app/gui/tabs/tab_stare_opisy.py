# -*- coding: utf-8 -*-
"""Zakładka MIETEK | Stare -> nowe szablony.

Przerabia WSZYSTKIE stare pliki Word z Pełnego Automatu na nowe szablony:

  - raporty OPTAX, REJESTR1, TAB_KLW3, WSKAZ1, WSK_ZB, WYK_NEG, ZEST1,
    HALIZNY — stary Word zawiera tę samą treść co TXT mietka (ramki │...│
    zapisane jako tekst w Wordzie), więc program wyciąga z niego linie,
    zapisuje jako TXT i przepuszcza przez zwykłe generatory nowych
    szablonów (szablony.py): HTML + PDF,
  - opisy ogólne „opis og_<wieś>.doc/.docx" — liczby (etaty, użytkowanie
    przedrębne) i formy ochrony czytane ze starego dokumentu, generacja
    na szablonie opis_og_szablon.docx (jak w zakładce „Opisy ogólne"),
  - STR_TYT — opcjonalnie: jeśli wskażesz swój nowy szablon strony
    tytułowej, program wypełni go danymi (nazwa wsi z „Obiekt:", powierzchnia
    „Razem" z OPTAX) i zapisze jako STR_TYT_<wieś>.docx.

Oryginalne pliki zostają nietknięte — wszystko ląduje w folderze docelowym:
  <out>/<wieś>/opis og_<wieś>.docx
  <out>/<wieś>/STR_TYT_<wieś>.docx
  <out>/<wieś>/nowe szablony/<TYP>.html
  <out>/<wieś>/nowe szablony/pdf/<TYP>.pdf
  <out>/<wieś>/nowe szablony/pdf/STR_TYT_<wieś>.pdf   (strona tytułowa jako PDF)
  <out>/<wieś>/nowe szablony/pdf/opis og_<wieś>.pdf (opis og na nowym szablonie)

Zależności: python-docx, szablony.py, przeglądarka (Edge/Chrome) do PDF;
pliki .doc są tymczasowo konwertowane na .docx przez Worda COM
(pythoncom.CoInitialize w wątku).
"""

import re
import tempfile
import threading
import traceback
from pathlib import Path

try:
    import pythoncom
except ImportError:            # testy na Linuxie
    pythoncom = None

TPL_STARE_FILENAME = "opis_og_szablon.docx"   # ten sam szablon co "Opisy ogólne"

# raporty, które umie wyrenderować szablony.py (stara nazwa pliku -> typ)
TYPY_RAPORTOW = ("OPTAX", "REJESTR1", "TAB_KLW3", "WSKAZ1", "WSK_ZB",
                 "WYK_NEG", "ZEST1", "HALIZNY")

# znaki typograficzne Worda, których nie ma w cp852 (format TXT mietka)
_TRANSLIT = str.maketrans({
    "\u201e": '"', "\u201d": '"', "\u201c": '"',
    "\u2019": "'", "\u2018": "'",
    "\u2013": "-", "\u2014": "-",
    "\u2026": "...", "\xa0": " ", "\u202f": " ",
})


class TabStareOpisyMixin:
    """Konwersja starych plików Word (Pełny Automat) na nowe szablony."""

    # ------------------------------------------------ ekstrakcja tekstu z Worda

    @staticmethod
    def _stare_linie(doc):
        """Wszystkie linie tekstu z .docx — także z dokumentów, w których
        cała treść siedzi w JEDNYM akapicie rozdzielonym <w:br> (tak były
        robione REJESTR1/WYK_NEG), z ramkami tekstowymi włącznie.

        <w:br>/<w:cr> -> nowa linia, <w:tab> -> spacja (TXT mietka jest
        wyrównany spacjami, nie tabulatorami).
        """
        from docx.oxml.ns import qn
        out = []
        for p in doc.element.body.iter(qn("w:p")):
            buf = []
            for el in p.iter():
                if el.tag == qn("w:t"):
                    buf.append(el.text or "")
                elif el.tag in (qn("w:br"), qn("w:cr")):
                    buf.append("\n")
                elif el.tag == qn("w:tab"):
                    buf.append(" ")
            out.extend("".join(buf).split("\n"))
        return out

    @classmethod
    def _stare_na_txt(cls, linie, txt_path):
        """Linie -> plik TXT w kodowaniu cp852 (czytany potem przez
        szablony.wczytaj jak zwykły TXT z mietka)."""
        tekst = "\n".join(linie).translate(_TRANSLIT)
        try:
            dane = tekst.encode("cp852")
        except UnicodeEncodeError:
            dane = tekst.encode("cp852", errors="replace")
        Path(txt_path).write_bytes(dane)
        return Path(txt_path)

    # ------------------------------------------------ parsowanie starego opisu og

    @staticmethod
    def _stare_liczba(s):
        """'1 234' / '1234,5' / '1234.5' -> float (albo None)."""
        s = re.sub(r"[\s\xa0\u202f]", "", str(s or ""))
        if not s:
            return None
        s = s.replace(",", ".")
        if not re.fullmatch(r"[\d.]+", s):
            return None
        try:
            return float(s)
        except ValueError:
            return None

    @staticmethod
    def _stare_etat_z_tabeli(doc):
        """Tabela etatów ze starego opisu: nagłówki wg słów kluczowych.

        Zwraca dict {ETAT_POTRZEBY, ETAT_PRZYJETY, POZOSTALE_REBNE,
        MAKS_MIAZSZOSC, (ETAT_OST_KL, ETAT_2_OST)} albo None.
        """
        for tabela in doc.tables:
            if len(tabela.rows) < 2:
                continue
            naglowki = [c.text.strip().lower() for c in tabela.rows[0].cells]
            if not any("etat" in h or "maksymalna" in h for h in naglowki):
                continue
            kolumny = {}
            for i, h in enumerate(naglowki):
                if "2-ch" in h or ("dw" in h and "ostatnich" in h):
                    kolumny["ETAT_2_OST"] = i
                elif "ostatniej" in h:
                    kolumny["ETAT_OST_KL"] = i
                elif "potrzeb" in h:
                    kolumny["ETAT_POTRZEBY"] = i
                elif "przyj" in h:
                    kolumny["ETAT_PRZYJETY"] = i
                elif "pozosta" in h:
                    kolumny["POZOSTALE_REBNE"] = i
                elif "maksymalna" in h:
                    kolumny["MAKS_MIAZSZOSC"] = i
            if "ETAT_PRZYJETY" not in kolumny:
                continue  # to nie jest tabela etatów
            wiersz = [c.text.strip() for c in tabela.rows[1].cells]
            vals = {}
            for klucz, i in kolumny.items():
                if i < len(wiersz):
                    n = TabStareOpisyMixin._stare_liczba(wiersz[i])
                    if n is not None:
                        vals[klucz] = n
            if vals:
                return vals
        return None

    @staticmethod
    def _stare_przedrzedne(doc):
        """'Użytkowanie przedrębne - 150 m3' -> 150.0 (albo None)."""
        for p in doc.paragraphs:
            t = p.text
            niskie = t.lower().replace(" ", "")
            if "przedr" not in niskie or "m3" not in niskie:
                continue
            m = re.search(r"([\d][\d\s\xa0.,]*)\s*m\s*3", t)
            if m:
                n = TabStareOpisyMixin._stare_liczba(m.group(1))
                if n is not None:
                    return n
        return None

    @staticmethod
    def _stare_formy(doc):
        """Formy ochrony przyrody z sekcji OCHRONA starego opisu.

        Zwraca listę akapitów w formacie jak z wyników GDOŚ:
        [0] = "Zlokalizowano następujące formy ochrony przyrody",
        dalej pozycje "- ...". None = brak form.

        UWAGA na "Lasy ochronne": nowy szablon ma już statyczną linię
        „Lasy ochronne – 0,00ha" — linię z zerową powierzchnią pomijamy
        (dublowałaby się), a niezerową przepisujemy (i sygnalizujemy
        w logu, bo statyczną linię trzeba wtedy poprawić ręcznie).
        """
        paras = [p.text.strip() for p in doc.paragraphs]
        lasy_niezero = False
        for i, t in enumerate(paras):
            if t.startswith("Zlokalizowano"):
                formy = [t]
                for nast in paras[i + 1:]:
                    if not nast:
                        continue
                    n = nast.lstrip()
                    if n.startswith(("-", "•", "·")):
                        formy.append(nast)
                    elif n.lower().startswith("lasy ochronne"):
                        cyfry = [float(x.replace(",", "."))
                                 for x in re.findall(r"\d+[.,]\d+|\d+", n)]
                        if any(c > 0 for c in cyfry):
                            lasy_niezero = True
                            formy.append(nast)
                        # zero — pomijamy: to samo co statyczna linia szablonu
                    else:
                        break
                return formy, lasy_niezero
        # stare opisy bez form ochrony przyrody
        if any("nie zlokalizowano" in t.lower() for t in paras):
            return None, False
        return None, False

    @staticmethod
    def _stare_wies(path, root):
        """Nazwa wsi z nazwy pliku 'opis og_<wieś>' (albo z folderu)."""
        m = re.match(r"opis\s*og[\s_\-]+(.+)", path.stem, re.I)
        if m and m.group(1).strip():
            return m.group(1).strip()
        if path.parent != Path(root):
            return path.parent.name
        return path.stem

    # ------------------------------------------------ konwersja .doc -> .docx

    @staticmethod
    def _stare_docx_dla(path, word, tmpdir):
        """Ścieżka .docx do odczytu: .docx bezpośrednio, .doc przez Word COM."""
        if path.suffix.lower() == ".docx":
            return path, None
        tmp = Path(tmpdir) / (path.stem + ".docx")
        doc = word.Documents.Open(str(path), ReadOnly=True)
        try:
            doc.SaveAs2(str(tmp), FileFormat=16)  # wdFormatXMLDocument
        finally:
            doc.Close(False)
        return tmp, tmp

    # ------------------------------------------------ Word COM / docx -> PDF

    def _stare_zapewnia_word(self):
        """Word COM (.doc -> .docx, docx -> PDF), start raz na całe
        uruchomienie — trzymany w self._stare_word."""
        word = getattr(self, "_stare_word", None)
        if word is None:
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            self._stare_word = word
        return word

    def _stare_docx_na_pdf(self, docx_path, pdf_path):
        """Nowy docx -> PDF przez Worda (ExportAsFixedFormat — ten sam
        mechanizm co konwersja Word->PDF w Pełnym Automacie)."""
        word = self._stare_zapewnia_word()
        doc = word.Documents.Open(str(docx_path), AddToRecentFiles=False)
        try:
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=17,      # wdExportFormatPDF
                OpenAfterExport=False,
                OptimizeFor=0,        # wdExportOptimizeForPrint
                Range=0,              # wdExportAllDocument
                Item=0,               # wdExportDocumentContent
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1,    # wdExportCreateHeadingBookmarks
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
        finally:
            doc.Close(False)

    # ------------------------------------------------ STR_TYT z szablonu

    @staticmethod
    def _stare_zamien_w_akapicie(paragraph, stary, nowy):
        if not stary:
            return
        if stary.lower() in paragraph.text.lower():
            pelny = "".join(r.text for r in paragraph.runs)
            nowy_pelny = re.compile(re.escape(stary), re.I).sub(nowy, pelny)
            if paragraph.runs:
                paragraph.runs[0].text = nowy_pelny
                for i in range(1, len(paragraph.runs)):
                    paragraph.runs[i].text = ""

    @staticmethod
    def _stare_zamien(doc, stary, nowy):
        """Zamiana placeholdera w całym dokumencie (akapity, tabele,
        nagłówki/stopki) — jak Kreator Stron tytułowych."""
        for p in doc.paragraphs:
            TabStareOpisyMixin._stare_zamien_w_akapicie(p, stary, nowy)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        TabStareOpisyMixin._stare_zamien_w_akapicie(p, stary, nowy)
        for s in doc.sections:
            for p in s.header.paragraphs:
                TabStareOpisyMixin._stare_zamien_w_akapicie(p, stary, nowy)
            for p in s.footer.paragraphs:
                TabStareOpisyMixin._stare_zamien_w_akapicie(p, stary, nowy)

    # ------------------------------------------------ przebieg zadania

    def start_stare_opisy_pipeline(self):
        """Zadanie z mapy zadań (web): przycisk 'Przerób na nowy szablon'."""
        self._disable_ui_for_process()

        def _run():
            if pythoncom is not None:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
            try:
                self._stare_opisy_run()
            except Exception:
                self.log("[STARE OPISY BŁĄD] " + traceback.format_exc())
                self.update_status("Błąd", "#D83B01", animate=False)
            finally:
                if pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
                self.restore_all_buttons()

        threading.Thread(target=_run, daemon=True).start()

    # -- jakie pliki jednego folderu wsi rozpoznajemy --
    @staticmethod
    def _stare_pliki_wsi(folder):
        """Mapa {TYP: ścieżka} dla plików starego Pełnego Automatu w folderze."""
        mapa = {}
        for p in sorted(Path(folder).iterdir()):
            if not p.is_file() or p.name.startswith("~$"):
                continue
            if p.suffix.lower() not in (".doc", ".docx"):
                continue
            nazwa = p.stem.upper()
            if nazwa in TYPY_RAPORTOW:
                mapa[nazwa] = p
            elif re.match(r"OPIS\s*OG[\s_\-]", p.name, re.I):
                mapa["OPIS_OG"] = p
            elif nazwa.startswith("STR_TYT"):
                mapa["STR_TYT"] = p
        return mapa

    def _stare_opisy_run(self):
        from app.core.word_worker import get_resource_path
        from app.core import szablony
        from docx import Document

        root_e = getattr(self, "stare_opisy_root_entry", None)
        root_raw = root_e.get().strip() if root_e else ""
        if not root_raw or not Path(root_raw).exists():
            self.log("[STARE OPISY] Wskaż najpierw folder ze starymi plikami "
                     "(foldery wsi albo folder jednej wsi — OPTAX, REJESTR1, "
                     "opis og_* itd.).")
            self.update_status("Brak folderu", "#D83B01", animate=False)
            return
        root = Path(root_raw)

        out_e = getattr(self, "stare_opisy_out_entry", None)
        out_raw = out_e.get().strip() if out_e else ""
        out = Path(out_raw) if out_raw else root
        out.mkdir(parents=True, exist_ok=True)
        self.last_output_dir = out

        # opcjonalny szablon STR_TYT (nowa strona tytułowa usera)
        strtyt_e = getattr(self, "stare_strtyt_tpl_entry", None)
        strtyt_tpl = None
        strtyt_raw = strtyt_e.get().strip() if strtyt_e else ""
        if strtyt_raw:
            if Path(strtyt_raw).exists():
                strtyt_tpl = Path(strtyt_raw)
            else:
                self.log(f"[STARE OPISY] Szablon STR_TYT nie istnieje: "
                         f"{strtyt_raw} — strony tytułowe zostaną pominięte.")

        tpl = get_resource_path(TPL_STARE_FILENAME)
        if not Path(tpl).exists():
            self.log(f"[STARE OPISY BŁĄD] Brak szablonu: {tpl}")
            self.update_status("Brak szablonu", "#D83B01", animate=False)
            return

        # --- odkryj wsie ---
        if self._stare_pliki_wsi(root):
            wsie = [root]
        else:
            wsie = sorted(d for d in root.iterdir()
                          if d.is_dir() and self._stare_pliki_wsi(d))
        if not wsie:
            self.log(f"[STARE OPISY] Nie znaleziono starych plików (OPTAX, "
                     f"REJESTR1, TAB_KLW3, WSK_ZB, ZEST1, HALIZNY, WYK_NEG, "
                     f"opis og_*) w: {root}")
            self.update_status("Brak plików", "#D83B01", animate=False)
            return

        self.log(f"[STARE OPISY] Start — wsi: {len(wsie)} "
                 f"(szablon opisu: {Path(tpl).name}"
                 + (f", szablon STR_TYT: {strtyt_tpl.name}" if strtyt_tpl
                    else "; bez STR_TYT (nie podano szablonu)") + ")")

        self._stare_word = None
        tmpdir = tempfile.mkdtemp(prefix="stare_szablony_")
        done_wsi = 0
        try:
            for i, folder in enumerate(wsie, 1):
                self.check_stop()
                wies = folder.name if folder != root else root.name
                self.update_status(f"Stare szablony: {wies} ({i}/{len(wsie)})")
                if self._stare_przerob_wies(folder, wies, out, tpl, strtyt_tpl,
                                           szablony, Document, tmpdir):
                    done_wsi += 1
        finally:
            word = getattr(self, "_stare_word", None)
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            import shutil as _sh
            _sh.rmtree(tmpdir, ignore_errors=True)

        self.log(f"[STARE OPISY] Zakończono — przerobiono {done_wsi}/{len(wsie)} wsi.")
        self.update_status(f"Gotowe: {done_wsi}/{len(wsie)}", "#107C10", animate=False)

    # ------------------------------------------------ jedna wieś

    def _stare_przerob_wies(self, folder, wies, out, tpl, strtyt_tpl,
                            szablony, Document, tmpdir):
        """Przerabia wszystkie pliki jednej wsi (raporty + opis og + STR_TYT).
        Zwraca True, gdy cokolwiek powstało. Word COM (dla .doc) startuje
        raz na całe uruchomienie — trzymany w self._stare_word."""
        log = self.log
        pliki = self._stare_pliki_wsi(folder)
        # katalog wyników wsi (zachowujemy nazwę podfolderu)
        docelowy = out / wies
        word = getattr(self, "_stare_word", None)

        # --- 1) wczytaj wszystkie stare wordy wsi do .docx ---
        docx_map = {}       # TYP -> (docx_path, tmp_or_None)
        txt_map = {}        # TYP raportu -> ścieżka TXT
        linie_map = {}      # TYP raportu -> linie (unicode, przed cp852)
        for typ, p in sorted(pliki.items()):
            try:
                if p.suffix.lower() == ".doc":
                    if word is None:
                        word = self._stare_zapewnia_word()
                    docx_path, tmp = self._stare_docx_dla(p, word, tmpdir)
                else:
                    docx_path, tmp = p, None
                docx_map[typ] = (docx_path, tmp)
            except Exception as e:
                log(f"[STARE OPISY] {wies}: nie mogę odczytać {p.name} "
                    f"({e}) — pomijam ten plik.")
                continue

        if not docx_map:
            log(f"[STARE OPISY] {wies}: żaden plik nie dał się odczytać.")
            return False

        cos_powstalo = False

        # --- 2) raporty TXT -> nowe szablony (HTML + PDF) ---
        for typ, (docx_path, _tmp) in sorted(docx_map.items()):
            if typ not in TYPY_RAPORTOW:
                continue
            try:
                doc = Document(str(docx_path))
                linie = self._stare_linie(doc)
            except Exception as e:
                log(f"[STARE OPISY] {wies}: {typ}: błąd odczytu ({e}) — pomijam.")
                continue
            if sum(len(l.strip()) for l in linie) < 40:
                log(f"[STARE OPISY] {wies}: {typ} jest pusty — pomijam.")
                continue
            txt_path = Path(tmpdir) / f"{wies}_{typ}.TXT"
            self._stare_na_txt(linie, txt_path)
            txt_map[typ] = txt_path
            linie_map[typ] = linie

            # sanity: parser musi coś z tego wyciągnąć
            try:
                if typ == "OPTAX":
                    n = len(szablony.parse_optax(txt_path))
                elif typ == "REJESTR1":
                    n = len(szablony.parse_rejestr1(txt_path))
                elif typ == "TAB_KLW3":
                    n = len(szablony.parse_tabklw3(txt_path))
                elif typ == "WSKAZ1":
                    n = len(szablony.parse_wskaz1(txt_path))
                elif typ == "WSK_ZB":
                    n = len(szablony.parse_wskzb(txt_path))
                elif typ == "ZEST1":
                    n = len(szablony.parse_zest1(txt_path))
                elif typ == "HALIZNY":
                    n = len(szablony.parse_halizny(txt_path))
                else:  # WYK_NEG
                    n = len(szablony.parse_wyk_neg(txt_path)[1])
                if n == 0:
                    log(f"[STARE OPISY] {wies}: {typ} — parser nie znalazł "
                        f"danych (pusty/zmieniony układ) — pomijam.")
                    continue
            except Exception as e:
                log(f"[STARE OPISY] {wies}: {typ}: błąd parsowania ({e}) — "
                    f"próbuję mimo to.")

            html_dir = docelowy / "nowe szablony"
            pdf_dir = docelowy / "nowe szablony" / "pdf"
            pdf_dir.mkdir(parents=True, exist_ok=True)
            try:
                szablony.generuj_raport_pdf(
                    typ, txt_path, pdf_dir / f"{typ}.pdf",
                    html_out=html_dir / f"{typ}.html")
                cos_powstalo = True
                log(f"[STARE OPISY] {wies}: {typ} → nowe szablony/{typ}.html "
                    f"+ nowe szablony/pdf/{typ}.pdf")
            except Exception as e:
                log(f"[STARE OPISY] {wies}: {typ}: BŁĄD generowania ({e}).")

        # --- 3) opis ogólny -> nowy szablon docx ---
        if "OPIS_OG" in docx_map:
            docx_path, _tmp = docx_map["OPIS_OG"]
            try:
                doc = Document(str(docx_path))
                from app.gui.tabs.tab_opis_og import TabOpisOgMixin
                _l = TabOpisOgMixin._opis_og_fmt

                etaty = self._stare_etat_z_tabeli(doc)
                przed = self._stare_przedrzedne(doc)
                if etaty and "ETAT_PRZYJETY" in etaty:
                    vals = {
                        "ETAT_POTRZEBY": _l(etaty.get("ETAT_POTRZEBY",
                                                        etaty["ETAT_PRZYJETY"])),
                        "ETAT_PRZYJETY": _l(etaty["ETAT_PRZYJETY"]),
                        "POZOSTALE_REBNE": _l(etaty.get("POZOSTALE_REBNE", 0)),
                    }
                    if "MAKS_MIAZSZOSC" in etaty:
                        vals["MAKS_MIAZSZOSC"] = _l(etaty["MAKS_MIAZSZOSC"])
                    else:
                        vals["MAKS_MIAZSZOSC"] = _l(
                            etaty["ETAT_PRZYJETY"] + etaty.get("POZOSTALE_REBNE", 0))
                    if przed is not None:
                        vals["UZYTK_PRZEDRZEBNE"] = _l(przed)
                    formy, lasy_niezero = self._stare_formy(doc)

                    docelowy.mkdir(parents=True, exist_ok=True)
                    out_path = docelowy / f"opis og_{wies}.docx"
                    TabOpisOgMixin._fill_template(tpl, vals, out_path, formy=formy)
                    cos_powstalo = True
                    # opis og na nowym szablonie -> też PDF do folderu pdf
                    try:
                        pdf_dir = docelowy / "nowe szablony" / "pdf"
                        pdf_dir.mkdir(parents=True, exist_ok=True)
                        self._stare_docx_na_pdf(out_path, pdf_dir / f"opis og_{wies}.pdf")
                        log(f"[STARE OPISY] {wies}: opis og → PDF "
                            f"(nowe szablony/pdf/opis og_{wies}.pdf)")
                    except Exception as e:
                        log(f"[STARE OPISY] {wies}: opis og — nie udało się "
                            f"wygenerować PDF ({e}); docx został zapisany.")
                    opis_form = (f"{len(formy) - 1} form ochrony"
                                 if formy and len(formy) > 1 else
                                 ("formy: brak" if formy is None
                                  else "formy: same nagłówki"))
                    log(f"[STARE OPISY] {wies}: zapisano opis og_{wies}.docx "
                        f"(rębne {vals['MAKS_MIAZSZOSC']} m3, przedrębne "
                        f"{vals.get('UZYTK_PRZEDRZEBNE', '?')} m3, {opis_form})"
                        + ("\n[STARE OPISY] ⚠ UWAGA: lasy ochronne mają NIEZEROWĄ "
                           "powierzchnię — nowy szablon ma statyczną linię "
                           "„Lasy ochronne – 0,00ha”; popraw ją ręcznie."
                           if lasy_niezero else ""))
                else:
                    log(f"[STARE OPISY] {wies}: nie znaleziono tabeli etatów "
                        f"w 'opis og' — pomijam opis ogólny tej wsi.")
            except Exception as e:
                log(f"[STARE OPISY] {wies}: opis og — błąd generowania ({e}).")

        # --- 4) STR_TYT z nowego szablonu usera (opcjonalnie) ---
        if strtyt_tpl is not None and "STR_TYT" in pliki:
            # nazwa wsi: najpierw "Obiekt:" z raportów, w ostateczności folder
            wies_doc = wies
            for typ in ("OPTAX", "REJESTR1", "WSK_ZB"):
                txt = txt_map.get(typ)
                if not txt:
                    continue
                obiekt, _stan, _okres = szablony.meta_z_pliku(txt)
                if obiekt:
                    wies_doc = obiekt
                    break
            # powierzchnia "Razem" z OPTAX
            pow_txt = ""
            if "OPTAX" in txt_map:
                try:
                    pow_txt = szablony.razem_z_optax(txt_map["OPTAX"])
                except Exception:
                    pow_txt = ""
            if not pow_txt or pow_txt == "[BRAK_DANYCH]":
                pow_txt = "[BRAK_DANYCH]"
            try:
                doc = Document(str(strtyt_tpl))
                self._stare_zamien(doc, "NAZWA WSI", wies_doc)
                self._stare_zamien(doc, "wielkość", pow_txt)
                safe = "".join(c for c in wies_doc
                               if c.isalpha() or c.isdigit() or c in " -_").strip()
                docelowy.mkdir(parents=True, exist_ok=True)
                out_path = docelowy / f"STR_TYT_{safe or wies}.docx"
                doc.save(str(out_path))
                cos_powstalo = True
                log(f"[STARE OPISY] {wies}: STR_TYT_{safe or wies}.docx "
                    f"(Pow: {pow_txt})")
                # strona tytułowa -> też PDF do folderu pdf
                try:
                    pdf_dir = docelowy / "nowe szablony" / "pdf"
                    pdf_dir.mkdir(parents=True, exist_ok=True)
                    self._stare_docx_na_pdf(out_path,
                                             pdf_dir / f"STR_TYT_{safe or wies}.pdf")
                    log(f"[STARE OPISY] {wies}: STR_TYT → PDF "
                        f"(nowe szablony/pdf/STR_TYT_{safe or wies}.pdf)")
                except Exception as e:
                    log(f"[STARE OPISY] {wies}: STR_TYT — nie udało się "
                        f"wygenerować PDF ({e}); docx został zapisany.")
            except Exception as e:
                log(f"[STARE OPISY] {wies}: STR_TYT — błąd generowania ({e}).")

        # --- 4b) bez szablonu STR_TYT: stara strona tytułowa jako PDF ---
        elif "STR_TYT" in docx_map:
            try:
                pdf_dir = docelowy / "nowe szablony" / "pdf"
                pdf_dir.mkdir(parents=True, exist_ok=True)
                stary = docx_map["STR_TYT"][0]
                self._stare_docx_na_pdf(stary, pdf_dir / f"{stary.stem}.pdf")
                cos_powstalo = True
                log(f"[STARE OPISY] {wies}: stara STR_TYT skopiowana do PDF "
                    f"(nowe szablony/pdf/{stary.stem}.pdf)")
            except Exception as e:
                log(f"[STARE OPISY] {wies}: stara STR_TYT — nie udało się "
                    f"skopiować do PDF ({e}).")

        if not cos_powstalo:
            log(f"[STARE OPISY] {wies}: nic nie powstało dla tej wsi.")
        return cos_powstalo
