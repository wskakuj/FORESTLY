"""
Forestly — Mixin: TabPdfMixin
"""

import customtkinter as ctk
import re
import time
from pathlib import Path

import pymupdf as fitz
from pypdf import PdfWriter, PdfReader
import win32com.client

from app.config import (
    get_saved_excluded_templates,
    PDF_ORDER_TEMPLATES, build_ordered_pdfs_from_templates, get_saved_template_order, is_file_locked, template_matches,
    load_order_store, set_saved_excluded_templates, set_saved_template_order,
)

class TabPdfMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def _setup_pdf_extras(self, card_frame, row_idx):
        self.pdf_merge_var = ctk.BooleanVar(value=True)
        cb = ctk.CTkCheckBox(
            card_frame,
            text="Po konwersji scal pliki w jeden dokument PDF",
            variable=self.pdf_merge_var,
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
        cb.grid(row=row_idx, column=0, columnspan=3, padx=15, pady=(0, 5), sticky="w")

        self.pdf_skroty_var = ctk.BooleanVar(value=True)
        cb_skroty = ctk.CTkCheckBox(
            card_frame,
            text="Dołącz 'Skróty i symbole' na końcu scalonego PDF",
            variable=self.pdf_skroty_var,
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
        cb_skroty.grid(row=row_idx + 1, column=0, columnspan=3, padx=15, pady=(0, 20), sticky="w")

    def task_convert_to_pdf(self, in_dir, out_dir):
        docs = [
            p
            for p in in_dir.rglob("*")
            if p.is_file() and p.suffix.lower() in {".doc", ".docx"}
        ]
        if not docs:
            return 0

        total_docs = len(docs)
        self.last_output_dir = Path(out_dir)
        self.start_progress_tracking(total_docs, "Konwersja Word -> PDF")

        # Inicjalizacja strumienia
        self.init_live_stream(total_docs)
        for doc_path in docs:
            rel_path = doc_path.relative_to(in_dir)
            target = out_dir / rel_path.parent / f"{doc_path.stem}.pdf"
            self.add_to_stream_queue(doc_path, target)

        word = None
        word_pid = None
        count = 0
        try:
            word = win32com.client.DispatchEx("Word.Application")
            try:
                from app.core import office_guard
                word_pid = office_guard.register(word)
            except Exception:
                pass
            word.Visible, word.DisplayAlerts = False, 0
            # --- OPTYMALIZACJA PRĘDKOŚCI ---
            word.Application.ScreenUpdating = False
            word.Options.BackgroundSave = False
            word.Options.CheckSpellingAsYouType = False
            word.Options.CheckGrammarAsYouType = False
            word.Options.UpdateFieldsAtPrint = False
            # --------------------------------
            for doc_path in docs:
                self.check_stop()
                if is_file_locked(doc_path):
                    self.log(f"Zablokowany: {doc_path.name}")
                    continue
                rel_path = doc_path.relative_to(in_dir)
                target = out_dir / rel_path.parent / f"{doc_path.stem}.pdf"
                target.parent.mkdir(parents=True, exist_ok=True)
                doc = None
                try:
                    self.set_progress(count / total_docs if total_docs else 1, current_file=doc_path.name,
                                      current=count)
                    self.start_stream_file(doc_path, target)
                    start_time = time.time()

                    doc = word.Documents.Open(str(doc_path), AddToRecentFiles=False)
                    # ExportAsFixedFormat jest znacznie szybszy niż SaveAs(FileFormat=17)
                    # nie wymaga ręcznego Repaginate() — Word robi to automatycznie
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(target),
                        ExportFormat=17,  # wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,     # wdExportOptimizeForPrint
                        Range=0,           # wdExportAllDocument
                        Item=0,            # wdExportDocumentContent
                        IncludeDocProps=True,
                        KeepIRM=True,
                        CreateBookmarks=1, # wdExportCreateHeadingBookmarks
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False,
                    )

                    duration = time.time() - start_time
                    self.complete_stream_file(doc_path, target, duration)
                    count += 1
                    self.set_progress(count / total_docs if total_docs else 1, current_file=doc_path.name, current=count)
                except Exception as e:
                    self.log(f"Problem konwersji obiektu {doc_path.name}: {e}")
                finally:
                    if doc is not None:
                        doc.Close(SaveChanges=False)
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
                try:
                    from app.core import office_guard
                    office_guard.unregister(word_pid)
                except Exception:
                    pass
        return count

    def task_merge_pdfs(self, in_dir, out_dir, mode_key="ALL"):
        # [NUMERACJA] po poprzednim biegu pliki mają prefiksy pozycji
        # ("04_OPTAX.pdf"); jeśli świeży przebieg wygenerował już czystą
        # nazwę ("OPTAX.pdf"), stary duplikat usuwamy, żeby nie wchodził
        # do scalanki podwójnie
        for f in in_dir.rglob("[0-9][0-9]_*.pdf"):
            base = f.with_name(f.name[3:])
            if base.exists():
                f.unlink()
                self.log(f"  [NUMERACJA] Usunięto duplikat z poprzedniego "
                         f"przebiegu: {f.name}")

        # INTELIGENTNY WYBÓR TRYBU: Jedna wieś vs Wiele wsi
        direct_pdfs = list(in_dir.glob("*.pdf"))
        pdf_dirs = set()

        # Podfoldery (wsie) — KAŻDY folder z PDF-ami dostaje własny scalony plik,
        # zawsze, niezależnie od tego, co leży w folderze głównym.
        for _p in in_dir.rglob("*.pdf"):
            if _p.parent != in_dir:
                pdf_dirs.add(_p.parent)

        # Pliki leżące bezpośrednio w folderze głównym — scalane jako jeden
        # pakiet. Wcześniej jeden luźny PDF wyłączał scalanie wszystkich wsi
        # (tryb "jednej wsi" łapał tylko luźne pliki i pomijał podfoldery).
        if direct_pdfs:
            pdf_dirs.add(in_dir)

        if not pdf_dirs:
            return 0

        if direct_pdfs and len(pdf_dirs) > 1:
            self.log(
                f"[SCALANIE] {len(direct_pdfs)} plik(ów) PDF bezpośrednio w folderze "
                f"głównym zostanie scalonych jako jeden pakiet, a każda wieś "
                f"({len(pdf_dirs) - 1}) osobno."
            )

        # --- KONTROLA KOMPLETNOŚCI ---
        warnings = []
        for folder in pdf_dirs:
            if folder == in_dir:
                # Bierzemy tylko pliki z głównego folderu
                pdfs = [p.name.lower() for p in in_dir.glob("*.pdf")]
                village_name = in_dir.parent.name
                if village_name.upper() in ["PDF", "WORD", "TXT"]:
                    village_name = in_dir.parent.parent.name
            else:
                pdfs = [p.name.lower() for p in folder.iterdir() if p.suffix.lower() == ".pdf"]
                village_name = folder.name

            _tpl = {t["key"]: t for t in PDF_ORDER_TEMPLATES}
            has_title = any(template_matches(_tpl["TITLE"], p) for p in pdfs)
            has_optax = any(template_matches(_tpl["OPTAX"], p) for p in pdfs)
            has_opis = any(template_matches(_tpl["OPIS"], p) for p in pdfs)
            has_rej = any(template_matches(_tpl["REJESTR1"], p) for p in pdfs)

            missing = []
            if not has_title: missing.append("STR_TYT")
            if not (has_optax or has_opis): missing.append("OPTAX / OPIS")
            if not has_rej: missing.append("REJESTR")

            if missing:
                warnings.append(f"• Wieś {village_name.upper()}: brak -> {', '.join(missing)}")

        if warnings:
            self.log("[KONTROLA] Wykryto braki w folderach do scalenia. Oczekiwanie na decyzję...")
            if not self.show_validation_window_sync("Wykryto brakujące pliki (niektóre wsie nie są kompletne):",
                                                    warnings):
                raise InterruptedError("Operacja scalania przerwana przez użytkownika.")
        # -----------------------------

        count = 0
        total_dirs = len(pdf_dirs)
        self.last_output_dir = Path(out_dir)
        self.start_progress_tracking(total_dirs, "Scalanie PDF")
        # [UKŁAD] gdy ten folder nie ma własnego układu (np. nowy folder wyników),
        # użyj globalnie zapamiętanej kolejności z ustawień programu
        _store = load_order_store(Path(in_dir))
        if not (isinstance(_store.get(mode_key), list) and _store.get(mode_key)):
            _glob = None
            _gs = getattr(self, "get_setting", None)
            if _gs is not None:
                try:
                    _glob = _gs(f"pdf_order.{mode_key}", None)
                except Exception:
                    _glob = None
            if isinstance(_glob, dict) and isinstance(_glob.get("order"), list) \
                    and _glob["order"]:
                set_saved_template_order(Path(in_dir), mode_key, _glob["order"])
                if isinstance(_glob.get("excluded"), list):
                    set_saved_excluded_templates(Path(in_dir), mode_key,
                                                 _glob["excluded"])
                self.log("[UKŁAD] Użyto zapamiętanej kolejności PDF z ustawień.")
        template_keys = get_saved_template_order(in_dir, mode_key)
        excluded_keys = get_saved_excluded_templates(in_dir, mode_key)
        # jawny zapis użytej kolejności — od razu widać w logach, czyje
        # ustawienie zadziałało (i w jakiej kolejności scalono)
        _lbl_map = {t["key"]: t["label"] for t in PDF_ORDER_TEMPLATES}
        self.log("[UKŁAD] Kolejność scalania (" + mode_key + "): "
                 + " → ".join(_lbl_map.get(k, k) for k in template_keys
                              if k not in excluded_keys))
        # od v2.0.32: 'Opis ogólny' to stała część zestawienia (generuje go
        # Pełny Automat) — nie scalamy bez niego
        if "OPIS" in excluded_keys:
            excluded_keys = [k for k in excluded_keys if k != "OPIS"]
            set_saved_excluded_templates(in_dir, mode_key, excluded_keys)
            set_saved_template_order(in_dir, mode_key, template_keys)
            self.log("[UKŁAD] 'Opis ogólny' był wykluczony — przywrócono go "
                     "do scalania (zaraz za stroną tytułową).")
        if excluded_keys:
            _lbl = ", ".join(
                t["label"] for t in PDF_ORDER_TEMPLATES if t["key"] in excluded_keys)
            self.log(f"[UKŁAD] Wykluczono z scalania: {_lbl}")

        for idx_dir, folder in enumerate(pdf_dirs, start=1):
            self.check_stop()

            if folder == in_dir:
                village_name = in_dir.parent.name
                if village_name.upper() in ["PDF", "WORD", "TXT"]:
                    village_name = in_dir.parent.parent.name
                target_dir = out_dir
                pdfs = sorted(list(in_dir.glob("*.pdf")))
            else:
                village_name = folder.name
                target_dir = out_dir / folder.relative_to(in_dir)
                pdfs = sorted([p for p in folder.iterdir() if p.suffix.lower() == ".pdf"])

            self.set_progress((idx_dir - 1) / total_dirs if total_dirs else 1, current_file=village_name,
                              current=idx_dir - 1)

            ordered_pdfs = build_ordered_pdfs_from_templates(pdfs, template_keys, excluded_keys)
            if not ordered_pdfs:
                continue

            target_dir.mkdir(parents=True, exist_ok=True)
            # nazwa pliku: UPUL/ISL wg wyboru z kreatora strony tytułowej
            target = target_dir / f"{village_name}_{self._nazwa_dokumentu()}.pdf"

            writer = PdfWriter()
            current_page = 0
            try:
                for pdf in ordered_pdfs:
                    friendly_name = pdf.stem
                    for tpl in PDF_ORDER_TEMPLATES:
                        if template_matches(tpl, pdf.name):
                            friendly_name = tpl["label"]
                            break

                    reader = PdfReader(str(pdf))
                    num_pages = len(reader.pages)

                    for page in reader.pages:
                        writer.add_page(page)

                    writer.add_outline_item(friendly_name, current_page)
                    current_page += num_pages

                writer.add_metadata({
                    "/Title": f"UPUL - {village_name.upper()}",
                    "/Author": "Agencja Cezar",
                    "/Creator": "Forestly",
                    "/Producer": "Forestly"
                })

                with open(target, "wb") as f_out:
                    writer.write(f_out)
                self.log(f"Połączono: {target.name}")
                count += 1
                self.set_progress(idx_dir / total_dirs if total_dirs else 1, current_file=village_name, current=idx_dir)
            except Exception as e:
                self.log(f"Błąd przy {target.name}: {e}")
            finally:
                writer.close()

        # [NUMERACJA] pliki w folderze PDF dostają prefiks z pozycją
        # w ustawionym układzie scalania (np. "03_OPTAX.pdf")
        try:
            self._numeruj_pdfy_wg_ukladu(pdf_dirs, in_dir, template_keys)
        except Exception as e:
            self.log(f"[NUMERACJA] Nie udało się dodać prefiksów: {e}")
        return count

    def _nazwa_dokumentu(self):
        """UPUL albo ISL — wg 'Typ dokumentu' z kreatora strony tytułowej.

        Używane w nazwie scalonego PDF (np. 'CHORZEWO_UPUL.pdf') zamiast
        dawnego przyrostka '_scalony'.
        """
        v = getattr(self, "all_tpl_doc_var", None)
        try:
            t = str(v.get()).strip().upper()
        except Exception:
            t = ""
        return t if t in ("UPUL", "ISL") else "UPUL"

    def _numeruj_pdfy_wg_ukladu(self, pdf_dirs, in_dir, template_keys):
        """Prefiks numeru pozycji z ustawionego układu PDF w nazwach plików.

        "OPTAX.pdf" -> "03_OPTAX.pdf", gdy OPTAX jest 3. pozycją układu.
        Pliki niedopasowane do żadnego szablonu zostają bez zmian;
        powtórne uruchomienie nie podwaja prefiksu (stary "NN_" zdejmujemy).
        """
        pozycje = {k: i + 1 for i, k in enumerate(template_keys)}
        wzorce = [t for t in PDF_ORDER_TEMPLATES if t["key"] in pozycje]
        if not wzorce:
            return
        for folder in pdf_dirs:
            pliki = (in_dir.glob("*.pdf") if folder == in_dir
                     else (p2 for p2 in folder.iterdir()
                           if p2.suffix.lower() == ".pdf"))
            for f in sorted(pliki):
                if not f.is_file():
                    continue
                klucz = None
                for t in wzorce:
                    if template_matches(t, f.name):
                        klucz = t["key"]
                        break
                if klucz is None:
                    continue
                czysta = re.sub(r"^\d\d_", "", f.name)
                docel = f.with_name(f"{pozycje[klucz]:02d}_{czysta}")
                if docel != f and not docel.exists():
                    f.rename(docel)
                    self.log(f"  [NUMERACJA] {f.name} → {docel.name}")

    def task_remove_blank_pages(self, in_dir, out_dir):
        pdfs = list(in_dir.rglob("*.pdf"))
        if not pdfs:
            return 0

        count = 0
        total_pdfs = len(pdfs)
        self.last_output_dir = Path(out_dir)
        self.start_progress_tracking(total_pdfs, "Usuwanie pustych stron")

        for idx_pdf, pdf_path in enumerate(pdfs, start=1):
            self.check_stop()
            self.set_progress((idx_pdf - 1) / total_pdfs if total_pdfs else 1, current_file=pdf_path.name,
                              current=idx_pdf - 1)

            target = out_dir / pdf_path.relative_to(in_dir)
            target.parent.mkdir(parents=True, exist_ok=True)

            doc = fitz.open(str(pdf_path))
            out = fitz.open()
            # mapa: stary indeks strony -> nowy indeks (żeby przenieść spis treści)
            mapa_stron = {}
            for i in range(doc.page_count):
                page = doc.load_page(i)
                pix = page.get_pixmap(
                    matrix=fitz.Matrix(100 / 72, 100 / 72),
                    colorspace=fitz.csGRAY,
                    alpha=False,
                )
                data = pix.samples
                white = sum(1 for v in data if v >= 250)
                if (white / len(data)) < 0.995:
                    mapa_stron[i] = out.page_count
                    out.insert_pdf(doc, from_page=i, to_page=i)

            # spis treści (zakładki) — przeniesiony z pominięciem usuniętych stron
            try:
                stary_toc = doc.get_toc(simple=True)
                if stary_toc:
                    nowy_toc = []
                    zachowane = sorted(mapa_stron)
                    for lvl, tytul, strona in stary_toc:
                        idx = strona - 1
                        if idx in mapa_stron:
                            nowy_toc.append([lvl, tytul, mapa_stron[idx] + 1])
                        else:
                            # strona docelowa usunięta jako pusta -> najbliższa zachowana
                            nastepna = next(
                                (mapa_stron[j] for j in zachowane if j > idx), None)
                            if nastepna is not None:
                                nowy_toc.append([lvl, tytul, nastepna + 1])
                    if nowy_toc:
                        out.set_toc(nowy_toc)
            except Exception:
                pass

            # Inteligentne wyciąganie nazwy wsi do metadanych
            if pdf_path.parent == in_dir:
                village_name = in_dir.parent.name.upper()
                if village_name in ["PDF POLACZONE", "PDF", "WORD", "TXT"]:
                    village_name = in_dir.parent.parent.name.upper()
            else:
                village_name = pdf_path.parent.name.upper()

            out.set_metadata({
                "title": f"UPUL - {village_name}",
                "author": "Agencja Cezar",
                "creator": "Forestly",
                "producer": "Forestly"
            })

            out.save(str(target))
            out.close()
            doc.close()
            count += 1
            self.set_progress(idx_pdf / total_pdfs if total_pdfs else 1, current_file=pdf_path.name, current=idx_pdf)

        return count

