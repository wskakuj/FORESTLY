"""
Forestly — Mixin: TabPdfMixin
"""

import customtkinter as ctk
import time
import pymupdf as fitz
from pypdf import PdfWriter
from pypdf import PdfReader
import win32com.client

from app.config import (
    get_saved_excluded_templates,
    PDF_ORDER_TEMPLATES, build_ordered_pdfs_from_templates, get_saved_template_order, is_file_locked, template_matches,
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
        cb.grid(row=row_idx, column=0, columnspan=3, padx=15, pady=(0, 20), sticky="w")

    def task_convert_to_pdf(self, in_dir, out_dir):
        docs = [
            p
            for p in in_dir.rglob("*")
            if p.is_file() and p.suffix.lower() in {".doc", ".docx"}
        ]
        if not docs:
            return 0

        total_docs = len(docs)
        self.start_progress_tracking(total_docs, "Konwersja Word -> PDF")

        # Inicjalizacja strumienia
        self.init_live_stream(total_docs)
        for doc_path in docs:
            rel_path = doc_path.relative_to(in_dir)
            target = out_dir / rel_path.parent / f"{doc_path.stem}.pdf"
            self.add_to_stream_queue(doc_path, target)

        word = None
        count = 0
        try:
            word = win32com.client.DispatchEx("Word.Application")
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
                word.Quit()
        return count

    def task_merge_pdfs(self, in_dir, out_dir, mode_key="ALL"):
        # INTELIGENTNY WYBÓR TRYBU: Jedna wieś vs Wiele wsi
        direct_pdfs = list(in_dir.glob("*.pdf"))

        if direct_pdfs:
            # TRYB JEDNEJ WSI (pliki leżą bezpośrednio w głównym folderze)
            pdf_dirs = {in_dir}
        else:
            # TRYB WIELU WSI (szukamy plików tylko w podfolderach)
            pdf_dirs = set(p.parent for p in in_dir.rglob("*.pdf"))

        if not pdf_dirs:
            return 0

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
        self.start_progress_tracking(total_dirs, "Scalanie PDF")
        template_keys = get_saved_template_order(in_dir, mode_key)
        excluded_keys = get_saved_excluded_templates(in_dir, mode_key)
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
            target = target_dir / f"{village_name}_scalony.pdf"

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
        return count

    def task_remove_blank_pages(self, in_dir, out_dir):
        pdfs = list(in_dir.rglob("*.pdf"))
        if not pdfs:
            return 0

        count = 0
        total_pdfs = len(pdfs)
        self.start_progress_tracking(total_pdfs, "Usuwanie pustych stron")

        for idx_pdf, pdf_path in enumerate(pdfs, start=1):
            self.check_stop()
            self.set_progress((idx_pdf - 1) / total_pdfs if total_pdfs else 1, current_file=pdf_path.name,
                              current=idx_pdf - 1)

            target = out_dir / pdf_path.relative_to(in_dir)
            target.parent.mkdir(parents=True, exist_ok=True)

            doc = fitz.open(str(pdf_path))
            out = fitz.open()
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
                    out.insert_pdf(doc, from_page=i, to_page=i)

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

