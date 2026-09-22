"""
Kombajn Leśny PRO — Mixin: TabAllMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import os
import re
import traceback
import shutil
import tempfile
import numpy as np
from docx import Document
import win32com.client
import pythoncom

from app.config import (
    is_file_locked, load_margins, save_margins,
)

from app.core.word_worker import (
    get_resource_path,
)

class TabAllMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def _setup_all_extras(self, card_frame, row_idx):
        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")

        # 1. STR_TYT Checkbox
        self.all_gen_str_tyt_var = ctk.BooleanVar(value=False)
        cb_str = ctk.CTkCheckBox(
            card_frame,
            text="Generuj strony tytułowe (STR_TYT) na podstawie OPTAX",
            variable=self.all_gen_str_tyt_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            command=self._toggle_all_template_ui,
        )
        cb_str.grid(
            row=row_idx, column=0, columnspan=3, padx=15, pady=(0, 5), sticky="w"
        )

        self.all_template_frame = ctk.CTkFrame(
            card_frame, fg_color="#1E1E1E", border_width=1, border_color="#333333"
        )
        self.all_template_frame.grid(
            row=row_idx + 1, column=0, columnspan=3, padx=15, pady=(0, 10), sticky="ew"
        )
        self.all_template_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self.all_template_frame,
            text="Plik szablonu (.docx):",
            font=font_label,
            text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=(10, 10), pady=5, sticky="w")
        self.all_template_entry = ctk.CTkEntry(
            self.all_template_frame,
            placeholder_text="Wskaż plik bazowy STR_TYT...",
            height=32,
        )
        self.all_template_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(
            self.all_template_frame,
            text="Wybierz",
            command=lambda: self.select_file(
                self.all_template_entry, [("Word", "*.docx")]
            ),
            width=90,
            height=32,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=(5, 10), pady=5)

        # 2. SKROTY - Zawsze włączone, opcjonalnie z własnego pliku
        self.all_custom_skroty_var = ctk.BooleanVar(value=False)
        cb_skroty = ctk.CTkCheckBox(
            card_frame,
            text="Użyj własnego pliku 'Skróty i symbole' (zamiast domyślnego z programu)",
            variable=self.all_custom_skroty_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            command=self._toggle_all_skroty_ui,
        )
        cb_skroty.grid(
            row=row_idx + 2, column=0, columnspan=3, padx=15, pady=(5, 5), sticky="w"
        )

        self.all_skroty_frame = ctk.CTkFrame(
            card_frame, fg_color="#1E1E1E", border_width=1, border_color="#333333"
        )
        self.all_skroty_frame.grid(
            row=row_idx + 3, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="ew"
        )
        self.all_skroty_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self.all_skroty_frame,
            text="Własny plik:",
            font=font_label,
            text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=(10, 10), pady=8, sticky="w")

        self.all_skroty_entry = ctk.CTkEntry(
            self.all_skroty_frame,
            placeholder_text="Wskaż własny plik ze skrótami...",
            height=32,
        )
        self.all_skroty_entry.grid(row=0, column=1, padx=5, pady=8, sticky="ew")

        ctk.CTkButton(
            self.all_skroty_frame,
            text="Wybierz",
            command=lambda: self.select_file(
                self.all_skroty_entry,
                [("Word/PDF", "*.docx *.doc *.pdf"), ("Wszystkie pliki", "*.*")],
            ),
            width=90,
            height=32,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=(5, 10), pady=8)

        self._toggle_all_template_ui()
        self._toggle_all_skroty_ui()

        # --- DODANA TABELA MARGINESÓW ---
        self._build_margins_ui(card_frame, row_idx + 4, "ALL")

    def _toggle_all_template_ui(self):
        state = (
            "normal"
            if getattr(self, "all_gen_str_tyt_var", None)
               and self.all_gen_str_tyt_var.get()
            else "disabled"
        )
        if hasattr(self, "all_template_frame"):
            for child in self.all_template_frame.winfo_children():
                try:
                    child.configure(state=state)
                except:
                    pass
                if hasattr(child, "winfo_children"):
                    for subchild in child.winfo_children():
                        try:
                            subchild.configure(state=state)
                        except:
                            pass

    def _toggle_all_skroty_ui(self):
        if getattr(self, "all_custom_skroty_var", None) and getattr(self, "all_skroty_frame", None):
            if self.all_custom_skroty_var.get():
                # Jeśli checkbox jest zaznaczony, przywracamy ramkę z powrotem na ekran
                self.all_skroty_frame.grid()
            else:
                # Jeśli checkbox jest odznaczony, całkowicie ukrywamy ramkę
                self.all_skroty_frame.grid_remove()

    def _resolve_skroty_path(self):
        """Zwraca ścieżkę do pliku 'Skróty i symbole' (własny użytkownika lub domyślny z zasobów programu)."""
        # Sprawdzamy, czy użytkownik chce użyć własnego pliku
        if getattr(self, "all_custom_skroty_var", None) and self.all_custom_skroty_var.get():
            custom = self.all_skroty_entry.get().strip()
            if custom and Path(custom).exists():
                return custom
            self.log(
                "[UWAGA] Wskazano własny plik 'Skróty i symbole', ale nie istnieje. Używam domyślnego z programu."
            )
        # Pobieranie domyślnego pliku z zasobów programu w tle
        domyslne = get_resource_path("Skroty.pdf")
        if not domyslne.exists():
            domyslne = get_resource_path("Skroty.docx")
        if domyslne.exists():
            return str(domyslne)
        return None

    def start_pipeline(self, mode):
        src_path = self.entries[mode]["src"].get()
        dst_path = self.entries[mode]["dst"].get()
        remove_names_flag = self.remove_names_var.get()
        if not src_path or not os.path.exists(src_path):
            messagebox.showwarning(
                "Nieprawidłowa ścieżka", "Wybierz istniejący folder źródłowy."
            )
            return
        if not dst_path:
            messagebox.showwarning("Nieprawidłowa ścieżka", "Wybierz folder docelowy.")
            return
        if self.running:
            return
        self.last_output_dir = Path(dst_path)
        self._disable_ui_for_process()
        self.log(f"[{mode}] URUCHOMIENIE ZADANIA\nZ: {src_path}\nDo: {dst_path}")
        self.set_progress(0)

        # Pobieranie i zapis marginesów
        margins_dict = {}
        if mode in ["ALL", "WORD"] and hasattr(self, "margin_vars") and mode in self.margin_vars:
            saved_config = load_margins()  # Odczytujemy stary plik, żeby nie nadpisać układu drugiej zakładki
            if mode not in saved_config:
                saved_config[mode] = {}

            for ftype, entries in self.margin_vars[mode].items():
                try:
                    t = float(entries["T"].get().replace(',', '.'))
                    b = float(entries["B"].get().replace(',', '.'))
                    l = float(entries["L"].get().replace(',', '.'))
                    r = float(entries["R"].get().replace(',', '.'))
                    margins_dict[ftype] = [t, b, l, r]

                    # Wrzucamy do struktury do zapisania na dysku
                    saved_config[mode][ftype] = {"T": t, "B": b, "L": l, "R": r}
                except ValueError:
                    messagebox.showwarning("Błąd", f"Marginesy dla {ftype} muszą być liczbami (np. 1.5).")
                    self.restore_all_buttons()
                    return

            # Zapis całego słownika z marginesami na dysk
            save_margins(saved_config)

        threading.Thread(
            target=self.run_logic_thread,
            args=(src_path, dst_path, mode, remove_names_flag, margins_dict),
            daemon=True,
        ).start()

    def task_generate_str_tyt(self, word_dir, template_path, village_ph, area_ph):
        word_dir = Path(word_dir)
        optax_files = sorted(
            [
                p
                for p in word_dir.rglob("OPTAX*.doc*")
                if p.is_file() and not p.name.startswith("~$")
            ]
        )
        if not optax_files:
            self.log(
                "[STR_TYT] Nie znaleziono plików OPTAX w folderze Word. Pomijam generowanie."
            )
            return

        self.log(
            f"[STR_TYT] Rozpoczynam generowanie stron tytułowych dla {len(optax_files)} wsi..."
        )
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0

        try:
            for optax_path in optax_files:
                self.check_stop()
                if is_file_locked(optax_path):
                    self.log(f"  [Pominięto] Plik zablokowany: {optax_path.name}")
                    continue

                try:
                    doc_word = word_app.Documents.Open(str(optax_path), ReadOnly=True)
                    text_content = doc_word.Content.Text
                    doc_word.Close(SaveChanges=False)

                    village_match = re.search(
                        r"Obiekt:\s*(.+?)(?=\s{2,}|\t|\r|\n|$)",
                        text_content,
                        re.IGNORECASE,
                    )
                    area_match = re.search(
                        r"Razem\s*[^0-9A-Za-z]*([\d\s]+(?:[\.,]\d+)?)",
                        text_content,
                        re.IGNORECASE,
                    )

                    village_name = (
                        village_match.group(1).strip().upper()
                        if village_match
                        else "NIEZNANA_WIES"
                    )
                    area_str = (
                        area_match.group(1).replace(" ", "").strip()
                        if area_match
                        else "[BRAK_DANYCH]"
                    )

                    doc = Document(template_path)
                    self.replace_text_robust(doc, village_ph, village_name)
                    self.replace_text_robust(doc, area_ph, area_str)

                    target_path = optax_path.parent / "STR_TYT.docx"
                    doc.save(str(target_path))
                    self.log(
                        f"  └─ Utworzono: {target_path.parent.name}/STR_TYT.docx (Wieś: {village_name})"
                    )

                except Exception as e:
                    self.log(
                        f"  [Błąd] Nie udało się wygenerować STR_TYT dla {optax_path.parent.name}: {e}"
                    )
        finally:
            try:
                word_app.Quit()
            except:
                pass

    # NOWA METODA: Wstrzykiwanie Skrótów i Symboli do pakietów wsi
    def task_inject_skroty(self, pdf_dir, skroty_source_path):
        pdf_dir = Path(pdf_dir)
        skroty_source_path = Path(skroty_source_path)

        if not skroty_source_path.exists():
            self.log("[SKROTY] Plik nie istnieje. Pomijam.")
            return 0

        ext = skroty_source_path.suffix.lower()
        temp_skroty_pdf = None
        skroty_pdf_to_copy = None

        if ext in {".doc", ".docx"}:
            self.log("[SKROTY] Konwertuję plik Word na PDF...")
            word_app = None
            try:
                word_app = win32com.client.DispatchEx("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0
                doc = word_app.Documents.Open(str(skroty_source_path.resolve()), AddToRecentFiles=False)
                temp_skroty_pdf = Path(tempfile.gettempdir()) / "skroty_temp.pdf"
                doc.ExportAsFixedFormat(
                    OutputFileName=str(temp_skroty_pdf),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    Range=0,
                    Item=0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )
                doc.Close(False)
                skroty_pdf_to_copy = temp_skroty_pdf
            except Exception as e:
                self.log(f"[SKROTY] Błąd konwersji: {e}")
                return 0
            finally:
                if word_app is not None:
                    try:
                        word_app.Quit()
                    except:
                        pass
        elif ext == ".pdf":
            skroty_pdf_to_copy = skroty_source_path
        else:
            self.log(f"[SKROTY] Nieobsługiwany format: {ext}")
            return 0

        # Znajdź wszystkie foldery które zawierają pliki PDF (oprócz skroty.pdf)
        pdf_folders = set()
        for p in pdf_dir.rglob("*.pdf"):
            if p.name.lower() != "skroty.pdf":
                pdf_folders.add(p.parent)

        count = 0

        if skroty_pdf_to_copy is not None:
            for folder in pdf_folders:
                target_skroty = folder / "skroty.pdf"
                try:
                    shutil.copy2(skroty_pdf_to_copy, target_skroty)
                    count += 1
                except Exception as e:
                    self.log(f"[SKROTY] Błąd kopiowania do {folder.name}: {e}")

        if temp_skroty_pdf and temp_skroty_pdf.exists():
            try:
                temp_skroty_pdf.unlink()
            except:
                pass

        return count

    def _flatten_001_subfolders(self, root_dir):
        """Wyciąga pliki z podfolderów *.001 do folderu nadrzędnego i usuwa puste podfoldery."""
        root_dir = Path(root_dir)
        if not root_dir.exists():
            return
        for sub in sorted(root_dir.rglob("*.001"), reverse=True):
            if not sub.is_dir():
                continue
            for f in sub.iterdir():
                if f.is_file():
                    target = sub.parent / f.name
                    if target.exists():
                        target.unlink()
                    shutil.move(str(f), str(target))
            try:
                sub.rmdir()
                self.log(f"  [SPŁASZCZONO] {sub.parent.name}/{sub.name} → {sub.parent.name}/")
            except Exception:
                pass

    def run_logic_thread(self, src_str, out_str, mode, remove_names, margins_dict=None):
        # --- INICJALIZACJA ZMIENNYCH ---
        in_root = None
        out_root = None
        dir_01, dir_02, dir_03, dir_04, dir_05 = None, None, None, None, None
        # -------------------------------

        pythoncom.CoInitialize()
        try:
            in_root = Path(src_str)
            out_root = Path(out_str)
            out_root.mkdir(parents=True, exist_ok=True)

            if mode == "ALL":
                dir_01, dir_02, dir_03, dir_04, dir_05 = (
                    out_root / "TXT",
                    out_root / "Word",
                    out_root / "PDF",
                    out_root / "PDF Polaczone",
                    out_root / "PDF bez pustych stron",
                )

                self.reset_dashboard()

                self.update_dashboard(0, "running", "Czyszczenie...")
                self.check_stop()
                c1 = self.task_clean_txt(in_root, dir_01)
                self._flatten_001_subfolders(dir_01)
                self.update_dashboard(0, "done", f"{c1} plików")
                self.set_progress(0.15)

                self.update_dashboard(1, "running", "Kompilacja...")
                self.check_stop()
                self.task_word_processing_subprocess(dir_01, dir_02, remove_names, margins_dict=margins_dict)
                self._flatten_001_subfolders(dir_02)
                self.update_dashboard(1, "done", "Gotowe")
                self.set_progress(0.30)

                # === GENEROWANIE STR_TYT ===
                if (
                        getattr(self, "all_gen_str_tyt_var", None)
                        and self.all_gen_str_tyt_var.get()
                ):
                    self.update_status(
                        "Generowanie stron tytułowych (STR_TYT)...", "#0078D7"
                    )
                    template_path = self.all_template_entry.get().strip()
                    v_ph = "NAZWA WSI"
                    a_ph = "wielkość"
                    if template_path and Path(template_path).exists():
                        self.task_generate_str_tyt(dir_02, template_path, v_ph, a_ph)
                    else:
                        self.log(
                            "[UWAGA] Zaznaczono generowanie STR_TYT, ale nie podano prawidłowego szablonu. Pomijam."
                        )
                self.set_progress(0.45)

                self.update_dashboard(2, "running", "Konwersja...")
                self.check_stop()
                c3 = self.task_convert_to_pdf(dir_02, dir_03)
                self._flatten_001_subfolders(dir_03)
                self.update_dashboard(2, "done", f"{c3} plików")
                self.set_progress(0.60)

                # === WSTRZYKIWANIE SKROTÓW (ZAWSZE WŁĄCZONE) ===
                self.update_status("Dołączanie 'Skrótów i symboli' do pakietów...", "#0078D7")

                skroty_path = self._resolve_skroty_path()

                # Przystępujemy do dołączenia pliku
                if skroty_path and Path(skroty_path).exists():
                    c_skroty = self.task_inject_skroty(dir_03, skroty_path)
                    self.log(f"[SKROTY] Dodano plik do {c_skroty} folderów wsi.")
                else:
                    self.log("[UWAGA] Nie znaleziono pliku ze skrótami (ani domyślnego, ani własnego). Pomijam.")

                self.update_dashboard(3, "running", "Scalanie...")
                self.check_stop()
                c4 = self.task_merge_pdfs(dir_03, dir_04, mode_key="ALL")
                self.update_dashboard(3, "done", f"{c4} pakietów")
                self.set_progress(0.80)

                self.update_dashboard(4, "running", "Weryfikacja...")
                self.check_stop()
                c5 = self.task_remove_blank_pages(dir_04, dir_05)
                self.update_dashboard(4, "done", f"{c5} plików")

            elif mode == "WORD":
                dir_01, dir_02 = out_root / "TXT", out_root / "Word"
                file_filter = self.get_selected_word_filters()
                filter_label = ", ".join(self.get_selected_word_filters())
                self.check_stop()
                self.update_status(
                    f"ETAP 1/2: Oczyszczanie plików TXT ({filter_label})", "#0078D7"
                )
                self.task_clean_txt(in_root, dir_01, file_filter)
                self.set_progress(0.5)
                self.check_stop()
                self.update_status(
                    f"ETAP 2/2: Przetwarzanie i konwersja Word ({filter_label})",
                    "#0078D7",
                )
                self.task_word_processing_subprocess(
                    dir_01, dir_02, remove_names, file_filter, margins_dict=margins_dict
                )
                self._flatten_001_subfolders(dir_01)
                self._flatten_001_subfolders(dir_02)

            elif mode == "PDF":
                dir_03, dir_04, dir_05 = (
                    out_root / "PDF",
                    out_root / "PDF Polaczone",
                    out_root / "PDF bez pustych stron",
                )
                do_merge = getattr(
                    self, "pdf_merge_var", ctk.BooleanVar(value=True)
                ).get()
                self.check_stop()
                (
                    self.update_status(
                        "ETAP 1/3: Zmiana formatu z Word na PDF", "#0078D7"
                    )
                    if do_merge
                    else self.update_status(
                        "Trwa zmiana formatu z Word na PDF...", "#0078D7"
                    )
                )
                self.task_convert_to_pdf(in_root, dir_03)
                self._flatten_001_subfolders(dir_03)

                # === WSTRZYKIWANIE SKROTÓW (AUTOMATYCZNIE Z SZABLONU Skroty.docx) ===
                self.update_status("Dołączanie 'Skrótów i symboli' do pakietów...", "#0078D7")
                skroty_path = self._resolve_skroty_path()
                if skroty_path and Path(skroty_path).exists():
                    c_skroty = self.task_inject_skroty(dir_03, skroty_path)
                    self.log(f"[SKROTY] Dodano plik do {c_skroty} folderów wsi.")
                else:
                    self.log(
                        "[UWAGA] Nie znaleziono pliku ze skrótami (ani domyślnego, ani własnego). Pomijam."
                    )

                self.set_progress(0.4 if do_merge else 1.0)
                if do_merge:
                    self.check_stop()
                    self.update_status(
                        "ETAP 2/3: Logiczna integracja dokumentacji", "#0078D7"
                    )
                    self.task_merge_pdfs(dir_03, dir_04, mode_key="PDF")
                    self.set_progress(0.7)
                    self.check_stop()
                    self.update_status("ETAP 3/3: Usuwanie anomalii", "#0078D7")
                    self.task_remove_blank_pages(dir_04, dir_05)

            self.log("\nZAKOŃCZONO POMYŚLNIE.")
            self.set_progress(1.0)
            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            self.after(0, lambda: messagebox.showinfo("Sukces", "Zadanie zakończone."))
        except InterruptedError:
            self.update_status("Przerwano", "#D83B01", animate=False)
            self.log("\nZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
        except Exception as e:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            pythoncom.CoUninitialize()
            self.running = False
            self.after(0, self.restore_all_buttons)

