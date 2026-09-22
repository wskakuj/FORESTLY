# -*- coding: utf-8 -*-
"""
Forestly — generator „opisu ogólnego" (opis og) dla wsi.
======================================================================
Zależności: python-docx, szablon opis_og_szablon.docx (repo / EXE)

Co robi:
  Dla każdego folderu wsi (folder z plikiem WSK_ZB.doc) tworzy plik
  „opis og_<nazwa wsi>.docx" — opis ogólny uproszczonego planu.

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

import re
import threading
import traceback
from pathlib import Path

TPL_FILENAME = "opis_og_szablon.docx"
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
        """Czyta liczby z binarnego WSK_ZB.doc (bez uruchamiania Worda)."""
        try:
            data = Path(path).read_bytes()
        except OSError as e:
            return None, f"nie można odczytać pliku ({e})"
        for enc in ("utf-16-le", "cp1250", "utf-8"):
            vals = cls._parse_wsk_zb_text(data.decode(enc, errors="ignore"))
            if vals:
                return vals, None
        return None, "nie znaleziono danych o użytkowaniu lasu w pliku"

    # ------------------------------------------------ Wypełnianie szablonu
    @staticmethod
    def _fill_template(tpl_path, values, out_path):
        """Wypełnia opis_og_szablon.docx liczbami i zapisuje jako .docx."""
        from docx import Document
        import copy

        doc = Document(str(tpl_path))
        vals = dict(values)
        vals["FORMY_OCHRONY"] = "Zlokalizowano następujące formy ochrony przyrody:"

        # tabela etatów (jedyna 6-kolumnowa tabela w szablonie)
        for table in doc.tables:
            if len(table.columns) == 6 and len(table.rows) >= 2:
                for cell in table.rows[1].cells:
                    ph = cell.text.strip()
                    if ph.startswith("{{") and ph.endswith("}}"):
                        key = ph[2:-2]
                        if key in vals and cell.paragraphs and cell.paragraphs[0].runs:
                            cell.paragraphs[0].runs[0].text = vals[key]

        # akapity + marker do ręcznego wpisania form ochrony przyrody
        marker_done = False
        for p in doc.paragraphs:
            txt = p.text
            m = re.search(r"\{\{(\w+)\}\}", txt)
            if not m or m.group(1) not in vals:
                continue
            new_txt = re.sub(r"\{\{(\w+)\}\}", lambda k: vals[k.group(1)], txt)
            if p.runs:
                p.runs[0].text = new_txt
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.add_run(new_txt)
            if m.group(1) == "FORMY_OCHRONY" and not marker_done:
                # dodaj pod spodem wyraźny marker do ręcznego wpisania
                new_p = copy.deepcopy(p._p)
                p._p.addnext(new_p)
                from docx.text.paragraph import Paragraph
                marker = Paragraph(new_p, p._parent)
                if marker.runs:
                    marker.runs[0].text = OCHRONA_MARKER
                    for r in marker.runs[1:]:
                        r.text = ""
                else:
                    marker.add_run(OCHRONA_MARKER)
                marker_done = True

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
        from app.core.word_worker import get_resource_path

        entry = getattr(self, "opis_og_root_entry", None)
        raw = entry.get().strip() if entry else ""
        if not raw:
            self.log("[OPIS OG] Wskaż najpierw folder (główny z folderami wsi albo pojedynczą wieś).")
            self.update_status("Brak folderu", "#D83B01", animate=False)
            return
        root = Path(raw)
        if not root.exists():
            self.log(f"[OPIS OG] Folder nie istnieje: {root}")
            self.update_status("Brak folderu", "#D83B01", animate=False)
            return

        # pojedyncza wieś (jest WSK_ZB.doc) czy folder nadrzędny z wsiami?
        if (root / "WSK_ZB.doc").exists():
            village_dirs = [root]
        else:
            village_dirs = sorted(
                d for d in root.iterdir()
                if d.is_dir() and (d / "WSK_ZB.doc").exists()
            )
        if not village_dirs:
            self.log("[OPIS OG] Nie znaleziono żadnego folderu wsi z plikiem WSK_ZB.doc "
                     f"w: {root}")
            self.update_status("Brak wsi", "#D83B01", animate=False)
            return

        tpl = get_resource_path(TPL_FILENAME)
        if not Path(tpl).exists():
            self.log(f"[OPIS OG BŁĄD] Brak szablonu: {tpl}")
            self.update_status("Brak szablonu", "#D83B01", animate=False)
            return

        self.log(f"[OPIS OG] Generator start — wsi: {len(village_dirs)}, "
                 f"szablon: {Path(tpl).name}")
        done = 0
        for i, d in enumerate(village_dirs, 1):
            self.update_status(f"Opis ogólny: {d.name} ({i}/{len(village_dirs)})",
                               "#0078D7")
            vals, err = self._read_wsk_zb_doc(d / "WSK_ZB.doc")
            if vals is None:
                self.log(f"[OPIS OG] {d.name}: POMINIĘTO — {err}")
                continue

            # nazwa pliku: zachowaj pisownię z istniejącego "opis og_*", jeśli jest
            out_base = f"opis og_{d.name}"
            for existing in d.glob("opis og_*"):
                out_base = existing.name.rsplit(".", 1)[0]
                break
            out_path = d / (out_base + ".docx")

            try:
                self._fill_template(tpl, vals, out_path)
            except Exception as e:
                self.log(f"[OPIS OG] {d.name}: BŁĄD zapisu — {e}")
                continue

            done += 1
            self.log(f"[OPIS OG] {d.name}: zapisano {out_path.name} "
                     f"(rębne {vals['MAKS_MIAZSZOSC']} m3, "
                     f"przedrębne {vals['UZYTK_PRZEDRZEBNE']} m3)")

        self.log(f"[OPIS OG] Zakończono — wygenerowano {done}/{len(village_dirs)}. "
                 "Pamiętaj o ręcznym wpisaniu form ochrony przyrody "
                 "(Natura 2000 itp.) w wygenerowanych plikach.")
        self.update_status("Opis ogólny gotowy", "#107C10", animate=False)

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
                 "(Natura 2000 itp.) wpisujesz ręcznie — w pliku zostanie wyraźnie "
                 "oznaczone miejsce.",
            font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#888888",
            wraplength=700, justify="left",
        ).grid(row=1, column=0, columnspan=3, padx=15, pady=(0, 15), sticky="w")
        ctk.CTkButton(
            scroll_frame, text="Generuj opisy ogólne", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0", hover_color="#005A9E", height=44, corner_radius=6,
            command=self.start_opis_og_pipeline,
        ).grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")
