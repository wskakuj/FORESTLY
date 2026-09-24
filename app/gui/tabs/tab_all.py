"""
Forestly — Mixin: TabAllMixin
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
    is_file_locked, load_margins, save_margins, add_tooltip,
)

from app.core.word_worker import (
    get_resource_path,
)
from app.core.wydruki import (
    generuj_wszystkie_po_przeniesieniu, generuj_halizny_txt,
)
from app.gui.tabs.tab_wydruki import AGENCJA_NAGLOWKA

class TabAllMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def _setup_all_extras(self, card_frame, row_idx):
        """Zakładka 1-Click: klasyczny formularz chowam — steruje kreator
        otwierany w osobnym oknie (spójnie z webowym GUI)."""
        scroll = card_frame.master
        card_frame.grid_remove()          # formularz (źródło/cel, przyciski) ukryty
        # dashboard i przyciski powstają DOPIERO PO extra_ui_setup — chowam
        # je odrobinę później (całość poza kartą startową)
        self.after(80, self._hide_all_tab_behind_launcher)

        box = ctk.CTkFrame(scroll, fg_color="#252526", corner_radius=8,
                           border_width=1, border_color="#333333")
        box.grid(row=0, column=0, padx=10, pady=(10, 8), sticky="new")
        self._all_launcher_box = box
        box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            box, text="Pełny Automat (1-Click)",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=20, pady=(18, 4), sticky="w")
        ctk.CTkLabel(
            box,
            text=("Halizny → TXT z DBF → pliki Word (ze stronami tytułowymi i opisami ogólnymi)\n"
                  "→ PDF → scalenie w jeden dokument.\n"
                  "Kreator przeprowadzi Cię przez cztery krótkie kroki — potem zrobi wszystko sam."),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#888888", justify="left",
        ).grid(row=1, column=0, padx=20, sticky="w")
        ctk.CTkButton(
            box, text="Zaczynamy — otwórz kreatora", height=46,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=self._open_all_wizard,
        ).grid(row=2, column=0, padx=20, pady=(10, 18), sticky="w")

    # ================== KREATOR 1-CLICK (osobne okno) ==================
    # Wejście w zakładkę (albo kliknięcie „Otwórz kreatora") otwiera okno:
    # opis → lokalizacje → strona tytułowa i daty → marginesy → nazwiska →
    # podsumowanie z „Generuj dokumenty" → postęp i „Ukończono!".
    # Kontrolki tworzone są RAZ (przy pierwszym otwarciu) i żyją w ukrytym
    # oknie — zamknięcie okna nie traci wpisanych danych.

    def _open_all_wizard(self):
        w = getattr(self, "_all_wiz", None)
        if w is not None:
            try:
                w.deiconify()
                w.lift()
                return
            except Exception:
                pass
        wiz = ctk.CTkToplevel(self)
        wiz.title("Pełny Automat (1-Click) — kreator")
        WW, WH = 900, 680
        try:
            self.update_idletasks()
            mx, my = self.winfo_rootx(), self.winfo_rooty()
            mw, mh = self.winfo_width(), self.winfo_height()
            x = mx + max(0, (mw - WW) // 2)
            y = max(0, my + max(0, (mh - WH) // 4))
        except Exception:
            x, y = 100, 100
        wiz.geometry(f"{WW}x{WH}+{x}+{y}")
        wiz.configure(fg_color="#1E1E1E")
        wiz.grid_columnconfigure(0, weight=1)
        wiz.grid_rowconfigure(1, weight=1)
        wiz.protocol("WM_DELETE_WINDOW", self._all_wiz_close)
        self._all_wiz = wiz
        self._all_wiz_step = 0
        self._all_wiz_started = False
        self._all_wiz_was_running = False
        self._all_wiz_dots_n = 0

        head = ctk.CTkFrame(wiz, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 2))
        ctk.CTkLabel(
            head, text="Pełny Automat (1-Click)",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color="#E0E0E0",
        ).pack(side="left")
        dots = ctk.CTkFrame(head, fg_color="transparent")
        dots.pack(side="left", expand=True)
        self._all_wiz_dots = []
        for _ in range(5):
            d = ctk.CTkLabel(dots, text="●", width=20,
                             font=ctk.CTkFont(family="Segoe UI", size=11),
                             text_color="#555555")
            d.pack(side="left")
            self._all_wiz_dots.append(d)
        ctk.CTkButton(head, text="✕", width=38, height=30,
                      fg_color="transparent", hover_color="#333333",
                      font=ctk.CTkFont(size=15),
                      command=self._all_wiz_close).pack(side="right")

        body = ctk.CTkFrame(wiz, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=18, pady=6)
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        nav = ctk.CTkFrame(wiz, fg_color="transparent")
        nav.grid(row=2, column=0, sticky="ew", padx=18, pady=(2, 14))
        self._all_wiz_back = ctk.CTkButton(
            nav, text="‹  Wstecz", width=110, height=38, fg_color="#333333",
            hover_color="#444444", command=self._all_wiz_prev)
        self._all_wiz_back.pack(side="left")
        self._all_wiz_next = ctk.CTkButton(
            nav, text="Zaczynamy  ›", width=210, height=42,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=self._all_wiz_next_step)
        self._all_wiz_next.pack(side="right")

        self._all_wiz_build_steps(body)
        self._all_wiz_show(0)

    def _all_wiz_close(self):
        """Chowam okno kreatora — bez utraty wpisanych danych."""
        w = getattr(self, "_all_wiz", None)
        if w is not None:
            try:
                w.withdraw()
            except Exception:
                pass

    def _all_wiz_prev(self):
        if getattr(self, "_all_wiz_step", 0) > 0 and not getattr(self, "_all_wiz_started", False):
            self._all_wiz_show(self._all_wiz_step - 1)

    def _all_wiz_next_step(self):
        krok = getattr(self, "_all_wiz_step", 0)
        if krok < 5:
            self._all_wiz_show(krok + 1)

    def _all_wiz_show(self, step):
        self._all_wiz_step = step
        for i, f in enumerate(self._all_wiz_steps):
            if i == step:
                f.grid()
            else:
                f.grid_remove()
        for i, d in enumerate(self._all_wiz_dots):
            d.configure(text_color="#2dd4a7" if i == step
                        else ("#3a7a68" if i < step else "#555555"))
        # przyciski nawigacji są w nav zarządzane przez pack — używamy
        # pack/pack_forget (grid na packowanym widżetie rzuca TclError)
        if step == 0:
            self._all_wiz_back.pack_forget()
            self._all_wiz_next.pack(side="right")
            self._all_wiz_next.configure(text="Zaczynamy  ›")
        elif step < 5:
            self._all_wiz_back.pack(side="left")
            self._all_wiz_next.pack(side="right")
            self._all_wiz_next.configure(text="Dalej  ›")
        elif step == 5:
            self._all_wiz_back.pack(side="left")
            self._all_wiz_next.pack_forget()
            self._all_wiz_refresh_summary()
        else:  # postęp
            self._all_wiz_back.pack_forget()
            self._all_wiz_next.pack_forget()

    def _toggle_all_gdos_ui(self):
        """Pole folderu GDOŚ widoczne tylko przy włączonych opisach ogólnych."""
        if getattr(self, "all_gen_opis_og_var", None) is None:
            return
        if self.all_gen_opis_og_var.get():
            self.all_gdos_frame.grid()
        else:
            self.all_gdos_frame.grid_remove()

    def _all_wiz_names_refresh(self):
        """Podpis pod przełącznikiem nazwisk — informacja o skutku wyboru."""
        on = bool(self.remove_names_var.get())
        self._all_wiz_names_cap.configure(
            text=("Nazwiska właścicieli zostaną usunięte z REJESTRU." if on
                  else "REJESTR zostanie wygenerowany z pełnymi\n"
                       "nazwiskami właścicieli."),
            text_color="#34d399" if on else "#888888")

    def _all_wiz_build_steps(self, body):
        from app.config import TERRITORY_DATA
        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_norm = ctk.CTkFont(family="Segoe UI", size=12)
        S = []

        # ---------- krok 0: ekran powitalny ----------
        f0 = ctk.CTkFrame(body, fg_color="transparent")
        f0.grid(row=0, column=0, sticky="nsew")
        f0.grid_columnconfigure(0, weight=1)
        powitalny = ctk.CTkFrame(f0, fg_color="transparent")
        powitalny.place(relx=0.5, rely=0.42, anchor="center")
        ctk.CTkLabel(
            powitalny, text="Cały proces jednym kliknięciem",
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
            text_color="#E0E0E0").pack(pady=(0, 10))
        ctk.CTkLabel(
            powitalny,
            text=("Halizny → TXT z DBF → pliki Word (ze stronami tytułowymi i opisami ogólnymi)\n"
                  "→ PDF → scalenie w jeden dokument.\n\n"
                  "Przeprowadzę Cię przez cztery krótkie kroki — potem zrobię wszystko sam."),
            font=font_norm, text_color="#888888", justify="center").pack()
        S.append(f0)

        # ---------- krok 1: lokalizacje ----------
        f1 = ctk.CTkFrame(body, fg_color="transparent")
        f1.grid(row=0, column=0, sticky="nsew")
        f1.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            f1, text="Gdzie są mietki i gdzie zapisać wyniki?",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#E0E0E0",
        ).grid(row=0, column=0, columnspan=3, padx=5, pady=(4, 14), sticky="w")
        ctk.CTkLabel(
            f1, text="Wskaż folder z danymi źródłowymi (obręby z plikami DBF)\noraz folder docelowy, w którym powstanie cała dokumentacja.",
            font=font_norm, text_color="#888888", justify="left",
        ).grid(row=1, column=0, columnspan=3, padx=5, pady=(0, 12), sticky="w")
        e_src = ctk.CTkEntry(f1, placeholder_text="Wskaż folder...", height=34)
        ctk.CTkLabel(f1, text="Folder źródłowy (MIETEK):", font=font_label,
                     text_color="#E0E0E0").grid(row=2, column=0, padx=(5, 10), pady=6, sticky="w")
        e_src.grid(row=2, column=1, padx=5, pady=6, sticky="ew")
        ctk.CTkButton(f1, text="Wybierz", width=90, height=34, fg_color="#333333",
                      hover_color="#444444",
                      command=lambda: self.select_dir(e_src)).grid(row=2, column=2, padx=(5, 5), pady=6)
        e_dst = ctk.CTkEntry(f1, placeholder_text="Wskaż folder...", height=34)
        ctk.CTkLabel(f1, text="Folder docelowy (wyniki):", font=font_label,
                     text_color="#E0E0E0").grid(row=3, column=0, padx=(5, 10), pady=6, sticky="w")
        e_dst.grid(row=3, column=1, padx=5, pady=6, sticky="ew")
        ctk.CTkButton(f1, text="Wybierz", width=90, height=34, fg_color="#333333",
                      hover_color="#444444",
                      command=lambda: self.select_dir(e_dst)).grid(row=3, column=2, padx=(5, 5), pady=6)
        # wpisy 1-Click wskazują teraz na pola kreatora („btn" zostaje — używa go UI)
        self.entries["ALL"]["src"] = e_src
        self.entries["ALL"]["dst"] = e_dst
        S.append(f1)

        # ---------- krok 2: strona tytułowa i daty ----------
        f2 = ctk.CTkScrollableFrame(body, fg_color="transparent")
        f2.grid(row=0, column=0, sticky="nsew")
        f2.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            f2, text="Strona tytułowa i daty",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=5, pady=(4, 4), sticky="w")
        ctk.CTkLabel(
            f2, text="Z tych danych powstanie strona tytułowa każdej wsi. „Stan na” zastępuje daty\nwe wszystkich dokumentach Word, a pola 10-lecia — okres w WSK_ZB.",
            font=font_norm, text_color="#888888", justify="left",
        ).grid(row=1, column=0, padx=5, pady=(0, 10), sticky="w")

        body2 = ctk.CTkFrame(f2, fg_color="#1E1E1E",
                             border_width=1, border_color="#333333")
        body2.grid(row=2, column=0, padx=5, pady=(0, 10), sticky="ew")
        body2.grid_columnconfigure(1, weight=1)

        woj_list = sorted(TERRITORY_DATA.keys()) if TERRITORY_DATA else ["BRAK DANYCH"]
        default_woj = ("KUJAWSKO-POMORSKIE" if "KUJAWSKO-POMORSKIE" in TERRITORY_DATA
                       else (woj_list[0] if woj_list else ""))
        powiat_list = sorted(TERRITORY_DATA.get(default_woj, {}).keys()) or ["BRAK DANYCH"]
        default_powiat = ("TUCHOLSKI" if "TUCHOLSKI" in TERRITORY_DATA.get(default_woj, {})
                          else (powiat_list[0] if powiat_list else ""))
        gmina_list = list(TERRITORY_DATA.get(default_woj, {}).get(default_powiat, []))
        default_gmina = ("LUBIEWO" if "LUBIEWO" in gmina_list
                         else (gmina_list[0] if gmina_list else ""))

        def _row(r, label, widget):
            ctk.CTkLabel(body2, text=label, font=font_label,
                         text_color="#E0E0E0").grid(
                row=r, column=0, padx=(10, 10), pady=4, sticky="w")
            widget.grid(row=r, column=1, padx=(0, 10), pady=4, sticky="ew")

        self.all_tpl_doc_var = ctk.StringVar(value="UPUL")
        _row(0, "Typ dokumentu:", ctk.CTkOptionMenu(
            body2, values=["UPUL", "ISL"], variable=self.all_tpl_doc_var, height=30))
        self.all_tpl_prefix_var = ctk.StringVar(value="położonych na terenie obrębu")
        _row(1, "Prefiks obrębu:", ctk.CTkOptionMenu(
            body2, values=["położonych na terenie obrębu", "Obręb:"],
            variable=self.all_tpl_prefix_var, height=30))
        self.all_tpl_woj_var = ctk.StringVar(value=default_woj)
        self.all_tpl_woj_box = ctk.CTkComboBox(
            body2, values=woj_list, variable=self.all_tpl_woj_var, height=30,
            command=lambda _v: self._all_tpl_refresh_powiat())
        _row(2, "Województwo (można wpisać własne):", self.all_tpl_woj_box)
        self.all_tpl_powiat_var = ctk.StringVar(value=default_powiat)
        self.all_tpl_powiat_box = ctk.CTkComboBox(
            body2, values=powiat_list, variable=self.all_tpl_powiat_var,
            height=30, command=lambda _v: self._all_tpl_refresh_gmina())
        _row(3, "Powiat (można wpisać własny):", self.all_tpl_powiat_box)
        self.all_tpl_gmina_var = ctk.StringVar(value=default_gmina)
        self.all_tpl_gmina_box = ctk.CTkComboBox(
            body2, values=gmina_list, variable=self.all_tpl_gmina_var, height=30)
        _row(4, "Gmina (można wpisać własną):", self.all_tpl_gmina_box)

        self.all_tpl_stan_na_entry = ctk.CTkEntry(body2, height=30)
        self.all_tpl_stan_na_entry.insert(0, "30.06.2026 r.")
        _row(5, "Stan na (także data we wszystkich Wordach):",
             self.all_tpl_stan_na_entry)
        self.all_tpl_okres_entry = ctk.CTkEntry(body2, height=30)
        self.all_tpl_okres_entry.insert(0, "01.01.2027 – 31.12.2036 r.")
        _row(6, "Na okres (strona tytułowa):", self.all_tpl_okres_entry)

        daty_frame = ctk.CTkFrame(body2, fg_color="#252526", corner_radius=6)
        daty_frame.grid(row=7, column=0, columnspan=2, padx=10, pady=(4, 8), sticky="ew")
        daty_frame.grid_columnconfigure((1, 3), weight=1)
        ctk.CTkLabel(daty_frame, text="WSK_ZB — 10-lecie:", font=font_label,
                     text_color="#888888").grid(row=0, column=0, padx=(10, 6),
                                                pady=6, sticky="w")
        self.all_wsk_od_entry = ctk.CTkEntry(daty_frame, height=28)
        self.all_wsk_od_entry.insert(0, "01-01-2027")
        self.all_wsk_od_entry.grid(row=0, column=1, padx=4, pady=6, sticky="ew")
        ctk.CTkLabel(daty_frame, text="do:", font=font_label,
                     text_color="#888888").grid(row=0, column=2, padx=4, sticky="w")
        self.all_wsk_do_entry = ctk.CTkEntry(daty_frame, height=28)
        self.all_wsk_do_entry.insert(0, "31-12-2036")
        self.all_wsk_do_entry.grid(row=0, column=3, padx=(4, 10), pady=6, sticky="ew")

        # --- skróty i symbole (własny plik, opcjonalnie) ---
        self.all_custom_skroty_var = ctk.BooleanVar(value=False)
        cb_skroty = ctk.CTkCheckBox(
            f2,
            text="Użyj własnego pliku 'Skróty i symbole' (zamiast domyślnego z programu)",
            variable=self.all_custom_skroty_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            command=self._toggle_all_skroty_ui,
        )
        cb_skroty.grid(row=3, column=0, padx=5, pady=(2, 2), sticky="w")
        self.all_skroty_frame = ctk.CTkFrame(
            f2, fg_color="#1E1E1E", border_width=1, border_color="#333333")
        self.all_skroty_frame.grid(row=4, column=0, padx=5, pady=(0, 10), sticky="ew")
        self.all_skroty_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.all_skroty_frame, text="Własny plik:", font=font_label,
                     text_color="#E0E0E0").grid(row=0, column=0, padx=(10, 10),
                                                pady=8, sticky="w")
        self.all_skroty_entry = ctk.CTkEntry(
            self.all_skroty_frame, placeholder_text="Wskaż własny plik ze skrótami...",
            height=32)
        self.all_skroty_entry.grid(row=0, column=1, padx=5, pady=8, sticky="ew")
        ctk.CTkButton(
            self.all_skroty_frame, text="Wybierz",
            command=lambda: self.select_file(
                self.all_skroty_entry,
                [("Word/PDF", "*.docx *.doc *.pdf"), ("Wszystkie pliki", "*.*")],
            ),
            width=90, height=32, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=(5, 10), pady=8)

        # --- opisy ogólne (zawsze z gotowych WSK_ZB.doc z pipeline) ---
        # Folder z wynikami GDOŚ nie jest już potrzebny — cała baza obszarów
        # ochrony przyrody jest w programie (zakładka GDOŚ / gdos_obszary.json)
        self.all_gen_opis_og_var = ctk.BooleanVar(value=True)
        cb_opis = ctk.CTkCheckBox(
            f2,
            text="Generuj opisy ogólne (opis og_<wieś>.docx) po plikach Word",
            variable=self.all_gen_opis_og_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            command=self._toggle_all_gdos_ui,
        )
        cb_opis.grid(row=5, column=0, padx=5, pady=(2, 2), sticky="w")

        # folder z wynikami GDOŚ — źródło form ochrony przyrody dla opisów
        self.all_gdos_frame = ctk.CTkFrame(
            f2, fg_color="#1E1E1E", border_width=1, border_color="#333333")
        self.all_gdos_frame.grid(row=6, column=0, padx=5, pady=(0, 10), sticky="ew")
        self.all_gdos_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.all_gdos_frame,
                     text="Folder z wynikami GDOŚ (opcjonalny):", font=font_label,
                     text_color="#E0E0E0").grid(row=0, column=0, padx=(10, 10),
                                                pady=5, sticky="w")
        self.all_gdos_entry = ctk.CTkEntry(
            self.all_gdos_frame,
            placeholder_text="Formy ochrony przyrody (NN_WIEŚ_wynik.xlsx) — puste = bez GDOŚ",
            height=32)
        self.all_gdos_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(
            self.all_gdos_frame, text="Wybierz",
            command=lambda: self.select_dir(self.all_gdos_entry),
            width=90, height=32, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=(5, 10), pady=5)

        self._toggle_all_skroty_ui()
        self._toggle_all_gdos_ui()
        S.append(f2)

        # ---------- krok 3: marginesy ----------
        f3 = ctk.CTkFrame(body, fg_color="transparent")
        f3.grid(row=0, column=0, sticky="nsew")
        f3.grid_columnconfigure(0, weight=1)
        self._build_margins_ui(f3, 0, "ALL", start_open=True)
        S.append(f3)

        # ---------- krok 4: nazwiska ----------
        f4 = ctk.CTkFrame(body, fg_color="transparent")
        f4.grid(row=0, column=0, sticky="nsew")
        f4.grid_columnconfigure(0, weight=1)
        karta = ctk.CTkFrame(f4, fg_color="#252526", corner_radius=10,
                             border_width=1, border_color="#333333")
        karta.grid(row=0, column=0, padx=10, pady=(6, 0), sticky="ew")
        karta.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            karta, text="Nazwiska w REJESTRZE",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=18, pady=(14, 4), sticky="w")
        ctk.CTkLabel(
            karta, text="Włącz, jeśli z wydruków REJESTR (oraz z 1. strony)\nmają zniknąć nazwiska właścicieli.",
            font=font_norm, text_color="#888888", justify="left",
        ).grid(row=1, column=0, padx=18, sticky="w")
        self._all_wiz_names_switch = ctk.CTkSwitch(
            karta, text="Usuwaj nazwiska z REJESTRU (oraz 1. stronę)",
            variable=self.remove_names_var,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            switch_width=52, switch_height=28,
            progress_color="#2dd4a7",
            command=self._all_wiz_names_refresh,
        )
        self._all_wiz_names_switch.grid(row=2, column=0, padx=18,
                                        pady=(10, 4), sticky="w")
        self._all_wiz_names_cap = ctk.CTkLabel(
            karta, text="", font=ctk.CTkFont(family="Segoe UI", size=13),
            justify="left")
        self._all_wiz_names_cap.grid(row=3, column=0, padx=18,
                                     pady=(0, 18), sticky="w")
        self._all_wiz_names_refresh()
        S.append(f4)

        # ---------- krok 5: podsumowanie i start ----------
        f5 = ctk.CTkFrame(body, fg_color="transparent")
        f5.grid(row=0, column=0, sticky="nsew")
        f5.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            f5, text="Wszystko gotowe!",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#E0E0E0",
        ).grid(row=0, column=0, padx=5, pady=(4, 4), sticky="w")
        ctk.CTkLabel(
            f5, text="Tak uruchomię proces — jeszcze możesz coś zmienić, wracając do poprzednich kroków.",
            font=font_norm, text_color="#888888",
        ).grid(row=1, column=0, padx=5, pady=(0, 10), sticky="w")
        self._all_wiz_sum = ctk.CTkFrame(f5, fg_color="transparent")
        self._all_wiz_sum.grid(row=2, column=0, padx=5, sticky="ew")
        self._all_wiz_sum.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            f5, text="▶   Generuj dokumenty", height=52,
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            command=self._all_wiz_run,
        ).grid(row=3, column=0, padx=5, pady=(16, 4), sticky="ew")
        ctk.CTkButton(
            f5, text="Skonfiguruj układ PDF…", height=36, fg_color="transparent",
            border_width=1, border_color="#555555", hover_color="#333333",
            command=lambda: self.open_mode_order_window(
                "ALL", self.entries["ALL"]["dst"]),
        ).grid(row=4, column=0, padx=5, pady=(2, 4), sticky="ew")
        S.append(f5)

        # ---------- krok 6: postęp ----------
        f6 = ctk.CTkFrame(body, fg_color="transparent")
        f6.grid(row=0, column=0, sticky="nsew")
        f6.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(f6, text="", height=40).grid(row=0, column=0)
        self._all_wiz_status = ctk.CTkLabel(
            f6, text="Rozpoczynam…",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#E0E0E0")
        self._all_wiz_status.grid(row=1, column=0, pady=(10, 8))
        self._all_wiz_bar = ctk.CTkProgressBar(
            f6, mode="indeterminate", height=14, corner_radius=7)
        self._all_wiz_bar.grid(row=2, column=0, padx=80, pady=(0, 10), sticky="ew")
        self._all_wiz_file = ctk.CTkLabel(
            f6, text="", font=font_norm, text_color="#888888")
        self._all_wiz_file.grid(row=3, column=0, pady=(0, 4))
        self._all_wiz_stop = ctk.CTkButton(
            f6, text="Przerwij zadanie", height=36, width=150,
            fg_color="#8B0000", hover_color="#A52A2A",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=self._all_wiz_stop_clicked)
        self._all_wiz_stop.grid(row=4, column=0, pady=(10, 2))
        self._all_wiz_done_lbl = ctk.CTkLabel(
            f6, text="", font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"))
        self._all_wiz_actions = ctk.CTkFrame(f6, fg_color="transparent")
        self._all_wiz_actions.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(
            self._all_wiz_actions, text="Otwórz folder wyników", height=40,
            fg_color="#333333", hover_color="#444444",
            command=self.open_last_output_dir,
        ).grid(row=0, column=0, padx=6, sticky="ew")
        ctk.CTkButton(
            self._all_wiz_actions, text="Zamknij", height=40,
            command=self._all_wiz_close,
        ).grid(row=0, column=1, padx=6, sticky="ew")
        S.append(f6)

        self._all_wiz_steps = S

    def _all_wiz_refresh_summary(self):
        """Odświeża listę podsumowania przed uruchomieniem."""
        w = self._all_wiz_sum
        for child in w.winfo_children():
            child.destroy()
        font_k = ctk.CTkFont(family="Segoe UI", size=12)
        font_v = ctk.CTkFont(family="Segoe UI", size=12, weight="bold")

        def _v(attr, domyslne=""):
            e = getattr(self, attr, None)
            if e is not None and hasattr(e, "get"):
                return e.get().strip() or domyslne
            return domyslne

        skroty = (self.all_custom_skroty_var.get()
                  if hasattr(self, "all_custom_skroty_var") else False)
        opis_og = (self.all_gen_opis_og_var.get()
                   if hasattr(self, "all_gen_opis_og_var") else False)
        wiersze = [
            ("Mietki (źródło)", self.entries["ALL"]["src"].get().strip() or "— nie wskazano —"),
            ("Folder wyników", self.entries["ALL"]["dst"].get().strip() or "— nie wskazano —"),
            ("Obszar", " / ".join(x for x in [
                _v("all_tpl_gmina_var"), _v("all_tpl_powiat_var"), _v("all_tpl_woj_var")] if x) or "—"),
            ("Stan na", _v("all_tpl_stan_na_entry") or "—"),
            ("Nazwiska w REJESTRZE",
             "usuwane z REJESTRU" if self.remove_names_var.get()
             else "REJESTR z pełnymi nazwiskami"),
            ("Własne skróty i symbole",
             (self.all_skroty_entry.get().strip() or "(nie wskazano pliku)") if skroty
             else "domyślne z programu"),
            ("Opisy ogólne",
             "z formami ochrony z GDOŚ" if (opis_og and self.all_gdos_entry.get().strip())
             else ("bez folderu GDOŚ" if opis_og else "pomijane")),
        ]
        for i, (k, val) in enumerate(wiersze):
            ctk.CTkLabel(w, text=k, font=font_k, text_color="#888888", anchor="w"
                         ).grid(row=i, column=0, padx=(10, 10), pady=3, sticky="w")
            ctk.CTkLabel(w, text=val, font=font_v, text_color="#E0E0E0", anchor="w"
                         ).grid(row=i, column=1, padx=(10, 10), pady=3, sticky="ew")

    def _hide_all_tab_behind_launcher(self):
        """Chowa na zakładce 1-Click wszystko poza kartą startową kreatora
        (dashboard powstaje po extra_ui_setup, więc dopiero tutaj)."""
        box = getattr(self, "_all_launcher_box", None)
        if box is None or not box.winfo_exists():
            return
        try:
            for child in box.master.winfo_children():
                if child is not box:
                    child.grid_remove()
        except Exception:
            pass

    def _all_wiz_run(self):
        try:
            self._all_wiz_stop.configure(state="normal", text="Przerwij zadanie")
            self._all_wiz_stop.grid()
        except Exception:
            pass
        if self.running:
            self._all_wiz_started = True
            self._all_wiz_show(6)
            self._all_wiz_set_done(
                False, "Zadanie już trwa — poczekaj na jego zakończenie.")
            return
        self._all_wiz_started = True
        self._all_wiz_was_running = False
        self._all_wiz_status.configure(text="Rozpoczynam…")
        self._all_wiz_done_lbl.grid_remove()
        self._all_wiz_actions.grid_remove()
        try:
            self._all_wiz_bar.configure(mode="indeterminate")
            self._all_wiz_bar.start()
        except Exception:
            pass
        self._all_wiz_show(6)
        self.start_pipeline("ALL")
        if not self.running:
            # walidacja odrzuciła zadanie (np. brak ścieżek) — nie czekamy
            self._all_wiz_set_done(False, "Nie udało się uruchomić — sprawdź ścieżki w kroku 1.")
        else:
            self._all_wiz_poll()

    def _all_wiz_poll(self):
        wiz = getattr(self, "_all_wiz", None)
        if wiz is None or not getattr(self, "_all_wiz_started", False):
            return
        if not wiz.winfo_exists() or getattr(self, "_all_wiz_step", -1) != 6:
            return
        txt = (getattr(self, "status_base_text", "") or "Przetwarzam…").strip()
        self._all_wiz_dots_n = (getattr(self, "_all_wiz_dots_n", 0) + 1) % 4
        self._all_wiz_status.configure(text=txt + "." * self._all_wiz_dots_n)
        czesci = []
        if getattr(self, "progress_current_file", ""):
            czesci.append(str(self.progress_current_file))
        if getattr(self, "progress_total", 0):
            czesci.append("({}/{})".format(
                getattr(self, "progress_current", 0), self.progress_total))
        self._all_wiz_file.configure(text="  ".join(czesci))
        if self._all_wiz_was_running and not self.running:
            ok = getattr(self, "status_color", "#0078D7") != "#D83B01"
            self._all_wiz_set_done(ok)
            return
        self._all_wiz_was_running = self.running
        wiz.after(350, self._all_wiz_poll)

    def _all_wiz_stop_clicked(self):
        """Bezpiecznie przerywa bieżące zadanie (jak przycisk w pasku statusu)."""
        try:
            self.cancel_process()
            self._all_wiz_stop.configure(state="disabled", text="Przerywanie…")
        except Exception:
            pass

    def _all_wiz_set_done(self, ok, text=None):
        try:
            self._all_wiz_bar.stop()
            self._all_wiz_bar.configure(mode="determinate")
            self._all_wiz_bar.set(1.0 if ok else 0.0)
        except Exception:
            pass
        self._all_wiz_done_lbl.configure(
            text=text if text else ("✓  Ukończono!" if ok
                                    else "✗  Zakończono z błędem — szczegóły w dzienniku"),
            text_color="#34d399" if ok else "#fb7185")
        self._all_wiz_done_lbl.grid(row=5, column=0, pady=(14, 4))
        self._all_wiz_actions.grid(row=6, column=0, padx=80, pady=(2, 10), sticky="ew")
        try:
            self._all_wiz_stop.grid_remove()
        except Exception:
            pass
        if not ok:
            self._all_wiz_started = False   # można wrócić (Wstecz) i poprawić
            self._all_wiz_back.grid()

    def _toggle_all_tpl_ui(self):
        self.all_tpl_open = not getattr(self, "all_tpl_open", False)
        if self.all_tpl_open:
            self.all_tpl_body.grid()
            self.all_tpl_header.configure(
                text="▾  Kreator strony tytułowej (strona tytułowa i daty)")
        else:
            self.all_tpl_body.grid_remove()
            self.all_tpl_header.configure(
                text="▸  Kreator strony tytułowej (strona tytułowa i daty)")

    def _all_tpl_refresh_powiat(self):
        from app.config import TERRITORY_DATA
        lista = sorted(TERRITORY_DATA.get(self.all_tpl_woj_var.get(), {}).keys()) \
            or ["BRAK DANYCH"]
        self.all_tpl_powiat_box.configure(values=lista)
        if self.all_tpl_powiat_var.get() not in lista:
            self.all_tpl_powiat_var.set(lista[0])
        self._all_tpl_refresh_gmina()

    def _all_tpl_refresh_gmina(self):
        from app.config import TERRITORY_DATA
        lista = list(TERRITORY_DATA.get(self.all_tpl_woj_var.get(), {})
                     .get(self.all_tpl_powiat_var.get(), [])) or ["BRAK DANYCH"]
        self.all_tpl_gmina_box.configure(values=lista)
        if self.all_tpl_gmina_var.get() not in lista:
            self.all_tpl_gmina_var.set(lista[0])

    def _zbuduj_szablon_str_tyt_dla_all(self):
        """Buduje tymczasowy szablon STR_TYT z ustawień kreatora w 1-Click."""
        def _v(attr, default=""):
            e = getattr(self, attr, None)
            return e.get().strip() if e is not None else default
        try:
            fd, tmp = tempfile.mkstemp(suffix=".docx", prefix="STR_TYT_ALL_")
            os.close(fd)
            out = self._build_str_tyt_doc(
                _v("all_tpl_doc_var", "UPUL") or "UPUL",
                _v("all_tpl_prefix_var") or "położonych na terenie obrębu",
                _v("all_tpl_woj_var").upper(),
                _v("all_tpl_powiat_var").upper(),
                _v("all_tpl_gmina_var").upper(),
                _v("all_tpl_stan_na_entry"),
                _v("all_tpl_okres_entry"),
                tmp, village="NAZWA WSI", keep_area=False)
            self.log("[STR_TYT] Zbudowano szablon bazowy z ustawień kreatora (1-Click).")
            return out
        except Exception:
            self.log("[STR_TYT] Nie udało się zbudować szablonu strony tytułowej:\n"
                     + traceback.format_exc())
            return None

    def _zamien_daty_txt(self, txt_dir):
        """Zamienia daty w wyczyszczonych TXT (przed konwersją na Word).

        'Stan na: <data>' — we wszystkich plikach (pole kreatora „Stan na");
        'w 10-leciu od <data> do <data>' — w plikach WSK_ZB (pola od/do)."""
        def _v(attr):
            e = getattr(self, attr, None)
            return e.get().strip() if e is not None else ""
        stan_na, wsk_od, wsk_do = (_v("all_tpl_stan_na_entry"),
                                   _v("all_wsk_od_entry"), _v("all_wsk_do_entry"))
        if not stan_na and not (wsk_od and wsk_do):
            return
        n_stan = n_wsk = 0
        for f in Path(txt_dir).rglob("*"):
            if not f.is_file() or f.suffix.lower() != ".txt":
                continue
            try:
                raw = f.read_bytes()
            except OSError:
                continue
            nowy = raw
            if stan_na:
                val = stan_na.encode("cp852", errors="replace")
                nowy, k = re.subn(
                    rb"(Stan\s+na:?\s*)[0-9]{2}[-./][0-9]{2}[-./][0-9]{4}(\s*r\.?)?",
                    lambda m, v=val: m.group(1) + v, nowy, flags=re.IGNORECASE)
                n_stan += k
            if wsk_od and wsk_do and f.stem.upper().startswith("WSK_ZB"):
                od = wsk_od.encode("cp852", errors="replace")
                do = wsk_do.encode("cp852", errors="replace")
                nowy, k = re.subn(
                    rb"(w\s+10-leciu\s+od\s+)[0-9][0-9.\-]*[0-9](\s+do\s+)[0-9][0-9.\-]*[0-9]",
                    lambda m, o=od, d=do: m.group(1) + o + m.group(2) + d, nowy)
                n_wsk += k
            if nowy != raw:
                try:
                    f.write_bytes(nowy)
                except OSError:
                    pass
        if n_stan or n_wsk:
            self.log(f"[DATY] Zamieniono \"Stan na\" {n_stan}×, "
                     f"okres 10-lecia WSK_ZB {n_wsk}×.")

    def _toggle_all_skroty_ui(self):
        if getattr(self, "all_custom_skroty_var", None) and getattr(self, "all_skroty_frame", None):
            if self.all_custom_skroty_var.get():
                # Jeśli checkbox jest zaznaczony, przywracamy ramkę z powrotem na ekran
                self.all_skroty_frame.grid()
            else:
                # Jeśli checkbox jest odznaczony, całkowicie ukrywamy ramkę
                self.all_skroty_frame.grid_remove()

    def task_generuj_txt(self, in_root):
        """Etap 0: generuje pliki TXT MIETEKA z DBF (jak 'Generowanie: MIETEK -> TXT').

        Szuka folderów z plikami O*.DBF (np. WOL.001) w drzewie źródłowym
        i generuje komplet wydruków w miejscu, obok plików DBF.
        Zwraca liczbę obrębów, dla których coś wygenerowano.
        """
        in_root = Path(in_root)
        # katalogi bezpośrednio zawierające O*.DBF
        kat_o = sorted({p.parent for p in in_root.rglob("*.DBF")
                        if p.name[:1].upper() == "O"})
        if not kat_o:
            self.log("[TXT] Nie znaleziono plików O*.DBF — pomijam generowanie z DBF.")
            return 0
        # obręb = folder nadrzędny (np. CHORZEWO nad WOL.001); gdy DBF-y leżą
        # bezpośrednio w folderze źródłowym, traktujemy go jako obręb
        obreby = sorted({d.parent if d.parent != in_root.parent else d for d in kat_o})
        n = 0
        for obr in obreby:
            self.check_stop()
            try:
                # 1) HALIZNY.TXT (przed przeniesieniem halizn)
                hp, hn = generuj_halizny_txt(obr, agencja=AGENCJA_NAGLOWKA)
                if hn:
                    self.log(f"  [TXT] {obr.name}: HALIZNY.TXT ({hn} wydzieleń)")
                    st, np_ = self.przenies_halizny_obreb(obr)
                    if st == "ok":
                        self.log(f"  [TXT] {obr.name}: przeniesiono halizny "
                                 f"w D*.DBF ({np_} rekordów)")
                    elif st == "blad":
                        self.log(f"  [TXT] {obr.name}: błąd przenoszenia halizn — "
                                 f"wynik może być niepełny.")
                # 2) komplet wydruków (po przeniesieniu halizn)
                out = generuj_wszystkie_po_przeniesieniu(obr, agencja=AGENCJA_NAGLOWKA)
                if out:
                    n += 1
                    self.log(f"  [TXT] {obr.name}: {', '.join(sorted(out))}")
                else:
                    self.log(f"  [TXT] {obr.name}: brak O*.DBF — pomijam.")
            except Exception as e:
                self.log(f"  [TXT] {obr.name}: błąd — {e}")
                traceback.print_exc()
        return n

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
        _zap = getattr(self, "_zapamietaj_folder_wynikow", None)
        if _zap is not None:
            _zap(self.last_output_dir)
        self._disable_ui_for_process()
        self.log(f"[{mode}] URUCHOMIENIE ZADANIA\nZ: {src_path}\nDo: {dst_path}\n"
                 f"Nazwiska w REJESTRZE: "
                 f"{'USUWANE' if remove_names_flag else 'zostają'}.")
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
    def _resolve_skroty_path(self):
        """Plik ze skrótami: własny (checkbox w Pełnym automacie) albo domyślny z zasobów."""
        if getattr(self, "all_custom_skroty_var", None) and self.all_custom_skroty_var.get():
            entry = getattr(self, "all_skroty_entry", None)
            if entry is not None:
                p = str(entry.get()).strip()
                if p:
                    return p
        # domyślny plik z zasobów programu
        domyslne = get_resource_path("Skroty.pdf")
        if not domyslne.exists():
            domyslne = get_resource_path("Skroty.docx")
        return str(domyslne) if domyslne.exists() else None

    def _inject_skroty_step(self, pdf_dir):
        """Dołącza 'Skróty i symbole' (skroty.pdf) do każdego folderu z PDF-ami."""
        skroty_path = self._resolve_skroty_path()
        if skroty_path and Path(skroty_path).exists():
            c = self.task_inject_skroty(pdf_dir, skroty_path)
            self.log(f"[SKROTY] Dodano plik do {c} folderów wsi.")
            return c
        self.log("[UWAGA] Nie znaleziono pliku ze skrótami (ani domyślnego, ani własnego). Pomijam.")
        return 0

    def task_inject_skroty(self, pdf_dir, skroty_source_path):
        pdf_dir = Path(pdf_dir)
        skroty_source_path = Path(skroty_source_path)

        if not pdf_dir.exists():
            self.log("[SKROTY] Brak folderu PDF — pomijam dołączanie skrótów.")
            return 0

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

                # === ETAP 0: GENEROWANIE TXT Z DBF MIETEKA ===
                if getattr(self, "all_gen_txt_var", None) is None or self.all_gen_txt_var.get():
                    self.update_status("Generowanie plików TXT z DBF mietka...", "#0078D7")
                    self.update_dashboard(0, "running", "Generowanie TXT...")
                    self.check_stop()
                    c0 = self.task_generuj_txt(in_root)
                    self.update_dashboard(0, "done", f"{c0} obrębów")
                else:
                    self.update_dashboard(0, "done", "Pominięto")
                self.set_progress(0.05)

                self.update_status("Czyszczenie plików TXT...", "#0078D7")
                self.update_dashboard(1, "running", "Czyszczenie...")
                self.check_stop()
                c1 = self.task_clean_txt(in_root, dir_01)
                self._flatten_001_subfolders(dir_01)
                self.update_dashboard(1, "done", f"{c1} plików")
                self.set_progress(0.15)

                if c1 == 0:
                    # nic do roboty — nie kontynuuj (Word/PDF nie mają na czym pracować)
                    self.update_dashboard(1, "error", "Brak TXT")
                    self.log(
                        "\n[BŁĄD] Brak plików TXT do przetworzenia.\n"
                        "Folder źródłowy nie zawiera plików TXT ani plików DBF mietka\n"
                        "(O*.DBF itd.), albo odznaczono 'Generuj pliki TXT z DBF mietka'.")
                    self.update_status("Błąd — brak plików TXT", "#D83B01", animate=False)
                    return

                self.update_status("Generowanie plików Word...", "#0078D7")
                self.update_dashboard(2, "running", "Kompilacja...")
                self.check_stop()
                self.task_word_processing_subprocess(dir_01, dir_02, remove_names, margins_dict=margins_dict)
                self._flatten_001_subfolders(dir_02)
                self.update_dashboard(2, "done", "Gotowe")
                self.set_progress(0.30)

                # === GENEROWANIE STR_TYT (zawsze — z kreatora w 1-Click) ===
                self.update_status(
                    "Generowanie stron tytułowych (STR_TYT)...", "#0078D7"
                )
                tpl_tmp = self._zbuduj_szablon_str_tyt_dla_all()
                if tpl_tmp:
                    try:
                        self.task_generate_str_tyt(dir_02, tpl_tmp,
                                                   "NAZWA WSI", "wielkość")
                    finally:
                        try:
                            Path(tpl_tmp).unlink()
                        except OSError:
                            pass
                self.set_progress(0.45)

                # === OPISY OGÓLNE (po plikach Word, przed konwersją do PDF) ===
                if (getattr(self, "all_gen_opis_og_var", None) is None
                        or self.all_gen_opis_og_var.get()):
                    self.update_status(
                        "Generowanie opisów ogólnych (opis og_<wieś>.docx)...",
                        "#0078D7",
                    )
                    self.check_stop()
                    gdos_raw = ""
                    ent = getattr(self, "all_gdos_entry", None)
                    if ent is not None:
                        gdos_raw = ent.get().strip()
                    try:
                        self._opis_og_generuj(
                            dir_02, Path(gdos_raw) if gdos_raw else None,
                            tylko_istniejace=True,
                        )
                    except Exception:
                        self.log("[OPIS OG] Błąd generowania opisów ogólnych:\n"
                                 + traceback.format_exc())

                self.update_status("Konwersja plików Word na PDF...", "#0078D7")
                self.update_dashboard(3, "running", "Konwersja...")
                self.check_stop()
                c3 = self.task_convert_to_pdf(dir_02, dir_03)
                self._flatten_001_subfolders(dir_03)
                self.update_dashboard(3, "done", f"{c3} plików")
                self.set_progress(0.60)

                # === WSTRZYKIWANIE SKROTÓW (ZAWSZE WŁĄCZONE) ===
                self.update_status("Dołączanie 'Skrótów i symboli' do pakietów...", "#0078D7")
                self._inject_skroty_step(dir_03)

                self.update_status("Scalanie pakietów PDF...", "#0078D7")
                self.update_dashboard(4, "running", "Scalanie...")
                self.check_stop()
                c4 = self.task_merge_pdfs(dir_03, dir_04, mode_key="ALL")
                self.update_dashboard(4, "done", f"{c4} pakietów")
                self.set_progress(0.80)

                # usuwanie pustych stron — bez osobnego kroku na dashboardzie
                self.update_status("Usuwanie pustych stron z PDF...", "#0078D7")
                self.check_stop()
                c5 = self.task_remove_blank_pages(dir_04, dir_05)

                # === PORZĄDKI: zostaje tylko finalny folder "PDF polaczone" ===
                self.update_status("Porządkowanie folderów wynikowych...", "#0078D7")
                try:
                    if dir_04 and dir_04.exists():
                        shutil.rmtree(dir_04)
                        self.log("[PORZĄDKI] Usunięto folder pośredni 'PDF Polaczone'.")
                except Exception as e:
                    self.log(f"[PORZĄDKI] Nie udało się usunąć 'PDF Polaczone': {e}")
                try:
                    if dir_05 and dir_05.exists():
                        dir_05.rename(out_root / "PDF polaczone")
                        self.log("[PORZĄDKI] Folder 'PDF bez pustych stron' "
                                 "przemianowano na 'PDF polaczone'.")
                except Exception as e:
                    self.log(f"[PORZĄDKI] Nie udało się zmienić nazwy folderu: {e}")

                # folder wyników = finalne pliki (dla przycisku "Otwórz folder
                # wyników" — także po zamknięciu i restarcie programu)
                _final = out_root / "PDF polaczone"
                self.last_output_dir = _final if _final.exists() else out_root
                _zap = getattr(self, "_zapamietaj_folder_wynikow", None)
                if _zap is not None:
                    _zap(self.last_output_dir)

            elif mode == "WORD":
                dir_01, dir_02 = out_root / "TXT", out_root / "Word"
                file_filter = self.get_selected_word_filters()
                filter_label = ", ".join(self.get_selected_word_filters())
                self.check_stop()
                self.update_status(
                    f"ETAP 1/2: Oczyszczanie plików TXT ({filter_label})", "#0078D7"
                )
                cw = self.task_clean_txt(in_root, dir_01, file_filter)
                if cw == 0:
                    self.log(
                        "\n[BŁĄD] Brak plików TXT do przetworzenia "
                        f"dla filtru: {filter_label}.")
                    self.update_status("Błąd — brak plików TXT", "#D83B01", animate=False)
                    return
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
                # UWAGA: domyślna wartość getattr jest tworzona ZAWSZE, nawet gdy
                # atrybut istnieje — ctk.BooleanVar bez okna Tk (web GUI) rzuca
                # RuntimeError "Too early to create variable", dlatego None + .get().
                _pmv = getattr(self, "pdf_merge_var", None)
                do_merge = True if _pmv is None else bool(_pmv.get())
                # przełącznik: czy dołączać 'Skróty i symbole' (checkbox w GUI)
                _skr = getattr(self, "pdf_skroty_var", None)
                add_skroty = True if _skr is None else bool(_skr.get())
                _total = 4 if (do_merge and add_skroty) else 3

                self.check_stop()
                (
                    self.update_status(
                        f"ETAP 1/{_total}: Zmiana formatu z Word na PDF", "#0078D7"
                    )
                    if do_merge
                    else self.update_status(
                        "Trwa zmiana formatu z Word na PDF...", "#0078D7"
                    )
                )
                self.task_convert_to_pdf(in_root, dir_03)
                self._flatten_001_subfolders(dir_03)
                self.set_progress(0.4 if do_merge else 1.0)
                if do_merge:
                    _e = 1
                    if add_skroty:
                        self.check_stop()
                        _e = 2
                        self.update_status(
                            "ETAP 2/4: Dołączanie 'Skrótów i symboli'", "#0078D7"
                        )
                        self._inject_skroty_step(dir_03)
                        self.set_progress(0.5)
                    self.check_stop()
                    self.update_status(
                        f"ETAP {_e + 1}/{_total}: Logiczna integracja dokumentacji",
                        "#0078D7",
                    )
                    self.task_merge_pdfs(dir_03, dir_04, mode_key="PDF")
                    self.set_progress(0.8)
                    self.check_stop()
                    self.update_status(
                        f"ETAP {_e + 2}/{_total}: Usuwanie anomalii", "#0078D7"
                    )
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

