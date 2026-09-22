"""
Forestly — Mixin: TabManualMergeMixin
"""

import customtkinter as ctk

from app.config import (
    add_tooltip,
)

class TabManualMergeMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_manual_merge_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)

        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)
        card = ctk.CTkFrame(
            scroll_frame,
            fg_color="#252526",
            corner_radius=8,
            border_width=1,
            border_color="#333333",
        )
        card.grid(row=0, column=0, padx=20, pady=(20, 15), sticky="new")
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            card, text="Wybierz folder PDF:", font=font_label, text_color="#E0E0E0"
        ).grid(row=0, column=0, padx=15, pady=(20, 10), sticky="w")
        self.manual_pdf_src = ctk.CTkEntry(
            card,
            placeholder_text="Wybierz lokalizację z plikami PDF...",
            height=36,
            border_width=1,
        )
        self.manual_pdf_src.grid(row=0, column=1, padx=5, pady=(20, 10), sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.manual_pdf_src),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(20, 10))
        ctk.CTkLabel(
            card, text="Wybierz folder docelowy:", font=font_label, text_color="#E0E0E0"
        ).grid(row=1, column=0, padx=15, pady=(0, 20), sticky="w")
        self.manual_pdf_dst = ctk.CTkEntry(
            card,
            placeholder_text="Gdzie zapisać plik wynikowy?",
            height=36,
            border_width=1,
        )
        self.manual_pdf_dst.grid(row=1, column=1, padx=5, pady=(0, 20), sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.manual_pdf_dst),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=1, column=2, padx=15, pady=(0, 20))
        btn = ctk.CTkButton(
            scroll_frame,
            text="Zarządzaj układem i scal pliki",
            image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            height=44,
            corner_radius=6,
            fg_color="#0067C0",
            hover_color="#005A9E",
            command=self.open_manual_merge_window,
        )
        btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")
        add_tooltip(
            btn,
            "Uruchamia interaktywne okno, w którym można poprzesuwać PDF-y góra/dół przed scaleniem.",
        )

