"""
Forestly — Widget: Okno walidacji
===========================================
Zależności: config.py (COLORS)
Odpowiada za: okno modalne z ostrzeżeniami/walidacją przed uruchomieniem procesu.
"""

import customtkinter as ctk
import threading

from app.config import COLORS

class ValidationWindow(ctk.CTkToplevel):
    def __init__(self, master, title_text, warnings_list, proceed_event, cancel_event):
        super().__init__(master)
        self.title("Kontrola kompletności plików")
        self.geometry("700x500")

        # Wycentrowanie okna
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 700) // 2
        y = (screen_height - 500) // 2
        self.geometry(f"700x500+{x}+{y}")

        self.proceed_event = proceed_event
        self.cancel_event = cancel_event

        # Budowa UI
        lbl = ctk.CTkLabel(self, text=title_text, font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"), text_color="#D83B01")
        lbl.pack(pady=(20, 10), padx=20, anchor="w")

        # Pole z listą braków
        self.textbox = ctk.CTkTextbox(
            self, font=ctk.CTkFont(family="Consolas", size=12),
            fg_color="#1E1E1E", text_color="#E0E0E0", border_width=1, border_color="#333333"
        )
        self.textbox.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        for w in warnings_list:
            self.textbox.insert("end", w + "\n")
        self.textbox.configure(state="disabled")

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))

        btn_cancel = ctk.CTkButton(
            btn_frame, text="Przerwij operację", fg_color="#8B0000", hover_color="#A52A2A",
            command=self.do_cancel, font=ctk.CTkFont(weight="bold")
        )
        btn_cancel.pack(side="right", padx=(10, 0))

        btn_proceed = ctk.CTkButton(
            btn_frame, text="Ignoruj i kontynuuj", fg_color="#0067C0", hover_color="#005A9E",
            command=self.do_proceed, font=ctk.CTkFont(weight="bold")
        )
        btn_proceed.pack(side="right")

        # Zabezpieczenie przed zamknięciem 'X'
        self.protocol("WM_DELETE_WINDOW", self.do_cancel)

        # Wymuszenie okna na wierzchu i przejęcie interakcji
        self.grab_set()
        self.lift()
        try:
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(200, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def do_proceed(self):
        self.proceed_event.set()
        self.destroy()

    def do_cancel(self):
        self.cancel_event.set()
        self.destroy()

# ==========================================
# ROZLICZANIE POWIERZCHNI (XLS + VAL)
# ==========================================

