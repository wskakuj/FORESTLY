"""
Forestly — Widget: Okno changelogu
=============================================
Zależności: config.py (COLORS)
Odpowiada za: okno wyświetlające listę zmian po aktualizacji programu.
"""

import customtkinter as ctk

import re

from app.config import COLORS

class ChangelogWindow(ctk.CTkToplevel):
    def __init__(self, master, version: str, changelog_text: str):
        super().__init__(master)
        self.title(f"Co nowego w wersji {version}?")
        self.geometry("600x450")
        self.resizable(False, False)

        # Wycentrowanie okna
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 600) // 2
        y = (screen_height - 450) // 2
        self.geometry(f"600x450+{x}+{y}")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Nagłówek
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(
            header_frame,
            text=f"Aplikacja została zaktualizowana do wersji {version}! 🎉",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color="#0078D7",
            wraplength=560,
            justify="left"
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_frame,
            text="Oto lista zmian i nowości wprowadzonych w tej wersji:",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#A0A0A0"
        ).pack(anchor="w", pady=(4, 0))

        # Treść opisu zmian
        self.textbox = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#1E1E1E",
            text_color="#E0E0E0",
            border_width=1,
            border_color="#333333",
            corner_radius=6,
            wrap="word",                # <--- KLUCZOWA ZMIANA: łamanie na całych słowach, nie w połowie wyrazu
            activate_scrollbars=True,
        )
        self.textbox.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="nsew")

        content = self.format_changelog_text(changelog_text)
        self.textbox.insert("0.0", content)
        self.textbox.configure(state="disabled")

        # Przycisk zamknięcia
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="e")

        ctk.CTkButton(
            bottom_frame,
            text="Zamknij",
            command=self.destroy,
            fg_color="#0067C0",
            hover_color="#005A9E",
            width=140,
            height=36,
            corner_radius=4,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        ).pack(side="right")

        # --- PRZYCIĄGANIE UWAGI / OTWIERANIE NA WIERZCHU ---
        self.grab_set()
        self.lift()
        try:
            self.focus_force()
        except Exception:
            pass
        try:
            self.attributes("-topmost", True)
            self.after(200, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    @staticmethod
    def format_changelog_text(text: str) -> str:
        """Czyści surowy opis z GitHuba (Markdown) do czytelnej postaci tekstowej."""
        if not text or not text.strip():
            return "Brak szczegółowego opisu zmian dla tej wersji."

        # Ujednolicenie końcówek linii (GitHub potrafi wysłać \r\n)
        text = text.replace("\r\n", "\n").replace("\r", "\n")

        lines = []
        for line in text.split("\n"):
            s = line.rstrip()

            # Nagłówki Markdown ("## Nowości") -> zwykły tekst
            s = re.sub(r"^\s*#{1,6}\s+", "", s)

            # Pogrubienia i podkreślenia Markdown
            s = s.replace("**", "").replace("__", "")

            # Kursywa (pojedyncze gwiazdki)
            s = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"\1", s)

            # Linki [tekst](adres) -> sam tekst
            s = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", s)

            # Punktory Markdown ("- ", "* ", "+ ") -> elegancka kropka
            s = re.sub(r"^\s*[-*+]\s+", "•  ", s)

            lines.append(s)

        # Sklej i zredukuj nadmiar pustych linii (maks. jedna z rzędu)
        cleaned = []
        prev_blank = False
        for line in lines:
            is_blank = not line.strip()
            if is_blank and prev_blank:
                continue
            cleaned.append(line)
            prev_blank = is_blank

        return "\n".join(cleaned).strip()

