"""
Forestly — Widget: Ręczne scalanie PDF
=================================================
Zależności: config.py (COLORS, add_tooltip)
Odpowiada za: okno do ręcznego wczytania, ułożenia i scalenia plików PDF.
"""

import customtkinter as ctk
import tkinter as tk
from pathlib import Path
from pypdf import PdfWriter

from tkinter import messagebox

from app.config import COLORS, add_tooltip

class ManualPdfMergeWindow(ctk.CTkToplevel):
    def __init__(self, master, src_folder: Path, dst_folder: Path):
        super().__init__(master)
        self.title("Ręczne scalanie PDF")
        self.geometry("860x600")
        self.src_folder = src_folder
        self.dst_folder = dst_folder
        self.drag_index = None
        self.selected_index = None
        self.pdf_files = []
        self.build_ui()
        self.load_files()
        self.grab_set()

    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(
            self,
            text="Konfiguracja ręcznego scalania",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 8), sticky="w")
        center = ctk.CTkFrame(self)
        center.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="nsew")
        center.grid_columnconfigure(0, weight=1)
        center.grid_rowconfigure(0, weight=1)
        self.listbox = tk.Listbox(
            center,
            font=("Segoe UI", 12),
            activestyle="none",
            selectmode=tk.SINGLE,
            bg="#1E1E1E",
            fg="#E0E0E0",
            selectbackground="#005A9E",
            borderwidth=0,
            highlightthickness=0,
        )
        self.listbox.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self.listbox.bind("<<ListboxSelect>>", self.on_select)
        self.listbox.bind("<ButtonPress-1>", self.on_drag_start)
        self.listbox.bind("<B1-Motion>", self.on_drag_motion)
        self.listbox.bind("<ButtonRelease-1>", self.on_drag_drop)
        scrollbar = ctk.CTkScrollbar(center, command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=2, pady=2)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        controls = ctk.CTkFrame(center, fg_color="transparent")
        controls.grid(row=0, column=2, padx=15, pady=10, sticky="n")
        btn_kwargs = {
            "width": 140,
            "height": 32,
            "fg_color": "#333333",
            "hover_color": "#444444",
            "text_color": "#FFFFFF",
            "corner_radius": 4,
        }
        ctk.CTkButton(
            controls, text="↑ W górę", command=self.move_up, **btn_kwargs
        ).pack(pady=4)
        ctk.CTkButton(
            controls, text="↓ W dół", command=self.move_down, **btn_kwargs
        ).pack(pady=4)
        ctk.CTkButton(
            controls, text="Sortuj alfabetycznie", command=self.sort_alpha, **btn_kwargs
        ).pack(pady=(20, 4))
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="e")
        btn_merge = ctk.CTkButton(
            bottom,
            text="Zastosuj i Scal",
            command=self.merge_now,
            fg_color="#0067C0",
            hover_color="#005A9E",
            width=140,
            height=36,
            corner_radius=4,
        )
        btn_merge.pack(side="right")
        add_tooltip(
            btn_merge,
            "Natychmiast scala pliki w podanej wyżej kolejności do jednego dokumentu PDF.",
        )
        ctk.CTkButton(
            bottom,
            text="Anuluj",
            command=self.destroy,
            fg_color="transparent",
            border_width=1,
            border_color="#555555",
            hover_color="#333333",
            width=100,
            height=36,
            corner_radius=4,
        ).pack(side="right", padx=10)

    def load_files(self):
        self.pdf_files = sorted(
            [
                p
                for p in self.src_folder.iterdir()
                if p.is_file() and p.suffix.lower() == ".pdf"
            ],
            key=lambda p: p.name.lower(),
        )
        self.selected_index = 0 if self.pdf_files else None
        self.refresh_listbox()

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for idx, pdf in enumerate(self.pdf_files, 1):
            self.listbox.insert(tk.END, f"  {idx:02d}.  {pdf.name}")
        if self.selected_index is not None and 0 <= self.selected_index < len(
                self.pdf_files
        ):
            self.listbox.selection_set(self.selected_index)
            self.listbox.activate(self.selected_index)

    def on_select(self, _event=None):
        sel = self.listbox.curselection()
        self.selected_index = sel[0] if sel else None

    def on_drag_start(self, event):
        self.drag_index = self.listbox.nearest(event.y)
        self.selected_index = self.drag_index

    def on_drag_motion(self, event):
        idx = self.listbox.nearest(event.y)
        self.listbox.selection_clear(0, tk.END)
        if 0 <= idx < len(self.pdf_files):
            self.listbox.selection_set(idx)
            self.listbox.activate(idx)

    def on_drag_drop(self, event):
        if self.drag_index is None:
            return
        drop_index = self.listbox.nearest(event.y)
        if (
                0 <= self.drag_index < len(self.pdf_files)
                and 0 <= drop_index < len(self.pdf_files)
                and drop_index != self.drag_index
        ):
            item = self.pdf_files.pop(self.drag_index)
            self.pdf_files.insert(drop_index, item)
            self.selected_index = drop_index
            self.refresh_listbox()
            self.drag_index = None

    def move_up(self):
        if self.selected_index is not None and self.selected_index > 0:
            (
                self.pdf_files[self.selected_index - 1],
                self.pdf_files[self.selected_index],
            ) = (
                self.pdf_files[self.selected_index],
                self.pdf_files[self.selected_index - 1],
            )
            self.selected_index -= 1
            self.refresh_listbox()

    def move_down(self):
        if (
                self.selected_index is not None
                and self.selected_index < len(self.pdf_files) - 1
        ):
            (
                self.pdf_files[self.selected_index + 1],
                self.pdf_files[self.selected_index],
            ) = (
                self.pdf_files[self.selected_index],
                self.pdf_files[self.selected_index + 1],
            )
            self.selected_index += 1
            self.refresh_listbox()

    def sort_alpha(self):
        self.pdf_files.sort(key=lambda p: p.name.lower())
        self.selected_index = 0 if self.pdf_files else None
        self.refresh_listbox()

    def merge_now(self):
        if not self.pdf_files:
            messagebox.showwarning("Informacja", "Brak plików do scalenia.")
            return
        self.dst_folder.mkdir(parents=True, exist_ok=True)
        target = self.dst_folder / f"{self.src_folder.name}_scalony_recznie.pdf"
        writer = PdfWriter()
        try:
            for pdf in self.pdf_files:
                writer.append(str(pdf))
            with open(target, "wb") as f_out:
                writer.write(f_out)
            messagebox.showinfo(
                "Zakończono", f"Zapisano poprawnie plik:\n{target.name}"
            )
            self.destroy()
        except Exception as e:
            messagebox.showerror("Błąd systemowy", str(e))
        finally:
            writer.close()


