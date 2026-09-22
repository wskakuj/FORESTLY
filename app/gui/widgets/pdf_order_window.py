"""
Forestly — Widget: Okno kolejności PDF
=================================================
Zależności: config.py (PDF_ORDER_TEMPLATES, COLORS, load_order_store, save_order_store, get_saved_template_order, set_saved_template_order, get_default_template_keys, add_tooltip)
Odpowiada za: okno modalne do ustawiania kolejności stron PDF w dokumencie końcowym.
"""

import customtkinter as ctk
import tkinter as tk
from pathlib import Path

from app.config import (
    PDF_ORDER_TEMPLATES, COLORS, add_tooltip,
    load_order_store, save_order_store,
    get_saved_template_order, set_saved_template_order,
    get_default_template_keys,
)

class PdfOrderWindow(ctk.CTkToplevel):
    def __init__(self, master, target_folder: Path, mode_key: str):
        super().__init__(master)
        self.title("Konfiguracja kolejności PDF")
        self.geometry("860x600")
        self.target_folder = Path(target_folder)
        self.mode_key = mode_key
        self.drag_index = None
        self.selected_index = None
        self.items = []
        self.build_ui()
        self.load_items()
        self.grab_set()

    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(
            self,
            text=f"Układ kolejności dokumentów: {self.mode_key}",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")
        info = ctk.CTkFrame(self, fg_color="transparent")
        info.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        ctk.CTkLabel(
            info,
            text="Zmień układ, przeciągając pozycje myszką lub używając przycisków.",
            text_color="#A0A0A0",
        ).pack(anchor="w")
        center = ctk.CTkFrame(self)
        center.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="nsew")
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
            controls, text="Na początek", command=self.move_top, **btn_kwargs
        ).pack(pady=4)
        ctk.CTkButton(
            controls, text="Na koniec", command=self.move_bottom, **btn_kwargs
        ).pack(pady=4)
        ctk.CTkButton(
            controls,
            text="Zresetuj domyślne",
            command=self.reset_default,
            width=140,
            height=32,
            fg_color="#8B0000",
            hover_color="#A52A2A",
            corner_radius=4,
        ).pack(pady=(20, 4))
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="e")
        ctk.CTkButton(
            bottom,
            text="Zapisz konfigurację",
            command=self.save_and_close,
            fg_color="#0067C0",
            hover_color="#005A9E",
            width=140,
            height=36,
            corner_radius=4,
        ).pack(side="right")
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

    def load_items(self):
        saved_keys = get_saved_template_order(self.target_folder, self.mode_key)
        template_map = {t["key"]: t for t in PDF_ORDER_TEMPLATES}
        self.items = [template_map[key] for key in saved_keys if key in template_map]
        for tpl in PDF_ORDER_TEMPLATES:
            if tpl not in self.items:
                self.items.append(tpl)
        self.selected_index = 0 if self.items else None
        self.refresh_listbox()

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for idx, item in enumerate(self.items, 1):
            aliases = ", ".join(item["aliases"])
            self.listbox.insert(tk.END, f"  {idx:02d}. {item['label']}  ({aliases})")
        if self.selected_index is not None and 0 <= self.selected_index < len(
                self.items
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
        if 0 <= idx < len(self.items):
            self.listbox.selection_set(idx)
            self.listbox.activate(idx)

    def on_drag_drop(self, event):
        if self.drag_index is None:
            return
        drop_index = self.listbox.nearest(event.y)
        if (
                0 <= self.drag_index < len(self.items)
                and 0 <= drop_index < len(self.items)
                and drop_index != self.drag_index
        ):
            item = self.items.pop(self.drag_index)
            self.items.insert(drop_index, item)
            self.selected_index = drop_index
            self.refresh_listbox()
            self.drag_index = None

    def move_up(self):
        if self.selected_index is not None and self.selected_index > 0:
            self.items[self.selected_index - 1], self.items[self.selected_index] = (
                self.items[self.selected_index],
                self.items[self.selected_index - 1],
            )
            self.selected_index -= 1
            self.refresh_listbox()

    def move_down(self):
        if (
                self.selected_index is not None
                and self.selected_index < len(self.items) - 1
        ):
            self.items[self.selected_index + 1], self.items[self.selected_index] = (
                self.items[self.selected_index],
                self.items[self.selected_index + 1],
            )
            self.selected_index += 1
            self.refresh_listbox()

    def move_top(self):
        if self.selected_index is None:
            return
        item = self.items.pop(self.selected_index)
        self.items.insert(0, item)
        self.selected_index = 0
        self.refresh_listbox()

    def move_bottom(self):
        if self.selected_index is None:
            return
        item = self.items.pop(self.selected_index)
        self.items.append(item)
        self.selected_index = len(self.items) - 1
        self.refresh_listbox()

    def reset_default(self):
        self.items = PDF_ORDER_TEMPLATES.copy()
        self.selected_index = 0
        self.refresh_listbox()

    def save_and_close(self):
        keys = [item["key"] for item in self.items]
        set_saved_template_order(self.target_folder, self.mode_key, keys)
        self.destroy()


