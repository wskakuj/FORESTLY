"""
Forestly — Modele danych
==================================
Zależności: config.py
Odpowiada za: struktury danych używane przez core i GUI.

Klasy dataclass zastępują krotki i słowniki w kodzie,
dając podpowiedzi typów w IDE i czytelniejszy kod.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class OrderStore:
    """Zapisana kolejność PDF dla danego folderu i trybu."""
    folder: Path
    data: dict = field(default_factory=dict)

    def get_order(self, mode_key: str) -> list[str]:
        saved = self.data.get(mode_key)
        if isinstance(saved, list) and saved:
            return saved
        from app.config import get_default_template_keys
        return get_default_template_keys()

    def set_order(self, mode_key: str, order_keys: list[str]):
        self.data[mode_key] = order_keys


@dataclass
class TerritoryEntry:
    """Pojedynczy wpis terytorialny (województwo/powiat/gmina)."""
    name: str
    children: dict[str, "TerritoryEntry"] = field(default_factory=dict)


@dataclass
class PipelineStep:
    """Jeden krok w dashboardzie pipeline'u."""
    title: str
    subtitle: str
    status: str = "pending"  # pending | running | done | error
    text: str = ""


@dataclass
class StreamFile:
    """Plik w kolejce live stream."""
    source: str
    target: str | None = None
    status: str = "queued"  # queued | running | completed
    start_time: float | None = None
    duration: float | None = None


@dataclass
class MarginConfig:
    """Konfiguracja marginesów dla danego typu pliku (w cm)."""
    top: float = 2.0
    bottom: float = 2.0
    left: float = 2.5
    right: float = 2.5

    def as_tuple(self) -> tuple[float, float, float, float]:
        return (self.top, self.bottom, self.left, self.right)


@dataclass
class ExcelSheetConfig:
    """Konfiguracja arkusza Excel (nazwa, czcionka, wiersze na stronę)."""
    name: str
    font_size: int = 9
    rows_per_page: int = 9
