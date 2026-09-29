# -*- coding: utf-8 -*-
"""Leniwe importy ciężkich bibliotek (v2.0.109) — szybszy start programu.

Dwa mechanizmy:

1) zainstaluj_leniwe([...]) — hak w meta-ścieżce Pythona: zwykłe
   `import pandas` / `import pandas as pd` / `import pymupdf as fitz`
   zwraca MODUŁ-LENIWCA, który ładuje się dopiero przy pierwszym
   użyciu atrybutu. Jeśli w czasie startu nic nie dotknie biblioteki —
   w ogóle nie wchodzi do pamięci.

2) leniwy(modul, nazwa) — obiekt-pośrednik dla importów z formy
   `from X import Y` na poziomie modułu (formuła FROM wymusza od razu
   załadowanie, więc hak tu nie pomoże). Prawdziwy Y ląduje w pamięci
   dopiero przy pierwszym wywołaniu, np.:

       Document = leniwy("docx", "Document")
       ...
       doc = Document("plik.docx")   # <- dopiero tu ładuje się python-docx

Oba mechanizmy są w 100% przezroczyste dla kodu wywołującego.
"""

import importlib
import sys


class _LeniwyZnajdzca:
    """Zwraca LazyLoader dla wskazanych modułów (zwykłe importy)."""

    def __init__(self, nazwy):
        self._nazwy = frozenset(nazwy)

    def find_spec(self, fullname, path=None, target=None):
        try:
            if fullname not in self._nazwy or fullname in sys.modules:
                return None
        except Exception:
            return None
        for znajdzca in list(sys.meta_path):
            if znajdzca is self:
                continue
            try:
                spec = znajdzca.find_spec(fullname, path, target)
            except Exception:
                spec = None
            if spec is not None and getattr(spec, "loader", None) is not None:
                try:
                    import importlib.util
                    spec.loader = importlib.util.LazyLoader(spec.loader)
                except Exception:
                    pass
                return spec
        return None


def zainstaluj_leniwe(moduly):
    """Włącza leniwe ładowanie zwykłych importów podanych modułów."""
    if not any(isinstance(z, _LeniwyZnajdzca) for z in sys.meta_path):
        sys.meta_path.insert(0, _LeniwyZnajdzca(moduly))


class _LeniwyModul:
    """Moduł-pośrednik: `import pandas as pd` zamienione na
    `pd = leniwy_modul("pandas")` — prawdziwy moduł ładuje się
    dopiero przy pierwszym `pd.<cokolwiek>`."""

    __slots__ = ("_nazwa",)

    def __init__(self, nazwa):
        object.__setattr__(self, "_nazwa", nazwa)

    def __getattr__(self, atrybut):
        m = importlib.import_module(object.__getattribute__(self, "_nazwa"))
        return getattr(m, atrybut)

    def __repr__(self):
        return f"<leniwy moduł: {object.__getattribute__(self, '_nazwa')}>"


def leniwy_modul(nazwa):
    """`pd = leniwy_modul("pandas")` — zamiast `import pandas as pd`
    na poziomie modułu (sam IMPORT z 'as' wymusza ładowanie,
    więc hak meta-ścieżki tu nie wystarcza)."""
    return _LeniwyModul(nazwa)


class _LeniwyObiekt:
    """Pośrednik dla `from modul import nazwa` — ładuje przy 1. użyciu."""

    __slots__ = ("_modul", "_nazwa")

    def __init__(self, modul, nazwa):
        object.__setattr__(self, "_modul", modul)
        object.__setattr__(self, "_nazwa", nazwa)

    def _rozwi(self):
        m = importlib.import_module(object.__getattribute__(self, "_modul"))
        return getattr(m, object.__getattribute__(self, "_nazwa"))

    def __getattr__(self, atrybut):
        return getattr(self._rozwi(), atrybut)

    def __call__(self, *a, **kw):
        return self._rozwi()(*a, **kw)

    def __repr__(self):
        return (f"<leniwy: {object.__getattribute__(self, '_modul')}."
                f"{object.__getattribute__(self, '_nazwa')}>")


def leniwy(modul, nazwa):
    """`Font = leniwy("openpyxl.styles", "Font")` — zamiast
    `from openpyxl.styles import Font` na poziomie modułu."""
    return _LeniwyObiekt(modul, nazwa)
