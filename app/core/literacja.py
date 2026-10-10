# -*- coding: utf-8 -*-
"""Literacja wydzieleń w plikach GEO-MAP (rdzeń algorytmu).

Przeniesione z programu Literacja_UPUL v1.3 (autorski program użytkownika),
z którego usunięto wyłącznie część okienkową (tkinter). Algorytm bez zmian:
  * warstwa 5310,
  * numer oddziału z początku A1,
  * literacja ciągła w obrębie kompleksu leśnego,
  * kolejność kompleksów: północny wschód -> południowy zachód,
  * domyślna tolerancja łączenia kompleksów 1,0 m.

Wejście: plik .MAP / .map / .~AP albo ZIP. Wyjście: ZIP z poprawionymi
plikami, kopia oryginału, raport CSV/TXT oraz podgląd SVG.
"""

# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import hashlib
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional


APP_NAME = "Literacja_UPUL"
APP_VERSION = "1.3"
DEFAULT_LAYER = "5310"
DEFAULT_COMPLEX_TOLERANCE = 1.0
ENCODING = "cp1250"
ALLOWED_BASE = ["a", "b", "c", "d", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "w", "x", "y", "z"]

ATTR_RE = re.compile(r"^:([A-Za-z0-9_]+)\[(.*)\]\s*$")
OBJECT_RE = re.compile(r"^\*(\d+)\b")
POINT_RE = re.compile(r"^P\s+\d+\s+([-+]?\d+(?:\.\d+)?)\s+([-+]?\d+(?:\.\d+)?)")
ODDZIAL_RE = re.compile(r"^\s*(\d+)")


class LiteracjaError(RuntimeError):
    pass


@dataclass
class GeoObject:
    start: int
    end: int
    layer: str
    lines: list[str]
    attrs: dict[str, str]
    points: list[tuple[float, float]]
    oddzial: Optional[str]
    tx: str
    area: float
    cx: float
    cy: float
    new_a1: str = ""
    sequence_no: int = 0
    complex_no: int = 0


def decode_map(data: bytes) -> str:
    for enc in ("cp1250", "latin1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    raise LiteracjaError("Nie udało się odczytać pliku MAP w kodowaniu CP1250.")


def natural_key(value: str) -> list[object]:
    value = (value or "").strip().lower()
    return [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", value)]


def polygon_area_centroid(points: list[tuple[float, float]]) -> tuple[float, float, float]:
    if not points:
        return 0.0, 0.0, 0.0
    pts = points[:]
    if len(pts) > 1 and pts[0] != pts[-1]:
        pts.append(pts[0])
    if len(pts) < 4:
        x = sum(p[0] for p in points) / len(points)
        y = sum(p[1] for p in points) / len(points)
        return 0.0, x, y
    twice_area = 0.0
    cx_num = 0.0
    cy_num = 0.0
    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        cross = x1 * y2 - x2 * y1
        twice_area += cross
        cx_num += (x1 + x2) * cross
        cy_num += (y1 + y2) * cross
    if abs(twice_area) < 1e-9:
        x = sum(p[0] for p in points) / len(points)
        y = sum(p[1] for p in points) / len(points)
        return 0.0, x, y
    signed_area = twice_area / 2.0
    cx = cx_num / (3.0 * twice_area)
    cy = cy_num / (3.0 * twice_area)
    return abs(signed_area), cx, cy


def suffixes_xyz() -> Iterable[str]:
    alphabet = "xyz"
    length = 1
    while True:
        total = len(alphabet) ** length
        for n in range(total):
            chars = []
            v = n
            for _ in range(length):
                chars.append(alphabet[v % len(alphabet)])
                v //= len(alphabet)
            yield "".join(reversed(chars))
        length += 1


def forestry_letters() -> Iterable[str]:
    for letter in ALLOWED_BASE:
        yield letter
    for suffix in suffixes_xyz():
        for letter in ALLOWED_BASE:
            yield letter + suffix


def parse_objects(lines: list[str]) -> list[GeoObject]:
    starts = [i for i, line in enumerate(lines) if OBJECT_RE.match(line)]
    objects: list[GeoObject] = []
    for pos, start in enumerate(starts):
        end = starts[pos + 1] if pos + 1 < len(starts) else len(lines)
        block = lines[start:end]
        m = OBJECT_RE.match(block[0])
        if not m:
            continue
        attrs: dict[str, str] = {}
        points: list[tuple[float, float]] = []
        for line in block[1:]:
            am = ATTR_RE.match(line)
            if am:
                attrs[am.group(1).upper()] = am.group(2)
            pm = POINT_RE.match(line)
            if pm:
                points.append((float(pm.group(1)), float(pm.group(2))))
        a1 = attrs.get("A1", "")
        om = ODDZIAL_RE.match(a1)
        oddzial = om.group(1) if om else None
        if not oddzial:
            # A1 bez numeru oddziału (np. "ax") — bierzemy go z A6 (np. "2ax")
            om6 = ODDZIAL_RE.match(attrs.get("A6", ""))
            if om6:
                oddzial = om6.group(1)
        area, cx, cy = polygon_area_centroid(points)
        objects.append(
            GeoObject(
                start=start,
                end=end,
                layer=m.group(1),
                lines=block,
                attrs=attrs,
                points=points,
                oddzial=oddzial,
                tx=attrs.get("TX", "").strip(),
                area=area,
                cx=cx,
                cy=cy,
            )
        )
    return objects


def spatial_order(group: list[GeoObject]) -> list[GeoObject]:
    """Zwraca kolejność od północnego wschodu do południowego zachodu.

    Położenie X i Y jest normalizowane osobno w każdym oddziale, dzięki czemu
    wschód i północ mają równy wpływ na kolejność. Większa suma znormalizowanych
    współrzędnych oznacza położenie bliższe północnemu wschodowi.
    """
    if not group:
        return []
    min_x = min(obj.cx for obj in group)
    max_x = max(obj.cx for obj in group)
    min_y = min(obj.cy for obj in group)
    max_y = max(obj.cy for obj in group)
    span_x = max(max_x - min_x, 1.0)
    span_y = max(max_y - min_y, 1.0)

    def key(obj: GeoObject) -> tuple:
        east = (obj.cx - min_x) / span_x
        north = (obj.cy - min_y) / span_y
        return (
            -(east + north),  # północny wschód -> południowy zachód
            -east,
            -north,
            obj.start,
        )

    return sorted(group, key=key)



def _segments(points: list[tuple[float, float]]) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    if len(points) < 2:
        return []
    pts = points[:] if points[0] == points[-1] else points + [points[0]]
    return list(zip(pts, pts[1:]))


def _sub(a: tuple[float, float], b: tuple[float, float]) -> tuple[float, float]:
    return a[0] - b[0], a[1] - b[1]


def _dot(a: tuple[float, float], b: tuple[float, float]) -> float:
    return a[0] * b[0] + a[1] * b[1]


def _cross(a: tuple[float, float], b: tuple[float, float]) -> float:
    return a[0] * b[1] - a[1] * b[0]


def _orientation(a: tuple[float, float], b: tuple[float, float], c: tuple[float, float]) -> float:
    return _cross(_sub(b, a), _sub(c, a))


def _on_segment(a: tuple[float, float], b: tuple[float, float], p: tuple[float, float], eps: float = 1e-8) -> bool:
    return (
        abs(_orientation(a, b, p)) <= eps
        and min(a[0], b[0]) - eps <= p[0] <= max(a[0], b[0]) + eps
        and min(a[1], b[1]) - eps <= p[1] <= max(a[1], b[1]) + eps
    )


def _segments_intersect(
    a: tuple[float, float], b: tuple[float, float],
    c: tuple[float, float], d: tuple[float, float],
) -> bool:
    o1 = _orientation(a, b, c)
    o2 = _orientation(a, b, d)
    o3 = _orientation(c, d, a)
    o4 = _orientation(c, d, b)
    if ((o1 > 0 > o2) or (o2 > 0 > o1)) and ((o3 > 0 > o4) or (o4 > 0 > o3)):
        return True
    return _on_segment(a, b, c) or _on_segment(a, b, d) or _on_segment(c, d, a) or _on_segment(c, d, b)


def _point_segment_distance(
    p: tuple[float, float], a: tuple[float, float], b: tuple[float, float]
) -> float:
    ab = _sub(b, a)
    ap = _sub(p, a)
    den = _dot(ab, ab)
    if den <= 1e-18:
        return math.dist(p, a)
    t = max(0.0, min(1.0, _dot(ap, ab) / den))
    q = (a[0] + t * ab[0], a[1] + t * ab[1])
    return math.dist(p, q)


def _segment_distance(
    a: tuple[float, float], b: tuple[float, float],
    c: tuple[float, float], d: tuple[float, float],
) -> float:
    if _segments_intersect(a, b, c, d):
        return 0.0
    return min(
        _point_segment_distance(a, c, d),
        _point_segment_distance(b, c, d),
        _point_segment_distance(c, a, b),
        _point_segment_distance(d, a, b),
    )


def _point_in_polygon(point: tuple[float, float], polygon: list[tuple[float, float]]) -> bool:
    if len(polygon) < 3:
        return False
    x, y = point
    inside = False
    pts = polygon[:] if polygon[0] == polygon[-1] else polygon + [polygon[0]]
    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        if _on_segment((x1, y1), (x2, y2), point):
            return True
        if (y1 > y) != (y2 > y):
            x_cross = (x2 - x1) * (y - y1) / (y2 - y1) + x1
            if x < x_cross:
                inside = not inside
    return inside


def _bbox(obj: GeoObject) -> tuple[float, float, float, float]:
    if not obj.points:
        return obj.cx, obj.cy, obj.cx, obj.cy
    xs = [p[0] for p in obj.points]
    ys = [p[1] for p in obj.points]
    return min(xs), min(ys), max(xs), max(ys)


def polygons_connected(a: GeoObject, b: GeoObject, tolerance: float) -> bool:
    """Czy dwa wydzielenia należą do jednego kompleksu przestrzennego."""
    aminx, aminy, amaxx, amaxy = _bbox(a)
    bminx, bminy, bmaxx, bmaxy = _bbox(b)
    if amaxx + tolerance < bminx or bmaxx + tolerance < aminx:
        return False
    if amaxy + tolerance < bminy or bmaxy + tolerance < aminy:
        return False

    seg_a = _segments(a.points)
    seg_b = _segments(b.points)
    for p1, p2 in seg_a:
        for q1, q2 in seg_b:
            if _segment_distance(p1, p2, q1, q2) <= tolerance:
                return True

    # Obsługa nakładania / enklawy bez przecięcia samych krawędzi.
    if a.points and _point_in_polygon(a.points[0], b.points):
        return True
    if b.points and _point_in_polygon(b.points[0], a.points):
        return True
    return False


def order_by_forest_complexes(
    group: list[GeoObject], tolerance: float
) -> tuple[list[GeoObject], list[list[GeoObject]]]:
    """Najpierw kończy kompleks, potem przechodzi do kolejnego NE -> SW."""
    if not group:
        return [], []

    adjacency: dict[int, set[int]] = {i: set() for i in range(len(group))}
    for i in range(len(group)):
        for j in range(i + 1, len(group)):
            if polygons_connected(group[i], group[j], tolerance):
                adjacency[i].add(j)
                adjacency[j].add(i)

    components_idx: list[list[int]] = []
    seen: set[int] = set()
    for start in range(len(group)):
        if start in seen:
            continue
        stack = [start]
        seen.add(start)
        component: list[int] = []
        while stack:
            current = stack.pop()
            component.append(current)
            for nxt in adjacency[current]:
                if nxt not in seen:
                    seen.add(nxt)
                    stack.append(nxt)
        components_idx.append(component)

    components = [[group[i] for i in comp] for comp in components_idx]
    min_x = min(sum(o.cx for o in comp) / len(comp) for comp in components)
    max_x = max(sum(o.cx for o in comp) / len(comp) for comp in components)
    min_y = min(sum(o.cy for o in comp) / len(comp) for comp in components)
    max_y = max(sum(o.cy for o in comp) / len(comp) for comp in components)
    span_x = max(max_x - min_x, 1.0)
    span_y = max(max_y - min_y, 1.0)

    def component_key(comp: list[GeoObject]) -> tuple[float, float, float]:
        cx = sum(o.cx for o in comp) / len(comp)
        cy = sum(o.cy for o in comp) / len(comp)
        east = (cx - min_x) / span_x
        north = (cy - min_y) / span_y
        return (-(east + north), -east, -north)

    components.sort(key=component_key)
    flattened: list[GeoObject] = []

    for complex_no, comp in enumerate(components, start=1):
        rank = {obj.start: pos for pos, obj in enumerate(spatial_order(comp))}
        comp_starts = {obj.start for obj in comp}
        idx_by_start = {obj.start: group.index(obj) for obj in comp}
        remaining = set(comp_starts)
        ordered: list[GeoObject] = []

        # Zaczynamy od wydzielenia najbardziej na północnym wschodzie.
        current_start = min(remaining, key=lambda st: rank[st])
        ordered.append(next(o for o in comp if o.start == current_start))
        remaining.remove(current_start)
        visited_starts = {current_start}

        while remaining:
            # Preferujemy obiekt stykający się z już przejrzaną częścią kompleksu.
            frontier = []
            for st in remaining:
                gi = idx_by_start[st]
                if any(group.index(next(o for o in comp if o.start == vst)) in adjacency[gi] for vst in visited_starts):
                    frontier.append(st)
            candidates = frontier if frontier else list(remaining)
            next_start = min(candidates, key=lambda st: rank[st])
            ordered.append(next(o for o in comp if o.start == next_start))
            remaining.remove(next_start)
            visited_starts.add(next_start)

        for obj in ordered:
            obj.complex_no = complex_no
        flattened.extend(ordered)

    return flattened, components


def replace_or_insert_a1(block: list[str], new_a1: str) -> list[str]:
    result = block[:]
    for i, line in enumerate(result):
        if line.upper().startswith(":A1["):
            result[i] = f":A1[{new_a1}]"
            return result
    insert_at = 1
    for i, line in enumerate(result[1:], start=1):
        if line.upper().startswith(":ID["):
            insert_at = i + 1
            break
    result.insert(insert_at, f":A1[{new_a1}]")
    return result


def update_checksum(text: str) -> tuple[str, str]:
    # GEO-MAP: CHK to MD5 wszystkich bajtów po pierwszej linii.
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    if not lines:
        raise LiteracjaError("Pusty plik MAP.")
    body = "\r\n".join(lines[1:]).encode(ENCODING, errors="replace")
    checksum = hashlib.md5(body).hexdigest()
    if "CHK=[" in lines[0]:
        lines[0] = re.sub(r"CHK=\[[^\]]*\]", f"CHK=[{checksum}]", lines[0], count=1)
    else:
        lines[0] = lines[0].rstrip() + f" CHK=[{checksum}]"
    return "\r\n".join(lines), checksum


def _napraw_nazwe_w_zipie(info: "zipfile.ZipInfo") -> str:
    """Odtwarza polskie znaki w nazwie pliku z ZIP-a.

    Windows często pakuje pliki BEZ flagi UTF-8 — Python czyta wtedy nazwy
    jako cp437 i „Ł” wychodzi jako „ú”. Bierzemy oryginalne bajty i próbujemy
    UTF-8, a gdy się nie uda — cp1250 (kodowanie GEO-MAP / Windows PL).
    """
    nazwa = info.filename
    if info.flag_bits & 0x800:            # ZIP sam mówi, że nazwa jest UTF-8
        return nazwa
    try:
        raw = nazwa.encode("cp437")
    except UnicodeEncodeError:
        return nazwa
    for kod in ("utf-8", "cp1250"):
        try:
            return raw.decode(kod)
        except UnicodeDecodeError:
            continue
    return nazwa


def safe_extract_zip(zip_path: Path, destination: Path) -> None:
    """Rozpakowuje ZIP bezpiecznie (bez wyjścia poza folder docelowy).

    Rozpakowujemy ręcznie, pozycja po pozycji, żeby móc poprawić nazwy plików
    z polskimi znakami (Windows zapisuje je czasem bez flagi UTF-8).
    """
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    baza = destination.resolve()
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            nazwa = _napraw_nazwe_w_zipie(info)
            cel = destination / nazwa
            try:
                cel.resolve().relative_to(baza)
            except ValueError as exc:
                raise LiteracjaError("ZIP zawiera niedozwoloną ścieżkę.") from exc
            if info.is_dir():
                cel.mkdir(parents=True, exist_ok=True)
                continue
            cel.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(cel, "wb") as out:
                shutil.copyfileobj(src, out)


def find_main_map(root: Path) -> Path:
    maps = [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() == ".map"]
    if not maps:
        raise LiteracjaError("Nie znaleziono pliku .MAP.")
    if len(maps) == 1:
        return maps[0]
    maps.sort(key=lambda p: p.stat().st_size, reverse=True)
    return maps[0]


def prepare_input(input_path: Path, workdir: Path) -> tuple[Path, Path]:
    source_root = workdir / "source"
    source_root.mkdir(parents=True, exist_ok=True)
    if input_path.suffix.lower() == ".zip":
        safe_extract_zip(input_path, source_root)
        main_map = find_main_map(source_root)
        return main_map.parent, main_map
    if input_path.is_file() and input_path.suffix.lower() in {".map", ".~ap"}:
        stem = input_path.stem
        # KWIATKOWO.~AP ma stem "KWIATKOWO".
        for sibling in input_path.parent.iterdir():
            if not sibling.is_file():
                continue
            if sibling.stem.lower() == stem.lower() and sibling.suffix.lower() in {".map", ".~ap", ".ini"}:
                shutil.copy2(sibling, source_root / sibling.name)
        main_map = find_main_map(source_root)
        return source_root, main_map
    raise LiteracjaError("Wskaż plik ZIP, MAP albo ~AP.")


def copy_tree_contents(source: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    for item in source.iterdir():
        target = destination / item.name
        if item.is_dir():
            shutil.copytree(item, target, dirs_exist_ok=True)
        else:
            shutil.copy2(item, target)


def create_svg(objects: list[GeoObject], output_path: Path) -> None:
    polys = [o for o in objects if len(o.points) >= 3]
    if not polys:
        output_path.write_text("<svg xmlns='http://www.w3.org/2000/svg' width='900' height='300'><text x='20' y='40'>Brak geometrii do podglądu.</text></svg>", encoding="utf-8")
        return
    all_points = [p for o in polys for p in o.points]
    minx = min(x for x, _ in all_points)
    maxx = max(x for x, _ in all_points)
    miny = min(y for _, y in all_points)
    maxy = max(y for _, y in all_points)
    width, height, margin = 1600, 1000, 35
    spanx = max(maxx - minx, 1.0)
    spany = max(maxy - miny, 1.0)
    scale = min((width - 2 * margin) / spanx, (height - 2 * margin) / spany)

    def tr(x: float, y: float) -> tuple[float, float]:
        sx = margin + (x - minx) * scale
        sy = height - margin - (y - miny) * scale
        return sx, sy

    parts = [
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}' viewBox='0 0 {width} {height}'>",
        "<rect width='100%' height='100%' fill='white'/>",
        "<style>text{font-family:Arial,sans-serif;font-size:11px;font-weight:bold;paint-order:stroke;stroke:white;stroke-width:3px;stroke-linejoin:round}.poly{fill:#e8f2ec;fill-opacity:.55;stroke:#254f3c;stroke-width:1.1}</style>",
    ]
    for obj in polys:
        pts = " ".join(f"{x:.2f},{y:.2f}" for x, y in (tr(x, y) for x, y in obj.points))
        parts.append(f"<polygon class='poly' points='{pts}'/>")
    for obj in polys:
        x, y = tr(obj.cx, obj.cy)
        parts.append(f"<text x='{x:.2f}' y='{y:.2f}' text-anchor='middle' dominant-baseline='central'>{obj.new_a1}</text>")
    parts.append("</svg>")
    output_path.write_text("\n".join(parts), encoding="utf-8")



def literate_plik(
    src: Path,
    dst: Path,
    layer: str = DEFAULT_LAYER,
    overwrite_existing: bool = True,
    complex_tolerance: float = DEFAULT_COMPLEX_TOLERANCE,
    log=None,
) -> dict:
    """Literuje POJEDYNCZY plik .MAP i zapisuje wynik pod ścieżką ``dst``.

    To wersja ``literate_map`` bez pakowania do ZIP — potrzebna w „Pełnym
    automacie”, gdzie zaraz po literacji ta sama mapa idzie do wpisania
    opisów i ułożenia. Zwraca słownik ze statystykami (jak ``literate_map``).
    """
    def _log(msg):
        if log is not None:
            log(msg)

    src = Path(src)
    dst = Path(dst)
    if not src.exists():
        raise LiteracjaError(f"Plik wejściowy nie istnieje: {src}")

    raw = src.read_bytes()
    text = decode_map(raw)
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    objects = parse_objects(lines)
    selected = [o for o in objects if o.layer == layer]
    if not selected:
        raise LiteracjaError(f"Nie znaleziono obiektów na warstwie {layer}.")
    without_oddzial = [o for o in selected if not o.oddzial]
    if without_oddzial:
        raise LiteracjaError(
            f"Na warstwie {layer} znaleziono {len(without_oddzial)} obiektów "
            "bez numeru oddziału (A1 i A6 puste)."
        )

    eligible: list[GeoObject] = []
    skipped_existing: list[GeoObject] = []
    for obj in selected:
        current_a1 = obj.attrs.get("A1", "").strip()
        has_letter = bool(re.search(r"[A-Za-z]", current_a1))
        if has_letter and not overwrite_existing:
            skipped_existing.append(obj)
        else:
            eligible.append(obj)
    if not eligible:
        raise LiteracjaError("Nie ma obiektów do zaliterowania przy wybranych ustawieniach.")

    groups: dict[str, list[GeoObject]] = {}
    for obj in eligible:
        groups.setdefault(obj.oddzial or "", []).append(obj)

    complex_counts: dict[str, int] = {}
    for oddzial, group in groups.items():
        ordered, complexes = order_by_forest_complexes(group, complex_tolerance)
        group[:] = ordered
        complex_counts[oddzial] = len(complexes)
        letters = forestry_letters()
        for idx, obj in enumerate(group, start=1):
            obj.sequence_no = idx
            obj.new_a1 = f"{oddzial}{next(letters)}"

    for obj in sorted(eligible, key=lambda o: o.start, reverse=True):
        lines[obj.start:obj.end] = replace_or_insert_a1(lines[obj.start:obj.end], obj.new_a1)

    output_text, checksum = update_checksum("\n".join(lines))
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(output_text.encode(ENCODING, errors="replace"))

    # Kontrola zapisu — tak jak w literate_map.
    saved_text = decode_map(dst.read_bytes()).replace("\r\n", "\n").replace("\r", "\n")
    saved_objects = [o for o in parse_objects(saved_text.split("\n")) if o.layer == layer]
    saved_a1 = {o.attrs.get("ID", ""): o.attrs.get("A1", "") for o in saved_objects}
    for obj in eligible:
        obj_id = obj.attrs.get("ID", "")
        if obj_id and saved_a1.get(obj_id) != obj.new_a1:
            raise LiteracjaError(f"Kontrola zapisu nie powiodła się dla obiektu {obj_id}.")

    _log(f"Zaliterowano {len(eligible)} wydzieleń w {len(groups)} oddziałach "
         f"(kompleksy: {complex_counts}).")
    return {
        "changed": len(eligible),
        "skipped": len(skipped_existing),
        "groups": {k: len(v) for k, v in sorted(groups.items(), key=lambda kv: int(kv[0]))},
        "complexes": complex_counts,
        "checksum": checksum,
        "sciezka": str(dst),
        "obiekty": sorted(eligible, key=lambda o: (int(o.oddzial or 0), o.sequence_no)),
    }

# ======================================================================
# ZBIERANIE PLIKÓW WEJŚCIOWYCH  (jeden plik / kilka plików / folder / ZIP)
# ======================================================================

def zbierz_pliki_wejsciowe(wpis, katalog_roboczy=None, log=None) -> tuple[list[Path], list[str]]:
    """Zamienia wpis użytkownika na listę plików .MAP do zaliterowania.

    Wpis może być:
      * jednym plikiem .MAP / .~AP,
      * kilkoma plikami rozdzielonymi średnikiem (tak zwraca wybór wielokrotny),
      * folderem — bierze wszystkie .MAP w nim i w podfolderach,
      * plikiem .ZIP — rozpakowuje i bierze wszystkie .MAP ze środka.

    Zwraca ``(lista_plików, ostrzeżenia)``.
    """
    def _log(msg):
        if log is not None:
            log(msg)

    wpis = str(wpis or "").strip()
    if not wpis:
        return [], ["Nie wskazano plików wejściowych."]

    czesci = [c.strip().strip('"') for c in re.split(r"[;\n]+", wpis) if c.strip()]
    pliki: list[Path] = []
    ostrzezenia: list[str] = []
    widziane: set[str] = set()

    def _dodaj(q: Path) -> None:
        klucz = str(q.resolve()).lower()
        if klucz not in widziane:
            widziane.add(klucz)
            pliki.append(q)

    for c in czesci:
        p = Path(c)
        if p.is_dir():
            znalezione = sorted(
                q for q in p.rglob("*")
                if q.is_file() and q.suffix.lower() in {".map", ".~ap"}
            )
            if not znalezione:
                ostrzezenia.append(f"W folderze „{p}” nie ma plików .MAP.")
            for q in znalezione:
                _dodaj(q)
        elif p.is_file() and p.suffix.lower() == ".zip":
            cel = Path(katalog_roboczy) if katalog_roboczy else (p.parent / "_ROZPAKOWANE")
            cel = cel / p.stem
            try:
                cel.mkdir(parents=True, exist_ok=True)
                safe_extract_zip(p, cel)
            except Exception as exc:                          # noqa: BLE001
                ostrzezenia.append(f"Nie udało się rozpakować „{p.name}”: {exc}")
                continue
            znalezione = sorted(
                q for q in cel.rglob("*")
                if q.is_file() and q.suffix.lower() in {".map", ".~ap"}
            )
            if not znalezione:
                ostrzezenia.append(f"W archiwum „{p.name}” nie ma plików .MAP.")
            for q in znalezione:
                _dodaj(q)
        elif p.is_file() and p.suffix.lower() in {".map", ".~ap"}:
            _dodaj(p)
        else:
            ostrzezenia.append(f"Pomijam „{c}” — nie istnieje albo to nie jest .MAP/.ZIP/folder.")

    return pliki, ostrzezenia


# ======================================================================
# LITERACJA WIELU PLIKÓW DO JEDNEGO FOLDERU  (wynik bez ZIP)
# ======================================================================

def literate_do_folder(
    pliki,
    out_dir: Path,
    layer: str = DEFAULT_LAYER,
    overwrite_existing: bool = True,
    complex_tolerance: float = DEFAULT_COMPLEX_TOLERANCE,
    log=None,
) -> dict:
    """Literuje listę plików .MAP i zapisuje wynik w FOLDERZE (bez ZIP-a).

    W folderze wynikowym powstają:
      * ``<NAZWA>_zaliterowane.MAP`` — po jednym na każdą mapę wejściową,
      * ``KOPIA_ORYGINALNA/``        — nietknięte kopie map źródłowych,
      * ``RAPORT_LITERACJI.csv`` i ``RAPORT_LITERACJI.txt`` — zbiorcze,
      * ``PODGLAD_LITERACJI.svg``    — podgląd wszystkich zaliterowanych map.

    Mapa, której nie da się zaliterować, jest POMIJANA, a powód trafia do
    raportu (nie przerywa pracy nad pozostałymi mapami).
    """
    def _log(msg):
        if log is not None:
            log(msg)

    baza = Path(out_dir)
    baza.mkdir(parents=True, exist_ok=True)

    # NOWY podfolder na każdy przebieg — żeby wyniki się nie mieszały
    # ani nie nadpisywały między kolejnymi uruchomieniami.
    stempel = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = baza / f"ZALITEROWANE_{stempel}"
    _nr = 2
    while out_dir.exists():
        out_dir = baza / f"ZALITEROWANE_{stempel}_{_nr}"
        _nr += 1
    out_dir.mkdir(parents=True, exist_ok=True)

    kopia_dir = out_dir / "KOPIA_ORYGINALNA"
    kopia_dir.mkdir(parents=True, exist_ok=True)

    wyniki: list[dict] = []
    bledy: list[tuple[str, str]] = []
    wszystkie_obiekty: list[GeoObject] = []
    wiersze_csv: list[list] = []

    for src in pliki:
        src = Path(src)
        try:
            dst = out_dir / f"{src.stem}_zaliterowane.MAP"
            st = literate_plik(src, dst, layer, overwrite_existing,
                               complex_tolerance, _log)
            try:
                shutil.copy2(src, kopia_dir / src.name)
            except OSError:
                pass
            objs = st.get("obiekty") or []
            wszystkie_obiekty += objs
            for obj in objs:
                litera = obj.new_a1[len(obj.oddzial or ""):]
                wiersze_csv.append([
                    src.name, obj.oddzial, obj.complex_no, obj.sequence_no,
                    obj.tx, obj.new_a1, litera, round(obj.area, 2),
                    round(obj.cx, 3), round(obj.cy, 3), obj.start + 1,
                ])
            wyniki.append({
                "mapa": src.name, "wynik": dst.name,
                "changed": st.get("changed", 0), "skipped": st.get("skipped", 0),
                "groups": st.get("groups", {}), "complexes": st.get("complexes", {}),
                "checksum": st.get("checksum", ""),
            })
        except Exception as exc:                              # noqa: BLE001
            bledy.append((src.name, str(exc)))
            _log(f"POMINIĘTO {src.name} — {exc}")

    # --- CSV -----------------------------------------------------------
    csv_path = out_dir / "RAPORT_LITERACJI.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(["mapa", "oddzial", "kompleks", "kolejnosc_w_oddziale",
                         "TX_numer_dzialki", "nowy_A1", "litera", "powierzchnia_m2",
                         "X", "Y", "linia_w_pliku"])
        writer.writerows(wiersze_csv)

    # --- TXT -----------------------------------------------------------
    txt_path = out_dir / "RAPORT_LITERACJI.txt"
    linie = [
        f"RAPORT LITERACJI WARSTWY {layer}",
        f"Program: {APP_NAME} {APP_VERSION}",
        f"Data: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        f"Zaliterowanych map: {len(wyniki)} z {len(pliki)}",
        "Sposób kolejności: najpierw cały kompleks leśny, potem kolejny kompleks "
        "od północnego wschodu do południowego zachodu. "
        f"Tolerancja łączenia kompleksu: {complex_tolerance:g} m.",
        "Zapis A1: pełny adres oddział + litera, np. 1a, 2ax.",
        "Ciąg liter: a,b,c,d,f,g,h,i,j,k,l,m,n,o,p,r,s,t,w,x,y,z, potem ax,bx,...",
        "",
    ]
    for w in wyniki:
        grupy = ", ".join(f"{k}: {v}" for k, v in sorted(w["groups"].items(), key=lambda kv: int(kv[0])))
        linie.append(f"  • {w['mapa']:<30} zmieniono: {w['changed']:4d}, "
                     f"pominięto istniejących: {w['skipped']:3d}, oddziały: {grupy or '—'} "
                     f"-> {w['wynik']}")
    linie += ["", "-" * 70, "MAPY POMINIĘTE (nie udało się zaliterować):", "-" * 70]
    if bledy:
        for nazwa, powod in bledy:
            linie.append(f"  x {nazwa}")
            linie.append(f"      powód: {powod}")
    else:
        linie.append("  (żadnej mapy nie pominięto)")
    linie.append("")
    txt_path.write_text("\n".join(linie), encoding="utf-8")

    # --- SVG -----------------------------------------------------------
    create_svg(wszystkie_obiekty, out_dir / "PODGLAD_LITERACJI.svg")

    razem = sum(w["changed"] for w in wyniki)
    _log(f"Utworzono nowy folder: {out_dir}")
    _log(f"Zapisano {len(wyniki)} map(y), wydzieleń: {razem}.")
    return {
        "folder": str(out_dir),
        "folder_bazowy": str(baza),
        "mapy": len(wyniki),
        "wydzielen": razem,
        "pominiete": bledy,
        "wyniki": wyniki,
        "raport_txt": str(txt_path),
        "raport_csv": str(csv_path),
    }


def literate_map(
    input_path: Path,
    output_dir: Path,
    layer: str,
    overwrite_existing: bool,
    complex_tolerance: float,
    log,
) -> tuple[Path, dict]:
    if not input_path.exists():
        raise LiteracjaError("Wskazany plik wejściowy nie istnieje.")
    output_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="literacja_upul_") as temp:
        workdir = Path(temp)
        source_root, main_map = prepare_input(input_path, workdir)
        rel_map = main_map.relative_to(source_root)

        raw = main_map.read_bytes()
        text = decode_map(raw)
        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        lines = normalized.split("\n")
        objects = parse_objects(lines)
        selected = [o for o in objects if o.layer == layer]
        if not selected:
            raise LiteracjaError(f"Nie znaleziono obiektów na warstwie {layer}.")

        without_oddzial = [o for o in selected if not o.oddzial]
        if without_oddzial:
            raise LiteracjaError(
                f"Na warstwie {layer} znaleziono {len(without_oddzial)} obiektów bez numeru oddziału w A1. "
                "Najpierw wpisz numer oddziału w A1."
            )

        eligible: list[GeoObject] = []
        skipped_existing: list[GeoObject] = []
        for obj in selected:
            current_a1 = obj.attrs.get("A1", "").strip()
            has_letter = bool(re.search(r"[A-Za-z]", current_a1))
            if has_letter and not overwrite_existing:
                skipped_existing.append(obj)
            else:
                eligible.append(obj)

        if not eligible:
            raise LiteracjaError("Nie ma obiektów do zaliterowania przy wybranych ustawieniach.")

        groups: dict[str, list[GeoObject]] = {}
        for obj in eligible:
            groups.setdefault(obj.oddzial or "", []).append(obj)

        complex_counts: dict[str, int] = {}
        for oddzial, group in groups.items():
            ordered, complexes = order_by_forest_complexes(group, complex_tolerance)
            group[:] = ordered
            complex_counts[oddzial] = len(complexes)
            letters = forestry_letters()
            for idx, obj in enumerate(group, start=1):
                letter = next(letters)
                obj.sequence_no = idx
                obj.new_a1 = f"{oddzial}{letter}"

        # Modyfikacja od końca pliku, aby indeksy się nie przesunęły po ewentualnym dodaniu A1.
        for obj in sorted(eligible, key=lambda o: o.start, reverse=True):
            lines[obj.start:obj.end] = replace_or_insert_a1(lines[obj.start:obj.end], obj.new_a1)

        output_text, checksum = update_checksum("\n".join(lines))
        output_bytes = output_text.encode(ENCODING, errors="replace")

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_name = f"{main_map.stem}_zaliterowane"
        result_root = workdir / result_name
        copy_tree_contents(source_root, result_root)

        # Kopia oryginalna całego wejścia.
        backup_root = result_root / "KOPIA_ORYGINALNA"
        copy_tree_contents(source_root, backup_root)

        target_map = result_root / rel_map
        target_map.parent.mkdir(parents=True, exist_ok=True)
        target_map.write_bytes(output_bytes)

        # Aktualizujemy odpowiadający plik ~AP, tak jak w GEO-MAP.
        for ap in result_root.rglob("*"):
            if "KOPIA_ORYGINALNA" in ap.parts:
                continue
            if ap.is_file() and ap.suffix.lower() == ".~ap" and ap.stem.lower() == main_map.stem.lower():
                ap.write_bytes(output_bytes)

        report_objects = sorted(eligible, key=lambda o: (int(o.oddzial or 0), o.sequence_no))
        csv_path = result_root / "RAPORT_LITERACJI.csv"
        with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow([
                "oddzial", "kompleks", "kolejnosc_w_oddziale", "TX_numer_dzialki", "nowy_A1",
                "litera", "powierzchnia_m2", "X", "Y", "linia_w_pliku"
            ])
            for obj in report_objects:
                letter = obj.new_a1[len(obj.oddzial or ""):]
                writer.writerow([
                    obj.oddzial, obj.complex_no, obj.sequence_no, obj.tx, obj.new_a1, letter,
                    round(obj.area, 2), round(obj.cx, 3), round(obj.cy, 3), obj.start + 1,
                ])

        txt_path = result_root / "RAPORT_LITERACJI.txt"
        report_lines = [
            f"RAPORT LITERACJI WARSTWY {layer}",
            f"Program: {APP_NAME} {APP_VERSION}",
            f"Plik: {main_map.name}",
            f"Liczba zmienionych obiektów: {len(eligible)}",
            f"Liczba pominiętych istniejących oznaczeń: {len(skipped_existing)}",
            f"Sposób kolejności: najpierw cały kompleks leśny, następnie kolejny kompleks od północnego wschodu do południowego zachodu. Tolerancja łączenia kompleksu: {complex_tolerance:g} m. Atrybut TX nie wpływa na kolejność.",
            "Zapis A1: pełny adres oddział + litera, np. 1a, 2ax.",
            "Ciąg liter: a,b,c,d,f,g,h,i,j,k,l,m,n,o,p,r,s,t,w,x,y,z, następnie ax,bx,...",
            "",
        ]
        for oddzial in sorted(groups, key=int):
            grp = sorted(groups[oddzial], key=lambda o: o.sequence_no)
            report_lines.append(f"Oddział {oddzial}: {len(grp)} obiektów, {complex_counts[oddzial]} kompleksów, od {grp[0].new_a1} do {grp[-1].new_a1}.")
        report_lines.extend(["", f"Suma kontrolna GEO-MAP CHK: {checksum}"])
        txt_path.write_text("\n".join(report_lines), encoding="utf-8")

        create_svg(report_objects, result_root / "PODGLAD_LITERACJI.svg")

        # Kontrola zapisu.
        saved_text = decode_map(target_map.read_bytes()).replace("\r\n", "\n").replace("\r", "\n")
        saved_objects = [o for o in parse_objects(saved_text.split("\n")) if o.layer == layer]
        saved_a1 = {o.attrs.get("ID", ""): o.attrs.get("A1", "") for o in saved_objects}
        for obj in eligible:
            obj_id = obj.attrs.get("ID", "")
            if obj_id and saved_a1.get(obj_id) != obj.new_a1:
                raise LiteracjaError(f"Kontrola zapisu nie powiodła się dla obiektu {obj_id}.")

        output_zip = output_dir / f"{result_name}_{stamp}.zip"
        with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for path in result_root.rglob("*"):
                if path.is_file():
                    zf.write(path, Path(result_name) / path.relative_to(result_root))

        stats = {
            "changed": len(eligible),
            "skipped": len(skipped_existing),
            "groups": {k: len(v) for k, v in sorted(groups.items(), key=lambda kv: int(kv[0]))},
            "complexes": complex_counts,
            "checksum": checksum,
            "zip": str(output_zip),
        }
        log(f"Zapisano: {output_zip}")
        return output_zip, stats
