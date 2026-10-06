# -*- coding: utf-8 -*-
"""Kodek Mazovia (CP667/CP790) — kodowanie DOS używane przez GEO-MAP.

Mazovia = CP437 z podmienionymi znakami (literami polskimi).
Pełne 256-pozycyjne mapowanie -> odwracalne dla każdego bajtu (bez surrogateescape).
"""
_CP437 = {b: bytes([b]).decode("cp437") for b in range(256)}

# bajt: (znak_w_CP437, znak_w_Mazovii)
_SUBST = {
    0x86: ("å", "ą"), 0x8D: ("ì", "ć"), 0x8F: ("Å", "Ą"), 0x90: ("É", "Ę"),
    0x91: ("æ", "ę"), 0x92: ("Æ", "ł"), 0x95: ("ò", "Ć"), 0x98: ("ÿ", "Ś"),
    0x9C: ("£", "Ł"), 0x9E: ("₧", "ś"), 0xA0: ("á", "Ź"), 0xA1: ("í", "Ż"),
    0xA3: ("ú", "Ó"), 0xA4: ("ñ", "ń"), 0xA5: ("Ñ", "Ń"), 0xA6: ("ª", "ź"),
    0xA7: ("º", "ż"),
}

DECODE_TABLE = {}   # bajt -> unicode
for b, ch in _CP437.items():
    DECODE_TABLE[b] = ch
for b, (_old, new) in _SUBST.items():
    DECODE_TABLE[b] = new

ENCODE_TABLE = {ch: b for b, ch in DECODE_TABLE.items()}


def decode(data, errors="strict"):
    if isinstance(data, str):
        data = data.encode("latin-1")
    out = []
    for b in data:
        ch = DECODE_TABLE[b]
        if ch is None and errors != "strict":
            ch = "\ufffd"
        out.append(ch)
    return "".join(out)


def encode(text, errors="strict"):
    out = bytearray()
    for ch in text:
        b = ENCODE_TABLE.get(ch)
        if b is None:
            if errors == "strict":
                raise UnicodeEncodeError("mazovia", text, 0, 1,
                                         "znak nie występuje w Mazovii: %r" % ch)
            b = 0x3F  # '?'
        out.append(b)
    return bytes(out)
