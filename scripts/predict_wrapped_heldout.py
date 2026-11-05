"""Predictions for measure_wrapped_heldout.py, written and committed before Excel is asked: fresh strings in fonts and sizes the rule was not fitted on, at widths past 1,040 pixels.

    python scripts/predict_wrapped_heldout.py

The rule: a line w pixels wide by GDI, whose digit 0 is d wide, stays on one line in a column of c pixels from c = w + 7 + 2 * max(0, (d - 5) // 4) + max(0, (w - 57) // 28).

Writes tests/fixtures/wrapped_heldout_predictions.json.
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as w
import json
import random
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/wrapped_heldout_predictions.json'
FONTS = (('Calibri', 8), ('Calibri', 9), ('Calibri', 10), ('Calibri', 12), ('Calibri', 14), ('Calibri', 18), ('Calibri', 20), ('Calibri', 24), ('Calibri', 36), ('Arial', 8), ('Arial', 12), ('Arial', 18), ('Times New Roman', 11), ('Times New Roman', 16),
         ('Tahoma', 11), ('Verdana', 10), ('Courier New', 11), ('Segoe UI', 11), ('Cambria', 12), ('Consolas', 14))
WORDS = 'quartz vexing jumbled glyphs zephyr blackout pivot ledger margin kerning oxide thimble cascade nimbus 1984 7.25 (draft) well-known Mixed CASE, stop. what? yes! 100% $5 a/b ~tilde under_score'.split()


def main() -> None:
    gdi = ctypes.windll.gdi32
    user = ctypes.windll.user32
    gdi.CreateFontW.restype = w.HFONT
    gdi.CreateFontW.argtypes = [ctypes.c_int] * 5 + [w.DWORD] * 8 + [w.LPCWSTR]
    gdi.SelectObject.argtypes = [w.HDC, w.HGDIOBJ]
    gdi.GetTextExtentPoint32W.argtypes = [w.HDC, w.LPCWSTR, ctypes.c_int, ctypes.POINTER(w.SIZE)]
    gdi.GetTextFaceW.argtypes = [w.HDC, ctypes.c_int, w.LPWSTR]
    hdc = user.GetDC(None)
    assert gdi.GetDeviceCaps(hdc, 88) == 96, 'predictions need the display at 96 DPI'
    generator = random.Random(20261104)

    def extent(text: str) -> int:
        size = w.SIZE()
        gdi.GetTextExtentPoint32W(hdc, text, len(text), ctypes.byref(size))
        return size.cx

    cases = []
    for font, points in FONTS:
        gdi.SelectObject(hdc, gdi.CreateFontW(-round(points * 96 / 72), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, font))
        face = ctypes.create_unicode_buffer(64)
        gdi.GetTextFaceW(hdc, 64, face)
        digit = extent('0')
        pad = 7 + 2 * max(0, (digit - 5) // 4)
        # Widths spread out to near the widest column, and some either side of a step of the term.
        for goal in (30, 56, 57, 58, 84, 85, 86, 113, 200, 300, 500, 800, 1100, 1400, 1650):
            text = ''
            while extent(text) < goal:
                text = (text + ' ' + generator.choice(WORDS)).strip()
            # Come back to the goal with single letters where the words overran it.
            while len(text) > 1 and extent(text[:-1]) >= goal:
                text = text[:-1].rstrip()
            wide = extent(text)
            if wide + pad + (wide // 28) > 1750:
                continue
            cases.append({'font': font, 'size': points, 'face': face.value, 'digit': digit, 'text': text, 'gdi': wide, 'predicted': wide + pad + max(0, (wide - 57) // 28)})
    user.ReleaseDC(None, hdc)
    OUT.write_text(json.dumps({'rule': 'c = w + 7 + 2 * max(0, (d - 5) // 4) + max(0, (w - 57) // 28)', 'cases': cases}, indent=0) + '\n', encoding='utf-8')
    print(len(cases), 'predictions', sorted({(one['font'], one['face']) for one in cases if one['font'] != one['face']}))


if __name__ == '__main__':
    main()
