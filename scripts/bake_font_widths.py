"""Bake GDI's character widths for Aptos Narrow and Calibri into src/pyopenvba/apps/excel/_font_widths.py.

Excel measures a cell's text for AutoFit and Range.Text as GDI's
GetTextExtentPoint32 does at the font's height in pixels, the size in
points times 4/3 at 96 DPI: whole pixels that add up along a text for
ASCII and Latin-1. This asks GDI for every such character's width, and the
font's ascent, at each height from 11 pixels (8pt) to 96 (72pt), for both
fonts in regular, bold, italic and bold italic. At 15 pixels, 11pt, the
widths have to be the ones live Excel showed for each character alone and
between two x's (tests/fixtures/column_autofit.json, which
scripts/measure_column_autofit.py measured), and this checks they are;
tests/fixtures/autofit_sizes.json checks the other heights.

    python scripts/bake_font_widths.py

It needs Windows, with Aptos Narrow in the cloud font cache where Office
puts it, and the widths it finds are the display's at 96 DPI.
"""

from __future__ import annotations

import base64
import json
import textwrap
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "column_autofit.json"
MODULE = ROOT / "src" / "pyopenvba" / "apps" / "excel" / "_font_widths.py"
FONTS = ("aptos_narrow", "calibri")
STYLES = {"regular": (False, False), "bold": (True, False), "italic": (False, True), "bold_italic": (True, True)}
CODES = [*range(0x20, 0x7F), *range(0xA0, 0x100)]
#: The heights in pixels baked: 8pt to 72pt.
HEIGHTS = range(11, 97)


def table(result: str) -> tuple[int, dict[int, int]]:
    """The overhang a text in the style has past its characters, and each character's width."""
    pieces = dict(piece.split(":") for piece in result.split())
    one, two = int(pieces.pop("x")), int(pieces.pop("xx"))
    overhang = one - 7 - (two - one)
    widths: dict[int, int] = {}
    for code, measured in pieces.items():
        alone, before, after, between = (int(value) for value in measured.split(","))
        width = between - two
        # Nothing reaches past a character at either end of a text but the style's overhang.
        assert alone - 7 == width + overhang and before == after == one + width, (code, measured)
        widths[int(code)] = width
    return overhang, widths


class _GDI:
    """GDI's text extents and font metrics on this display, Aptos Narrow loaded from the cloud font cache."""

    def __init__(self) -> None:
        import ctypes
        import os
        from ctypes import wintypes

        self.ctypes, self.wintypes = ctypes, wintypes
        self.gdi32, user32 = ctypes.WinDLL("gdi32"), ctypes.WinDLL("user32")
        cache = Path(os.environ["LOCALAPPDATA"]) / "Microsoft/FontCache/4/CloudFonts/Aptos Narrow"
        fonts = list(cache.glob("*.ttf"))
        assert fonts, f"Aptos Narrow is not in {cache}: open a workbook in it in Excel once"
        for path in fonts:
            # FR_PRIVATE: for this process alone.
            self.gdi32.AddFontResourceExW(str(path), 0x10, None)

        class LOGFONTW(ctypes.Structure):
            _fields_ = [("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG), ("lfEscapement", wintypes.LONG),
                        ("lfOrientation", wintypes.LONG), ("lfWeight", wintypes.LONG), ("lfItalic", wintypes.BYTE),
                        ("lfUnderline", wintypes.BYTE), ("lfStrikeOut", wintypes.BYTE), ("lfCharSet", wintypes.BYTE),
                        ("lfOutPrecision", wintypes.BYTE), ("lfClipPrecision", wintypes.BYTE),
                        ("lfQuality", wintypes.BYTE), ("lfPitchAndFamily", wintypes.BYTE),
                        ("lfFaceName", wintypes.WCHAR * 32)]

        class TEXTMETRICW(ctypes.Structure):
            _fields_ = [("tmHeight", wintypes.LONG), ("tmAscent", wintypes.LONG), ("tmDescent", wintypes.LONG),
                        ("tmInternalLeading", wintypes.LONG), ("tmExternalLeading", wintypes.LONG),
                        ("tmAveCharWidth", wintypes.LONG), ("tmMaxCharWidth", wintypes.LONG),
                        ("tmWeight", wintypes.LONG), ("tmOverhang", wintypes.LONG),
                        ("tmDigitizedAspectX", wintypes.LONG), ("tmDigitizedAspectY", wintypes.LONG),
                        ("tmFirstChar", wintypes.WCHAR), ("tmLastChar", wintypes.WCHAR),
                        ("tmDefaultChar", wintypes.WCHAR), ("tmBreakChar", wintypes.WCHAR),
                        ("tmItalic", wintypes.BYTE), ("tmUnderlined", wintypes.BYTE), ("tmStruckOut", wintypes.BYTE),
                        ("tmPitchAndFamily", wintypes.BYTE), ("tmCharSet", wintypes.BYTE)]

        self.LOGFONTW, self.TEXTMETRICW = LOGFONTW, TEXTMETRICW
        user32.GetDC.restype = wintypes.HDC
        self.gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
        self.gdi32.SelectObject.restype = wintypes.HGDIOBJ
        self.gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
        self.gdi32.GetTextExtentPoint32W.argtypes = [wintypes.HDC, wintypes.LPCWSTR, ctypes.c_int,
                                                     ctypes.POINTER(wintypes.SIZE)]
        self.gdi32.GetTextMetricsW.argtypes = [wintypes.HDC, ctypes.POINTER(TEXTMETRICW)]
        self.gdi32.CreateFontIndirectW.restype = wintypes.HFONT
        self.dc = user32.GetDC(None)

    def measure(self, face: str, bold: bool, italic: bool, pixels: int) -> tuple[list[int], int]:
        """Each of CODES' widths, and the font's ascent, at a height in pixels."""
        logical = self.LOGFONTW(lfHeight=-pixels, lfWeight=700 if bold else 400, lfItalic=int(italic), lfCharSet=1,
                                lfFaceName=face)
        font = self.gdi32.CreateFontIndirectW(self.ctypes.byref(logical))
        old = self.gdi32.SelectObject(self.dc, font)
        size = self.wintypes.SIZE()
        widths = []
        for code in CODES:
            self.gdi32.GetTextExtentPoint32W(self.dc, chr(code), 1, self.ctypes.byref(size))
            widths.append(size.cx)
        metrics = self.TEXTMETRICW()
        self.gdi32.GetTextMetricsW(self.dc, self.ctypes.byref(metrics))
        self.gdi32.SelectObject(self.dc, old)
        self.gdi32.DeleteObject(font)
        return widths, metrics.tmAscent


def main() -> None:
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    results = {probe["name"]: probe["result"] for probe in fixture["probes"]}
    gdi = _GDI()
    data = bytearray()
    for font in FONTS:
        face = font.replace("_", " ").title()
        for style, (bold, italic) in STYLES.items():
            _, live = table(results[f"glyphs_aptos_narrow_{font}_{style}"])
            for pixels in HEIGHTS:
                widths, ascent = gdi.measure(face, bold, italic, pixels)
                if pixels == 15:
                    assert widths == [live[code] for code in CODES], (font, style)
                assert max(widths) < 256 and ascent < 256, (font, style, pixels)
                data += bytes(widths) + bytes([ascent])
    blob = base64.b64encode(zlib.compress(bytes(data), 9)).decode("ascii")
    chunks = "\n".join(f'    "{chunk}"' for chunk in textwrap.wrap(blob, 100))
    MODULE.write_text(f'''"""How wide Excel measures each character for AutoFit and Range.Text, on a 96-DPI display.

Generated by scripts/bake_font_widths.py from GDI, whose GetTextExtentPoint32
is how Excel measures a cell's text, at the font's height in pixels: the
size in points times 4/3. For Aptos Narrow and Calibri from 8pt to 72pt, in
regular, bold, italic and bold italic, a table holds each character's width
in pixels, for ASCII from the space to the tilde and Latin-1 from the
no-break space to the y with diaeresis, and the font's ascent. The widths
are whole pixels that add up along a text; for Windows-1252's other
characters GDI's widths depend on their neighbours, so they are not here.
At 11pt they are what live Excel showed for each character
(tests/fixtures/column_autofit.json); tests/fixtures/autofit_sizes.json
checks the other sizes. Do not edit by hand.
"""

from __future__ import annotations

import base64
import zlib
from functools import cache

#: The characters each table holds, in order.
CODES = (*range(0x20, 0x7F), *range(0xA0, 0x100))
#: The fonts and styles the tables hold, in the order the data has them.
_FACES = tuple((font, bold, italic) for font in ("aptos narrow", "calibri")
               for bold, italic in ((False, False), (True, False), (False, True), (True, True)))
#: The heights in pixels the tables hold, in order: {HEIGHTS.start} ({8}pt) to {HEIGHTS.stop - 1} ({72}pt).
_HEIGHTS = range({HEIGHTS.start}, {HEIGHTS.stop})
#: For each font and style, then each height: each character's width, then the ascent, a byte each, compressed.
_DATA = (
{chunks}
)


@cache
def _tables() -> bytes:
    return zlib.decompress(base64.b64decode(_DATA))


def pixel_height(size: float) -> int | None:
    """The height in pixels of a font ``size`` points tall at 96 DPI, which the tables are kept by; None for a size
    they do not hold, past 8 to 72pt or between half points."""
    if size * 2 != int(size * 2):
        return None
    pixels = int(size * 4 / 3 + 0.5)
    return pixels if pixels in _HEIGHTS else None


def _row(font: str, bold: bool, italic: bool, size: float) -> bytes | None:
    pixels = pixel_height(size)
    key = (font.lower(), bold, italic)
    if pixels is None or key not in _FACES:
        return None
    width = len(CODES) + 1
    start = (_FACES.index(key) * len(_HEIGHTS) + _HEIGHTS.index(pixels)) * width
    return _tables()[start:start + width]


@cache
def widths(font: str, bold: bool, italic: bool, size: float = 11.0) -> dict[str, int] | None:
    """Each character's width in pixels in a font, style and size, or None for one that was not measured."""
    row = _row(font, bold, italic, size)
    return None if row is None else {{chr(code): width for code, width in zip(CODES, row[:-1], strict=True)}}


def ascent(font: str, bold: bool, italic: bool, size: float = 11.0) -> int | None:
    """How far the font reaches above its baseline in pixels, or None for one that was not measured."""
    row = _row(font, bold, italic, size)
    return None if row is None else row[-1]
''', encoding="utf-8")
    print(f"wrote {MODULE.relative_to(ROOT)}: {len(data)} bytes in {len(blob)} characters")


if __name__ == "__main__":
    main()
