"""A database's Office theme, and the colours and fonts it gives a design.

A form or report is drawn with the theme its database keeps in
`MSysResources`: the `.thmx` package Office saves, in an attachment on the
row MSysDb's `Theme Resource Name` property names.  A design stores what
the theme gives it -- a font face beside the theme font it came from, a
colour beside the theme slot, tint and shade it was worked out from -- so
the same new form differs between databases whose themes differ.  A
database with no theme yet gets the one Access installs with its first form
or report, `OFFICE_2023`.
"""

from __future__ import annotations

import colorsys
import io
import math
import re
import zipfile
from dataclasses import dataclass
from xml.etree import ElementTree

#: Access's theme colour slots, in the order its ThemeColorIndex counts them.
SLOTS = (
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)
_DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


@dataclass(frozen=True)
class Theme:
    """What a design takes from its database's theme."""

    major_font: str
    minor_font: str
    #: `RRGGBB` for each slot, in `SLOTS` order.
    colors: tuple[str, ...]


#: Office's 2007 theme, the one the shipped blank database holds and the
#: library's design templates were captured in.
OFFICE_2007 = Theme(
    "Cambria", "Calibri",
    ("000000", "FFFFFF", "1F497D", "EEECE1", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646",
     "0000FF", "800080"),
)
#: Office's 2023 theme, which Access 16 installs in a database with its
#: first form or report.
OFFICE_2023 = Theme(
    "Aptos Display", "Aptos",
    ("000000", "FFFFFF", "0E2841", "E8E8E8", "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E",
     "467886", "96607D"),
)


def parse_theme(thmx: bytes) -> Theme:
    """The fonts and colours of a `.thmx` package."""
    with zipfile.ZipFile(io.BytesIO(thmx)) as package:
        name = next(n for n in package.namelist() if re.fullmatch(r"theme/theme/theme\d+\.xml", n))
        root = ElementTree.fromstring(package.read(name))
    scheme = root.find(f"{_DRAWING}themeElements/{_DRAWING}clrScheme")
    fonts = root.find(f"{_DRAWING}themeElements/{_DRAWING}fontScheme")
    if scheme is None or fonts is None:
        raise ValueError("the theme has no colour or font scheme")
    colors: list[str] = []
    for slot in SLOTS:
        entry = scheme.find(f"{_DRAWING}{slot}")
        value = None
        if entry is not None and len(entry):
            spec = entry[0]
            value = spec.get("lastClr") if spec.tag == f"{_DRAWING}sysClr" else spec.get("val")
        if value is None:
            raise ValueError(f"the theme gives no colour for {slot}")
        colors.append(value.upper())

    def face(kind: str) -> str:
        latin = fonts.find(f"{_DRAWING}{kind}/{_DRAWING}latin")
        typeface = latin.get("typeface") if latin is not None else None
        if not typeface:
            raise ValueError(f"the theme names no {kind}")
        return typeface

    return Theme(face("majorFont"), face("minorFont"), tuple(colors))


# --- theme colours -------------------------------------------------------------
# A tint lightens a slot's colour and a shade darkens it, both percentages,
# 100 leaving it alone: in HSL, luminance goes to L*t + (1 - t) for a tint
# and to L*s for a shade.  That gives every colour Access was asked for --
# each slot of both Office themes at every whole tint and shade, 4,444 in
# all -- to within rounding.  Where a channel lands on or next to a half,
# Access rounds as its own arithmetic falls, which is not reproduced; the
# table holds the 148 answers of those it gave that differ from rounding
# half up.  A colour of a theme not measured can differ by one there.
_MEASURED = " ".join((
    "000000t010E5E5E5 000000t0507F7F7F 0000FFs010000019 0000FFs05000007F 0000FFs0900000E5 0000FFt010E5E5FF",
    "0000FFt0507F7FFF 0000FFt0901919FF 0E2841s010010406 0E2841s025030A10 0E2841s050071420 0E2841s0700A1C2E",
    "0E2841t0543F8CD6 0F9ED5t0674FC5F3 156082s005010506 156082s025051820 156082s03507222E 156082s0550C3547",
    "156082s08512526F 156082s095145B7C 156082t095186B91 156082t099166285 196B24s002000201 196B24s014030F05",
    "196B24s026061C09 196B24s03809290E 196B24s0500D3512 196B24s0580E3E15 196B24s066104718 196B24s070114B19",
    "196B24s07813531C 196B24s08214581E 196B24s090166020 196B24s098186923 196B24t03987E394 1F497Ds006020408",
    "1F497Ds014040A11 1F497Ds02207101B 1F497Ds026081320 1F497Ds030091626 1F497Ds0340B192A 1F497Ds0420D1F35",
    "1F497Ds05010253F 1F497Ds054112743 1F497Ds058122A48 1F497Ds082193C67 1F497Ds0861B3F6C 1F497Ds0981E487B",
    "467886s005030607 467886s0150B1214 467886s025111E22 467886s035192A2F 467886s0451F363C 467886s05526424A",
    "467886s075355A64 467886s09542727F 467886t010EBF2F4 467886t030C3D9DF 467886t0509BC0CA 467886t07072A7B5",
    "4BACC6t050A5D6E2 4EA72Es025132A0C 4EA72Es050275317 4F81BDs013091119 800080t025FF9FFF 800080t075DF00DF",
    "8064A2t050BFB2D0 8064A2t070A692BE 8064A2t0908D74AB 96607Ds002030203 96607Ds006090608 96607Ds0100F0A0D",
    "96607Ds013130C10 96607Ds015160E13 96607Ds017191015 96607Ds0211F141A 96607Ds027281A22 96607Ds0292C1C24",
    "96607Ds0302D1D26 96607Ds0312E1E27 96607Ds033312029 96607Ds03737242E 96607Ds038392430 96607Ds0413E2733",
    "96607Ds0423F2835 96607Ds043402936 96607Ds045432B38 96607Ds046452C3A 96607Ds047462D3B 96607Ds0504B303F",
    "96607Ds057563747 96607Ds0615C3B4C 96607Ds0625D3C4E 96607Ds066633F53 96607Ds067644054 96607Ds069684256",
    "96607Ds0736D465B 96607Ds0746F475D 96607Ds078754B62 96607Ds0827B4F67 96607Ds0837C5068 96607Ds08681536C",
    "96607Ds08782546D 96607Ds090875671 96607Ds0938B5974 96607Ds0958F5B77 96607Ds097915D79 96607Ds098935E7B",
    "96607Ds099955F7C 9BBB59s01213180A 9BBB59t075B4CC82 A02B93s05050164A C0504Ds032401816 C0504Dt002FEFBFB",
    "C0504Dt014F6E6E6 C0504Dt018F4DFDF C0504Dt022F1D8D8 C0504Dt025EFD3D2 C0504Dt026EFD1D1 C0504Dt030ECCBCA",
    "C0504Dt038E7BDBB C0504Dt046E2AEAD C0504Dt050DFA8A6 C0504Dt062D89291 C0504Dt075D07C79 C0504Dt078CE7774",
    "C0504Dt086C96866 C0504Dt090C6615F C0504Dt094C45B58 E8E8E8t050F4F4F4 E97132s020331506 E97132t010FDF1EA",
    "E97132t025FADCCC E97132t075EF9465 EEECE1t025FBFAF7 EEECE1t050F6F5F0 EEECE1t055F6F5EF EEECE1t075F2F1E8",
    "EEECE1t095EFEDE2 F79646s013271302 F79646s0193A1B03 F79646s087F5801F F79646t010FEF4EC F79646t030FDDFC8",
    "F79646t090F8A059 FFFFFFs010191919 FFFFFFs0507F7F7F FFFFFFs090E5E5E5",
))
_CORRECTIONS = {entry[:10]: entry[10:] for entry in _MEASURED.split()}


def tinted(color: str, tint: float = 100.0, shade: float = 100.0) -> str:
    """`color` (`RRGGBB`) lightened by `tint` or darkened by `shade`, as
    Access works it out."""
    whole = (tint, "t") if shade == 100 else (shade, "s") if tint == 100 else None
    if whole is not None and whole[0] == int(whole[0]):
        found = _CORRECTIONS.get(f"{color.upper()}{whole[1]}{int(whole[0]):03d}")
        if found is not None:
            return found
    red, green, blue = (int(color[at : at + 2], 16) / 255 for at in (0, 2, 4))
    hue, luminance, saturation = colorsys.rgb_to_hls(red, green, blue)
    # In this order: the table above corrects exactly this arithmetic.
    if tint != 100:
        fraction = tint / 100
        luminance = luminance * fraction + (1 - fraction)
    if shade != 100:
        luminance = luminance * (shade / 100)
    channels = colorsys.hls_to_rgb(hue, luminance, saturation)
    return "".join(f"{math.floor(channel * 255 + 0.5):02X}" for channel in channels)


def theme_color(theme: Theme, index: int, tint: float = 100.0, shade: float = 100.0) -> bytes:
    """A slot's colour as a design record holds it: red, green, blue, 0."""
    return bytes.fromhex(tinted(theme.colors[index], tint, shade)) + b"\x00"
