"""TextRange2.Font.Fill and Line: the paint of a shape's letters, over all of its text or part of it, as
tests/fixtures/shape_font2_paint.json measured them. Their rules differ from a shape's own fill and line, so they keep to
their own classes. Gradients, patterns and textures on the letters report themselves."""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._held_shapes import ShapePart
from pyopenvba.apps.excel._shape_format import _THEME_SLOT, SCHEMES, _lum, _scheme_number, _tint, painted  # pyright: ignore[reportPrivateUsage]
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, VBASingle, error, to_bool, to_integer, to_number
from pyopenvba.shapes._values import TextLine, emu

if TYPE_CHECKING:
    from pyopenvba.apps.excel._shapes import _ShapeFontBase  # pyright: ignore[reportPrivateUsage]

_INVALID = (-2147024809, "The parameter is incorrect.")
_MIXED_LONG = VBAInt(-2147483648, "Long")
_MIXED_SINGLE = VBASingle(-2147483648.0)
_MIXED_TRI = VBAInt(-2, "Long")
#: What a letter's colour is made of, across a range.
_INK_FIELDS = ("font_rgb", "font_index", "font_tint", "font_theme", "font_mods", "font_no_fill")
_DASHES = {1: "solid", 4: "dash"}


def _hex(value: object) -> str:
    colour = int(to_number(value)) & 0xFFFFFF
    return f"{colour & 255:02X}{colour >> 8 & 255:02X}{colour >> 16 & 255:02X}"


def _theme_colour(font: _ShapeFontBase, scheme: str) -> str:
    return font.characters.sheet.book.stylesheet.colors.theme[_THEME_SLOT[scheme]]


def font_ink(font: _ShapeFontBase) -> int:
    """The colour the letters read as, as a BGR Long. With no fill they read as white, in a text box too."""
    look = font.look()
    if look.font_no_fill:
        return 16777215
    mods = (("tint", look.font_tint),) if look.font_tint is not None else look.font_mods
    return painted(look.font_rgb or _theme_colour(font, look.font_theme or look.font_scheme), mods)


def _positive(value: object, what: str) -> float:
    wanted = float(VBASingle(float(to_number(value))))
    if wanted <= 0:
        raise VBAUnsupportedError(f"This {what} of a shape's letters is not implemented")
    return wanted


def _opacity(value: object) -> int | None:
    """The alpha a Transparency writes, none for 0, which also writes the colour out."""
    wanted = float(VBASingle(float(to_number(value))))
    if not 0 <= wanted <= 1:
        raise error(*_INVALID)
    return None if wanted == 0 else round((1 - wanted) * 100000)


class _Paint(ShapePart):
    """Of a Font2's letters."""

    vba_type_name = ""

    def __init__(self, font: _ShapeFontBase) -> None:
        self.font = font

    @member
    def Parent(self) -> object:
        return self.font


class TextFill(_Paint):
    vba_type_name = "FillFormat"

    def _mixed(self, *names: str) -> bool:
        return len(set(self.font.values(names))) > 1

    @member
    def Visible(self) -> object:
        # Any letter with a fill makes it visible.
        return VBAInt(-1 if any(not row[0] for row in self.font.values(("font_no_fill",))) else 0, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        if to_bool(value):
            if self.font.look().font_no_fill:
                # Showing it again brings back no colour of the macro's, whatever it had: always this one.
                look = self.font.changed()
                look.font_no_fill, look.font_rgb = False, "BBE0E3"
            return
        look = self.font.changed()
        look.font_no_fill = True
        look.font_rgb, look.font_index, look.font_tint, look.font_theme, look.font_mods = None, None, None, None, ()
        look.font_alpha, look.font_solid = None, False

    @member
    def Type(self) -> object:
        return _MIXED_TRI if self._mixed("font_no_fill") or self.font.look().font_no_fill else VBAInt(1, "Long")

    @member
    def ForeColor(self) -> object:
        return TextFillColour(self.font)

    @member
    def BackColor(self) -> object:
        return TextBackColour(self.font)

    @member
    def Transparency(self) -> object:
        if self._mixed("font_alpha", "font_no_fill"):
            return _MIXED_SINGLE
        look = self.font.look()
        if look.font_no_fill:
            return VBASingle(1.0)
        return VBASingle(0.0 if look.font_alpha is None else 1 - look.font_alpha / 100000)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        alpha = _opacity(value)
        if self.font.look().font_no_fill:
            raise VBAUnsupportedError("Transparency of letters with no fill is not implemented")
        look = self.font.changed()
        look.font_alpha = alpha
        look.font_solid = alpha is None

    @method
    def Solid(self) -> object:
        if self.font.look().font_no_fill:
            raise VBAUnsupportedError("Solid on letters with no fill is not implemented")
        self.font.changed().font_solid = True
        return EMPTY

    @member
    def Pattern(self) -> object:
        return _MIXED_TRI

    @setter("Pattern")
    def _pattern(self, value: object) -> None:
        raise VBAUnsupportedError("A pattern on a shape's letters is not implemented")


class TextFillColour(_Paint):
    vba_type_name = "ColorFormat"

    def _mixed(self) -> bool:
        return len(set(self.font.values(_INK_FIELDS))) > 1

    @member
    def RGB(self) -> object:
        return _MIXED_LONG if self._mixed() else VBAInt(font_ink(self.font), "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        look = self.font.changed()
        look.font_rgb, look.font_index, look.font_tint, look.font_theme, look.font_mods = _hex(value), None, None, None, ()
        look.font_no_fill = False

    @member
    def ObjectThemeColor(self) -> object:
        if self._mixed():
            return _MIXED_TRI
        look = self.font.look()
        if look.font_rgb is not None or look.font_no_fill:
            return VBAInt(0, "Long")
        return VBAInt(_scheme_number(look.font_theme or look.font_scheme), "Long")

    @setter("ObjectThemeColor")
    def _object_theme_color(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted not in SCHEMES:
            raise VBAUnsupportedError("This ObjectThemeColor of a shape's letters is not implemented")
        look = self.font.changed()
        look.font_theme, look.font_rgb, look.font_index, look.font_tint, look.font_mods = SCHEMES[wanted], None, None, None, ()
        look.font_no_fill = False

    @member
    def Type(self) -> object:
        if self._mixed():
            return _MIXED_TRI
        look = self.font.look()
        return VBAInt(1 if look.font_rgb is not None or look.font_no_fill else 2, "Long")

    @member
    def TintAndShade(self) -> object:
        if self._mixed():
            return _MIXED_SINGLE
        tint = self.font.look().font_tint
        return VBASingle(0.0 if tint is None else 1 - tint / 100000)

    @setter("TintAndShade")
    def _tint_and_shade(self, value: object) -> None:
        wanted = _positive(value, "TintAndShade")
        current = self.font.look()
        if current.font_rgb is not None or current.font_no_fill or current.font_mods:
            raise VBAUnsupportedError("TintAndShade of letters with this colour is not implemented")
        self.font.changed().font_tint = round((1 - wanted) * 100000)

    @member
    def Brightness(self) -> object:
        if self._mixed():
            return _MIXED_SINGLE
        off = dict(self.font.look().font_mods).get("lumOff")
        return VBASingle(0.0 if off is None else off / 100000)

    @setter("Brightness")
    def _brightness(self, value: object) -> None:
        wanted = _positive(value, "Brightness")
        current = self.font.look()
        if current.font_no_fill or current.font_tint is not None:
            raise VBAUnsupportedError("Brightness of letters with this colour is not implemented")
        self.font.changed().font_mods = _lum(wanted)


class TextBackColour(_Paint):
    """A second colour, which solid letters do not use. Setting it writes their own colour out."""

    vba_type_name = "ColorFormat"

    @member
    def RGB(self) -> object:
        return _MIXED_LONG if len(set(self.font.values(("font_no_fill",)))) > 1 else VBAInt(16777215, "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        if self.font.look().font_no_fill:
            raise VBAUnsupportedError("BackColor of letters with no fill is not implemented")
        self.font.changed().font_solid = True


def _apply(font: _ShapeFontBase, change: Callable[[TextLine | None], TextLine | None]) -> None:
    """Change each character's line on its own terms: over the whole text, those with a line of their own keep it."""
    base = font.frame._changed()  # pyright: ignore[reportPrivateUsage]
    shape = font.characters.owner.shape
    at, count = font.characters.span()
    if count == len(shape.text):
        for mark in [*shape.text_marks, *([] if shape.text_end is None else [shape.text_end])]:
            if "font_line" in mark:
                mark["font_line"] = change(mark["font_line"])  # pyright: ignore[reportArgumentType]
        base.font_line = change(base.font_line)
        return
    if not shape.text_marks:
        shape.text_marks = [{} for _ in shape.text]
    for mark in shape.text_marks[at:at + count]:
        mark["font_line"] = change(mark.get("font_line", base.font_line))  # pyright: ignore[reportArgumentType]


def _hidden_or(change: Callable[[TextLine], TextLine]) -> Callable[[TextLine | None], TextLine]:
    """A colour or transparency set where no line shows leaves a line of no fill, and nothing else."""
    def changed(line: TextLine | None) -> TextLine:
        return TextLine(hidden=True) if line is None or line.hidden else change(line)
    return changed


def _shown_or_new(change: Callable[[TextLine], TextLine]) -> Callable[[TextLine | None], TextLine]:
    """A width, dash or style shows a line where none does, in the text colour."""
    def changed(line: TextLine | None) -> TextLine:
        return change(TextLine(scheme="tx1") if line is None or line.hidden else line)
    return changed


class TextLineFormat(_Paint):
    vba_type_name = "LineFormat"

    def lines(self) -> list[TextLine | None]:
        """The line each character shows."""
        found: list[TextLine | None] = []
        for row in self.font.values(("font_line",)):
            line = row[0]
            found.append(line if isinstance(line, TextLine) and not line.hidden else None)
        return found

    def shown(self) -> list[TextLine]:
        return [line for line in self.lines() if line is not None]

    @member
    def Visible(self) -> object:
        return VBAInt(-1 if self.shown() else 0, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        # Setting it False has no effect, even on a line that shows.
        if to_bool(value):
            _apply(self.font, lambda line: TextLine(preset=True) if line is None or line.hidden else line)

    @member
    def ForeColor(self) -> object:
        return TextLineColour(self.font, self)

    @member
    def BackColor(self) -> object:
        return TextLineBack(self.font, self)

    @member
    def Weight(self) -> object:
        widths = {9525 if line.width is None else line.width for line in self.shown()}
        return VBASingle(widths.pop() / 12700) if len(widths) == 1 else _MIXED_SINGLE

    @setter("Weight")
    def _weight(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if wanted < 0:
            raise error(*_INVALID)
        _apply(self.font, _shown_or_new(lambda line: replace(line, width=emu(wanted))))

    @member
    def DashStyle(self) -> object:
        dashes = {line.dash or "solid" for line in self.shown()}
        if len(dashes) != 1:
            return _MIXED_TRI
        dash = dashes.pop()
        return VBAInt(next(key for key, name in _DASHES.items() if name == dash), "Long")

    @setter("DashStyle")
    def _dash_style(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted not in _DASHES:
            raise VBAUnsupportedError("This DashStyle of a shape's letters is not implemented")
        _apply(self.font, _shown_or_new(lambda line: replace(line, dash=_DASHES[wanted])))

    @member
    def Style(self) -> object:
        styles = {line.compound for line in self.shown()}
        return VBAInt(2 if styles.pop() else 1, "Long") if len(styles) == 1 else _MIXED_TRI

    @setter("Style")
    def _style(self, value: object) -> None:
        if int(to_integer(value, "Long")) != 2:
            raise VBAUnsupportedError("This Style of a shape's letters is not implemented")
        _apply(self.font, _shown_or_new(lambda line: replace(line, compound="dbl")))

    @member
    def Transparency(self) -> object:
        lines = self.lines()
        if lines[0] is None or len({None if line is None else (line.alpha,) for line in lines}) != 1:
            return _MIXED_SINGLE
        alpha = lines[0].alpha
        return VBASingle(0.0 if alpha is None else 1 - alpha / 100000)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        alpha = _opacity(value)
        _apply(self.font, _hidden_or(lambda line: replace(line, alpha=alpha, full=True)))

    @member
    def Pattern(self) -> object:
        return _MIXED_TRI

    @setter("Pattern")
    def _pattern(self, value: object) -> None:
        raise VBAUnsupportedError("A patterned line on a shape's letters is not implemented")


class TextLineColour(_Paint):
    vba_type_name = "ColorFormat"

    def __init__(self, font: _ShapeFontBase, line: TextLineFormat) -> None:
        super().__init__(font)
        self.line = line

    def _one(self) -> tuple[bool, TextLine | None]:
        """Whether the characters differ in colour, and the colour they have."""
        lines = self.line.lines()
        colours = {None if line is None else (line.rgb, line.scheme, line.mods) for line in lines}
        return len(colours) > 1, lines[0] if lines else None

    @member
    def RGB(self) -> object:
        mixed, line = self._one()
        if mixed:
            return _MIXED_LONG
        if line is None:
            return VBAInt(16777215, "Long")
        return VBAInt(painted(line.rgb or ("000000" if line.scheme is None else _theme_colour(self.font, line.scheme)), line.mods), "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        wanted = _hex(value)
        _apply(self.font, _hidden_or(lambda line: replace(line, rgb=wanted, scheme=None, preset=False, mods=())))

    @member
    def ObjectThemeColor(self) -> object:
        mixed, line = self._one()
        if mixed:
            return _MIXED_TRI
        return VBAInt(0 if line is None or line.scheme is None else _scheme_number(line.scheme), "Long")

    @setter("ObjectThemeColor")
    def _object_theme_color(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted not in SCHEMES:
            raise VBAUnsupportedError("This ObjectThemeColor of a shape's letters is not implemented")
        _apply(self.font, _hidden_or(lambda line: replace(line, rgb=None, scheme=SCHEMES[wanted], preset=False, mods=())))

    @member
    def Type(self) -> object:
        mixed, line = self._one()
        if mixed:
            return _MIXED_TRI
        return VBAInt(1 if line is None or line.scheme is None else 2, "Long")

    @member
    def TintAndShade(self) -> object:
        mixed, line = self._one()
        if mixed:
            return _MIXED_SINGLE
        tint = None if line is None else dict(line.mods).get("tint")
        return VBASingle(0.0 if tint is None else 1 - tint / 100000)

    @setter("TintAndShade")
    def _tint_and_shade(self, value: object) -> None:
        wanted = _positive(value, "TintAndShade")
        _apply(self.font, _hidden_or(lambda line: replace(line, mods=_tint(wanted))))

    @member
    def Brightness(self) -> object:
        mixed, line = self._one()
        if mixed:
            return _MIXED_SINGLE
        off = None if line is None else dict(line.mods).get("lumOff")
        return VBASingle(0.0 if off is None else off / 100000)

    @setter("Brightness")
    def _brightness(self, value: object) -> None:
        wanted = _positive(value, "Brightness")
        _apply(self.font, _hidden_or(lambda line: replace(line, mods=_lum(wanted))))


class TextLineBack(_Paint):
    vba_type_name = "ColorFormat"

    def __init__(self, font: _ShapeFontBase, line: TextLineFormat) -> None:
        super().__init__(font)
        self.line = line

    @member
    def RGB(self) -> object:
        return VBAInt(16777215, "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        if self.line.shown():
            raise VBAUnsupportedError("BackColor of a line on a shape's letters is not implemented")
        _apply(self.font, lambda _: TextLine(hidden=True))
