"""TextRange2.ParagraphFormat, and the Font2 effects a shape's letters read as having none of, as
tests/fixtures/loaded_shape_text.json and tests/fixtures/loaded_shape_sets.json measured them. Settings not yet measured report
themselves.

Each paragraph's settings are its pPr, as the file spells it. A shape holds them in text_paragraphs only where the paragraphs are not
all the look's plain algn and rtl."""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._held import HeldPart
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, setter
from pyopenvba.interpreter._values import VBAInt, VBASingle, error, to_number
from pyopenvba.shapes._read_text import paragraph_settings
from pyopenvba.shapes._values import Shape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._shapes import TextRange2

#: ParagraphFormat.Alignment for each algn.
_ALIGNMENTS = {"l": 1, "ctr": 2, "r": 3, "just": 4, "dist": 5, "thaiDist": 6, "justLow": 7}
#: BaselineAlignment for each fontAlgn.
_FONT_ALIGNS = {"auto": 5, "base": 1, "t": 2, "ctr": 3, "b": 4}
_MIXED = VBAInt(-2, "Long")
_INVALID = (-2147024809, "The specified value is out of range.")
#: pPr's attributes and children, in the schema's order.
_ATTRIBUTES = ("marL", "marR", "lvl", "indent", "algn", "defTabSz", "rtl", "eaLnBrk", "fontAlgn", "latinLnBrk", "hangingPunct")
_CHILDREN = ("a:lnSpc", "a:spcBef", "a:spcAft", "a:buClrTx", "a:buClr", "a:buSzTx", "a:buSzPct", "a:buSzPts", "a:buFontTx", "a:buFont",
             "a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip", "a:tabLst", "a:defRPr", "a:extLst")
_CHILD = re.compile(r"<(a:\w+)\b[^>]*?(?:/>|>.*?</\1>)", re.DOTALL)


def _children(markup: str) -> dict[str, str]:
    """A pPr's children, by tag."""
    start = markup.find(">")
    if start < 0 or markup[start - 1] == "/":
        return {}
    return {found.group(1): found.group(0) for found in _CHILD.finditer(markup[start + 1:markup.rfind("</a:pPr>")])}


def _written(attributes: dict[str, str], children: dict[str, str]) -> str:
    """A pPr, in the schema's order."""
    head = "".join(f' {name}="{attributes[name]}"' for name in _ATTRIBUTES if name in attributes)
    inside = "".join(children[tag] for tag in _CHILDREN if tag in children)
    if inside:
        return f"<a:pPr{head}>{inside}</a:pPr>"
    return f"<a:pPr{head}/>" if head else ""


def _plain(shape: Shape) -> str:
    look = shape.text_look
    if look is None:
        return ""
    return _written({**({"algn": look.align} if look.align else {}), **({} if look.rtl is None else {"rtl": "1" if look.rtl else "0"})}, {})


def _all_formats(shape: Shape) -> list[str]:
    return list(shape.text_paragraphs) if shape.text_paragraphs is not None else [_plain(shape)] * (shape.text.count("\n") + 1)


def _settled(shape: Shape, formats: list[str]) -> None:
    """Keep the paragraphs' settings, in the look where it can hold them."""
    look = shape.text_look
    assert look is not None
    first = paragraph_settings(formats[0]) if formats else {}
    look.align = first.get("algn")
    look.rtl = None if "rtl" not in first else first["rtl"] == "1"
    shape.text_paragraphs = None if all(one == _plain(shape) for one in formats) else formats


def paragraphs_set(shape: Shape, name: str, value: str) -> None:
    """A setting TextFrame gives every paragraph."""
    if shape.text_paragraphs is not None:
        shape.text_paragraphs = [_written({**paragraph_settings(one), name: value}, _children(one)) for one in shape.text_paragraphs]


def _tri(value: bool) -> object:
    return VBAInt(-1 if value else 0, "Long")


class ParagraphFormat2(HeldPart):
    """The paragraphs a TextRange2 reaches into, each with its pPr. Where they differ it reads as mixed."""

    vba_type_name = "ParagraphFormat2"

    def __init__(self, text: TextRange2) -> None:
        self.text = text

    def _range(self) -> tuple[int, int]:
        """Which paragraphs this range touches: the first, and how many."""
        text = self.text.owner.shape.text
        at, count = self.text.span()
        inside = text[at:at + count]
        return text[:at].count("\n"), inside.count("\n") + 1 - (1 if inside.endswith("\n") else 0)

    def _formats(self) -> list[str]:
        first, many = self._range()
        return _all_formats(self.text.owner.shape)[first:first + many]

    def _attribute(self, name: str) -> tuple[bool, str | None]:
        """Whether the paragraphs differ in an attribute, and what it is."""
        found = {paragraph_settings(one).get(name) for one in self._formats()}
        return len(found) > 1, next(iter(found))

    def _change(self, name: str | None, value: str | None, tag: str | None = None, child: str = "") -> None:
        """An attribute or a child of each paragraph here, set or, with None, gone."""
        from pyopenvba.apps.excel._shapes import TextFrame

        shape = self.text.owner.shape
        TextFrame(self.text.owner)._changed()  # pyright: ignore[reportPrivateUsage]
        formats = _all_formats(shape)
        first, many = self._range()
        for place in range(first, first + many):
            attributes, children = paragraph_settings(formats[place]), _children(formats[place])
            if name is not None:
                if value is None:
                    attributes.pop(name, None)
                else:
                    attributes[name] = value
            if tag is not None:
                children[tag] = child
            formats[place] = _written(attributes, children)
        _settled(shape, formats)

    @member
    def Application(self) -> object:
        return self.text.owner.sheet.book.application

    @member
    def Parent(self) -> object:
        return self.text

    @member
    def Alignment(self) -> object:
        mixed, algn = self._attribute("algn")
        if mixed:
            return _MIXED
        if algn is None and self._attribute("rtl")[1] == "1":
            # With none named, right to left reads as right-aligned.
            return VBAInt(3, "Long")
        return VBAInt(_ALIGNMENTS.get(algn or "l", 1), "Long")

    @setter("Alignment")
    def _alignment(self, value: object) -> None:
        wanted = next((name for name, number in _ALIGNMENTS.items() if number == int(to_number(value))), None)
        if wanted is None:
            raise error(*_INVALID)
        self._change("algn", wanted)

    @member
    def BaselineAlignment(self) -> object:
        mixed, chosen = self._attribute("fontAlgn")
        return _MIXED if mixed else VBAInt(_FONT_ALIGNS.get(chosen or "auto", 5), "Long")

    def _indent(self, name: str, unset: float) -> object:
        mixed, chosen = self._attribute(name)
        if mixed:
            raise VBAUnsupportedError("Paragraphs differing in their indents are not implemented")
        return VBASingle(unset if chosen is None else int(chosen) / 12700)

    def _set_indent(self, name: str, value: object) -> None:
        self._change(name, str(round(float(VBASingle(float(to_number(value)))) * 12700)))

    @member
    def FirstLineIndent(self) -> object:
        return self._indent("indent", 0.0)

    @setter("FirstLineIndent")
    def _first_line_indent(self, value: object) -> None:
        self._set_indent("indent", value)

    @member
    def LeftIndent(self) -> object:
        # With none named, each level goes 36 points further in.
        level = self._attribute("lvl")[1]
        return self._indent("marL", 36.0 * int(level or 0))

    @setter("LeftIndent")
    def _left_indent(self, value: object) -> None:
        self._set_indent("marL", value)

    @member
    def RightIndent(self) -> object:
        return self._indent("marR", 0.0)

    @setter("RightIndent")
    def _right_indent(self, value: object) -> None:
        self._set_indent("marR", value)

    def _space(self, tag: str) -> tuple[bool, float | None, bool | None]:
        """Whether the paragraphs differ in a spacing, and what it is: its size, and whether in lines; None if unsaid."""
        found: set[str | None] = {_children(one).get(tag) for one in self._formats()}
        one = next(iter(found))
        if len(found) > 1 or one is None:
            return len(found) > 1, None, None
        points = re.search(r'<a:spcPts val="(-?\d+)"', one)
        if points is not None:
            return False, int(points.group(1)) / 100, False
        lines = re.search(r'<a:spcPct val="(-?\d+)"', one)
        if lines is not None:
            return False, int(lines.group(1)) / 100000, True
        raise VBAUnsupportedError("This spacing of a paragraph in a file is not implemented")

    def _spacing(self, tag: str, unset: float) -> object:
        mixed, size, _ = self._space(tag)
        if mixed:
            raise VBAUnsupportedError("Paragraphs differing in their spacing are not implemented")
        return VBASingle(unset if size is None else size)

    def _rule(self, tag: str, unset: bool) -> object:
        mixed, _, lines = self._space(tag)
        if mixed:
            raise VBAUnsupportedError("Paragraphs differing in their spacing are not implemented")
        return _tri(unset if lines is None else lines)

    def _set_spacing(self, tag: str, value: object, in_lines: bool) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        inner = f'<a:spcPct val="{round(wanted * 100000)}"/>' if in_lines else f'<a:spcPts val="{round(wanted * 100)}"/>'
        self._change(None, None, tag, f"<{tag}>{inner}</{tag}>")

    @member
    def SpaceAfter(self) -> object:
        return self._spacing("a:spcAft", 0.0)

    @setter("SpaceAfter")
    def _space_after(self, value: object) -> None:
        self._set_spacing("a:spcAft", value, bool(self._space("a:spcAft")[2]))

    @member
    def SpaceBefore(self) -> object:
        return self._spacing("a:spcBef", 0.0)

    @setter("SpaceBefore")
    def _space_before(self, value: object) -> None:
        self._set_spacing("a:spcBef", value, bool(self._space("a:spcBef")[2]))

    @member
    def SpaceWithin(self) -> object:
        return self._spacing("a:lnSpc", 1.0)

    @setter("SpaceWithin")
    def _space_within(self, value: object) -> None:
        lines = self._space("a:lnSpc")[2]
        self._set_spacing("a:lnSpc", value, lines is None or lines)

    @member
    def LineRuleAfter(self) -> object:
        return self._rule("a:spcAft", False)

    @member
    def LineRuleBefore(self) -> object:
        return self._rule("a:spcBef", False)

    @member
    def LineRuleWithin(self) -> object:
        return self._rule("a:lnSpc", True)

    def _flag(self, name: str) -> object:
        mixed, chosen = self._attribute(name)
        return _MIXED if mixed else _tri(chosen != "0")

    @member
    def WordWrap(self) -> object:
        # The file's says to break inside Latin words.
        mixed, chosen = self._attribute("latinLnBrk")
        return _MIXED if mixed else _tri(chosen != "1")

    @member
    def FarEastLineBreakLevel(self) -> object:
        return self._flag("eaLnBrk")

    @member
    def HangingPunctuation(self) -> object:
        return self._flag("hangingPunct")

    @member
    def TextDirection(self) -> object:
        mixed, chosen = self._attribute("rtl")
        if mixed:
            return _MIXED
        return VBAInt({"0": 1, "1": 2}.get(chosen or "", -2), "Long")

    @setter("TextDirection")
    def _text_direction(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in (1, 2):
            raise error(*_INVALID)
        self._change("rtl", "1" if wanted == 2 else "0")

    @member
    def IndentLevel(self) -> object:
        mixed, chosen = self._attribute("lvl")
        return _MIXED if mixed else VBAInt(int(chosen or 0) + 1, "Long")

    @setter("IndentLevel")
    def _indent_level(self, value: object) -> None:
        wanted = int(to_number(value))
        if not 1 <= wanted <= 9:
            raise error(*_INVALID)
        self._change("lvl", str(wanted - 1))

    @member
    def Bullet(self) -> object:
        return Bullet2(self)


class Bullet2(HeldPart):
    """A paragraph's bullet: whether it has one. Turned on it is numbered, standing out in the paragraph's own margin."""

    vba_type_name = "BulletFormat2"

    def __init__(self, paragraph: ParagraphFormat2) -> None:
        self.paragraph = paragraph

    @member
    def Parent(self) -> object:
        return self.paragraph

    def _shown(self) -> set[bool]:
        return {any(tag in _children(one) for tag in ("a:buAutoNum", "a:buChar", "a:buBlip"))
                for one in self.paragraph._formats()}  # pyright: ignore[reportPrivateUsage]

    @member
    def Visible(self) -> object:
        # Any paragraph with one makes it visible.
        return _tri(True in self._shown())

    @setter("Visible")
    def _visible(self, value: object) -> None:
        if not to_number(value):
            if True in self._shown():
                raise VBAUnsupportedError("Taking a paragraph's bullet off is not implemented")
            return
        if self._shown() == {True}:
            return
        paragraph = self.paragraph
        paragraph._change("marL", "228600")  # pyright: ignore[reportPrivateUsage]
        paragraph._change("indent", "-228600", "a:buFont", '<a:buFont typeface="+mj-lt"/>')  # pyright: ignore[reportPrivateUsage]
        paragraph._change(None, None, "a:buAutoNum", '<a:buAutoNum type="arabicPeriod"/>')  # pyright: ignore[reportPrivateUsage]


class NoEffect(HeldPart):
    """Font2's Shadow, Glow, Reflection and the like, on letters with none: the file's effects are not read."""

    vba_type_name = ""

    def __init__(self, font: object, type_name: str, answers: dict[str, object]) -> None:
        self.font = font
        self.vba_type_name = type_name
        self.answers = answers

    @member
    def Parent(self) -> object:
        return self.font

    def vba_get(self, name: str, args: object = (), named: object = None) -> object:
        for key, value in self.answers.items():
            if key.lower() == name.lower():
                return value
        return super().vba_get(name, args, named)  # pyright: ignore[reportArgumentType]
