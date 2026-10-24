"""Number format codes: as Range.NumberFormat spells them, and as the file does.

Setting Range.NumberFormat does not keep the code as written: Excel
parses it and writes it back out. What scripts/measure_typing_formats.py
and scripts/measure_format_codes.py measured of that:

- a quoted ``"$"`` becomes a bare ``$``, while ``"-"``, ``" "`` and other
  quoted text stay quoted; an escaped character becomes bare unless bare
  it would mean something -- ``\\-``, ``\\(``, ``\\T`` and ``\\1`` lose
  the backslash, while ``\\$``, ``\\%``, ``\\@``, ``\\0``, ``\\E`` and
  the other date and era letters keep it (scripts/measure_format_escapes.py);
- date and time codes are lower case, a lone ``b`` or ``y`` doubled,
  ``General`` and colour names are capitalised, and ``am/pm`` becomes
  ``AM/PM`` while ``a/p`` stays; a locale tag Excel knows by its LCID
  takes it, ``[$-en-US]`` reading ``[$-409]``;
- a character shown as it is, that bare would mean something -- ``A``,
  ``R``, ``]``, or ``:`` outside a date -- reads back escaped, and an
  accented letter is read as the letter under its accent;
- a code Excel cannot read is refused with error 1004: a lower-case
  exponent, ``0.00e+00``, a section mixing a number's, a date's and text's
  characters, ``0 m``, and the rest :func:`_refused` names.

The file spells a code another way: every character shown as it is
escaped with a backslash, the currency symbol NumberFormat shows as a
bare ``$`` quoted as ``"$"``, and a bare ``$`` in a file a plain dollar
sign, which NumberFormat shows as ``\\$``. Which characters a code shows
as they are depends on its section: a number's digits, ``#``, ``?``,
``.``, ``%`` and ``/`` are its own, a comma after a digit, and the digits
of a fraction's denominator; a date's codes, ``/``, ``:`` and a fraction
of a second; a text section's ``@`` (scripts/measure_conditional_number_formats.py,
tests/fixtures/conditional_number_formats/, every printable character a
rule and a cell were given, bare and escaped, in each kind of section).
Excel's built-in formats have ids of their own and are not written out,
except the currency and accounting ones, 5 to 8 and 41 to 44, whose code
Excel spells out in the file as well. A conditional format's numFmt
always holds its code, and numbers a built-in format by its place in the
list Excel's Custom category shows (LISTED) instead.

Excel also moves a thousands separator when a quoted literal splits the
digits -- ``#,##0"."00`` comes back ``###,0"."00`` -- which is not
modelled; such a code is kept as written.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Callable
from functools import lru_cache
from typing import Final

from pyopenvba.formula._display import tokens
from pyopenvba.interpreter._values import error

_COLOURS = {name.lower(): name for name in ("Black", "Blue", "Cyan", "Green", "Magenta", "Red", "White", "Yellow")}
_ELAPSED = re.compile(r"[hms]+", re.IGNORECASE)
#: A condition's bracket, [>100] or [<=-1.5].
_CONDITION = re.compile(r"[<>=]{1,2}-?[0-9.]+(?:[eE][+-]?[0-9]+)?")
#: The other brackets Excel reads: a locale or currency, an elapsed time, a palette colour, a numeral system.
_BRACKETS = re.compile(r"\$.*|[hms]+|color[0-9]+|dbnum[0-9]|natnum[0-9]+", re.IGNORECASE | re.DOTALL)
#: Characters an escape stays on, since bare they would be part of the format: the rest lose it.
_KEPT_ESCAPED = frozenset('"#$%*/0:;?@[\\]_ABDEGHMNRSYabdeghmnrsy')
#: The letters a date or time section reads as its codes, in either case.
_DATE_LETTERS = frozenset("dmyhsegbDMYHSEGB")
#: Characters a text section refuses bare, a number's and a date's; and those a date refuses, a number's but the
#: zeros of a fraction of a second.
_NOT_TEXT = frozenset("0123456789#?%./") | _DATE_LETTERS
_NOT_DATE = frozenset("123456789#?%")
#: Letters outside ASCII Excel reads as a letter of its own though Unicode gives them no parts to fold to.
_FOLDED = {"Ð": "D", "ð": "D", "Đ": "D", "đ": "D", "Ħ": "H", "ħ": "H", "ŉ": "N"}
#: Locale tags Excel spells by their LCID once it has read them.
_LOCALES = {"x-sysdate": "F800", "x-systime": "F400", "en-US": "409", "x-euro2": "2"}
#: Words Excel reads as one piece: the General format, and the morning and afternoon markers.
_WORDS = (("general", "general"), ("am/pm", "ampm"), ("a/p", "ampm"))


def _word_at(listed: list[tuple[str, str]], index: int) -> tuple[str, int] | None:
    """The word Excel reads whole that the bare characters from ``index`` spell, and where it ends."""
    for name, part in _WORDS:
        chunk = listed[index:index + len(name)]
        if len(chunk) == len(name) and all(kind == "char" for kind, _ in chunk) and \
                "".join(text for _, text in chunk).lower() == name:
            return part, index + len(name)
    return None


def _folded(char: str) -> str:
    """The letter Excel reads a character as: a letter outside ASCII the letter it is made from, in upper case,
    so that ``Ý`` in a date is its year and ``é`` the exponent ``E`` (scripts/measure_number_format_refusals.py)."""
    if char.isascii() or not char.isalpha():
        return char
    if char in _FOLDED:
        return _FOLDED[char]
    base = unicodedata.normalize("NFD", char)[0]
    return base.upper() if base != char and base.isascii() else char


def _sections(code: str) -> list[tuple[str, list[tuple[str, str]]]]:
    """A code's sections, each with its kind, "number", "date" or "text", and its pieces: those ``tokens`` gives,
    each bare character given its part, ("syntax", c) for one the format reads and ("literal", c) for one it shows
    as it is, and ("currency", "$"), ("general", word) and ("ampm", word) for the pieces Excel reads whole."""
    listed = list(tokens(code))
    merged: list[tuple[str, str]] = []
    index = 0
    while index < len(listed):
        word = _word_at(listed, index)
        if word is None:
            merged.append(listed[index])
            index += 1
        else:
            merged.append((word[0], "".join(text for _, text in listed[index:word[1]])))
            index = word[1]
    grouped: list[list[tuple[str, str]]] = [[]]
    for piece in merged:
        if piece == ("char", ";"):
            grouped.append([])
        else:
            grouped[-1].append(piece)
    out: list[tuple[str, list[tuple[str, str]]]] = []
    for members in grouped:
        if ("char", "@") in members:
            kind = "text"
        elif any(part == "ampm" or (part == "char" and _folded(text) in _DATE_LETTERS)
                 or (part == "bracket" and _ELAPSED.fullmatch(text[1:-1])) for part, text in members):
            kind = "date"
        else:
            kind = "number"
        out.append((kind, _parts(kind, members)))
    return out


def _parts(kind: str, members: list[tuple[str, str]]) -> list[tuple[str, str]]:
    """A section's pieces, each bare character given its part in a section of ``kind``; a date's code is the
    letter Excel reads it as."""
    out: list[tuple[str, str]] = []
    for at, (part, text) in enumerate(members):
        if part != "char":
            out.append((part, text))
            continue
        before = out[-1] if out else ("", "")
        after = members[at + 1] if at + 1 < len(members) else ("", "")
        if text == "$":
            role = "currency"
        elif kind == "text":
            role = "syntax" if text == "@" else "literal"
        elif kind == "date":
            if _folded(text) in _DATE_LETTERS:
                text, role = _folded(text), "syntax"
            elif text in "/:":
                role = "syntax"
            elif text == ".":
                # A fraction of a second: a point before a zero, after any of the date's codes.
                role = "syntax" if after == ("char", "0") else "literal"
            elif text == "0":
                role = "syntax" if before in (("syntax", "."), ("syntax", "0")) else "literal"
            else:
                role = "literal"
        elif text in "0#?.%/":
            role = "syntax"
        elif text == ",":
            role = "syntax" if before[0] == "syntax" and before[1] in "0#?," else "literal"
        elif text in "123456789":
            # A fraction's denominator: the digits after its stroke.
            role = "syntax" if before[0] == "syntax" and (before[1] == "/" or before[1] in "123456789") else "literal"
        else:
            role = "literal"
        out.append((role, text))
    return out


def _pieces(code: str) -> list[tuple[str, str]]:
    """A code's pieces, section after section, each section closed by ("syntax", ";")."""
    out: list[tuple[str, str]] = []
    for place, (_, pieces) in enumerate(_sections(code)):
        if place:
            out.append(("syntax", ";"))
        out.extend(pieces)
    return out


def _known_bracket(text: str) -> bool:
    """Whether Excel reads a bracket: closed, and a colour, a condition, a locale or currency, an elapsed time or a
    numeral system."""
    inner = text[1:-1]
    return text.endswith("]") and (inner.lower() in _COLOURS or _CONDITION.fullmatch(inner) is not None
                                   or _BRACKETS.fullmatch(inner) is not None)


def _refused(sections: list[tuple[str, list[tuple[str, str]]]]) -> bool:
    """Whether Excel refuses a code with error 1004, for a cell's NumberFormat and a rule's alike.

    Measured over every printable character in each kind of section and
    over whole codes (scripts/measure_conditional_number_formats.py and
    measure_number_format_refusals.py): more than four sections, a text
    section before another, or a fourth section that is not text; and in
    any section a bracket left open or not Excel's, a fill, a skip or an
    escape with nothing to act on, an ``E`` outside an exponent, or an
    ``N``.
    """
    if len(sections) > 4:
        return True
    return any((kind == "text" and place < len(sections) - 1)
               or _section_refused("text" if place == 3 else kind, pieces)
               for place, (kind, pieces) in enumerate(sections))


def _section_refused(kind: str, pieces: list[tuple[str, str]]) -> bool:
    """Whether Excel refuses one section of a code: what :func:`_refused` names, and a section mixing kinds. A
    text section takes none of a number's or a date's characters; a date none of a number's but the zeros of a
    fraction of a second, three at most; General no digits and no text; and a number an exponent and a fraction's
    stroke only between digits, and no colon before one."""
    bare = [(part, text) for part, text in pieces if part in ("syntax", "literal")]
    if any(_folded(text) in ("E", "N", "n") or text in ("*", "_", "\\") for _, text in bare):
        return True
    if any(part == "bracket" and not _known_bracket(text) for part, text in pieces):
        return True
    digits = [at for at, (part, text) in enumerate(pieces) if part == "syntax" and text in "0#?"]
    exponents = [at for at, (part, _) in enumerate(pieces) if part == "exponent"]
    if exponents and (kind != "number" or len(exponents) > 1 or pieces[exponents[0]][1][0] == "e"
                      or not digits or digits[0] > exponents[0] or digits[-1] < exponents[0]):
        return True
    if any(part == "general" for part, _ in pieces) and (kind == "text" or digits):
        return True
    if kind == "text":
        # Nor a condition: [>5]@ is refused (tests/fixtures/text_cells.json).
        return any(_folded(text) in _NOT_TEXT for _, text in bare) or any(
            part == "bracket" and _CONDITION.fullmatch(text[1:-1]) for part, text in pieces)
    if kind == "date":
        run = 0
        for part, text in pieces:
            run = run + 1 if (part, text) == ("syntax", "0") else 0
            if run > 3 or (part in ("syntax", "literal") and text in _NOT_DATE) or (part, text) == ("literal", "0"):
                return True
        return False
    for at, piece in enumerate(pieces):
        if piece == ("syntax", "/") and not (any(one < at for one in digits) and any(
                part == "syntax" and (text in "0#?" or text.isdigit()) for part, text in pieces[at + 1:])):
            return True
        if piece == ("literal", ":") and any(one > at for one in digits):
            return True
    return False


def _bracket(text: str) -> str:
    """A bracket as Excel keeps it: a colour's name capitalised, an elapsed code lower case, and a locale tag
    Excel knows by its LCID taking it."""
    inner = text[1:-1]
    if inner.lower() in _COLOURS:
        return f"[{_COLOURS[inner.lower()]}]"
    if _ELAPSED.fullmatch(inner):
        return f"[{inner.lower()}]"
    found = re.fullmatch(r"\$([^-]*)-(.+)", inner)
    if found is not None and found.group(2) in _LOCALES:
        return f"[${found.group(1)}-{_LOCALES[found.group(2)]}]"
    return text

#: Excel's built-in formats on a US English system, by id, as NumberFormat spells them.
BUILTIN: Final[dict[int, str]] = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    5: "$#,##0_);($#,##0)", 6: "$#,##0_);[Red]($#,##0)", 7: "$#,##0.00_);($#,##0.00)",
    8: "$#,##0.00_);[Red]($#,##0.00)", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
    20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm",
    37: "#,##0_);(#,##0)", 38: "#,##0_);[Red](#,##0)", 39: "#,##0.00_);(#,##0.00)", 40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', 42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', 44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
}
#: The built-in formats Excel still spells out in the file's numFmts.
SPELLED_OUT: Final = frozenset({5, 6, 7, 8, 41, 42, 43, 44})
#: The id a conditional format's numFmt gives each built-in format: its place in the list Excel's Custom category
#: shows, not the id a cell's xf gives it (tests/fixtures/conditional_number_formats/).
LISTED: Final[dict[str, int]] = {BUILTIN[number]: place for place, number in enumerate(
    (0, 1, 2, 3, 4, 37, 38, 39, 40, 5, 6, 7, 8, 9, 10, 11, 48, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
     45, 47, 49, 46, 42, 41, 44, 43))}


def _rewritten(code: str, change: Callable[[str, str], str]) -> str:
    return "".join(change(kind, text) for kind, text in tokens(code))


def normalized(code: str, *, strict: bool = True) -> str:
    """A number format as Excel keeps it once Range.NumberFormat is set to ``code``.

    ``strict`` refuses what Excel refuses to set; a code read from a file
    is taken as it comes. A section holding a colour and nothing else
    gains General, and a code of one section with a condition gains a
    second, General: ``[Red]`` reads ``[Red]General`` and ``[>0]0`` reads
    ``[>0]0;General`` (tests/fixtures/conditional_number_formats/refusals.json).
    """
    sections = _sections(code)
    if strict and _refused(sections):
        raise error(1004, "Unable to set the NumberFormat property of the Range class")
    written = [_spelled(pieces) for _, pieces in sections]
    for place, (_, pieces) in enumerate(sections):
        if pieces and all(part == "bracket" and text[1:-1].lower() in _COLOURS for part, text in pieces):
            written[place] += "General"
    if len(sections) == 1 and any(part == "bracket" and _CONDITION.fullmatch(text[1:-1])
                                  for part, text in sections[0][1]):
        written.append("General")
    if len(sections) == 2 and not sections[0][1]:
        # An empty first section of two gains an empty third, ;0 reading ;0; (conditional_number_formats, chars),
        # or with text the third and fourth, ;@ reading ;;;@ (tests/fixtures/text_cells.json).
        written = ["", "", "", written[1]] if sections[1][0] == "text" else [*written, ""]
    return ";".join(written)


def _spelled(pieces: list[tuple[str, str]]) -> str:
    """A section's pieces as NumberFormat spells them."""
    # Whether the section has a digit placeholder, which an escaped dot needs to stay escaped.
    digits = any(kind == "syntax" and text in "0#?" for kind, text in pieces)
    out: list[str] = []
    previous = ""
    index = 0
    while index < len(pieces):
        kind, text = pieces[index]
        after, previous = previous, text[-1] if kind in ("syntax", "literal", "currency", "general", "ampm") else ""
        index += 1
        if kind == "syntax" and text in _DATE_LETTERS:
            # A date code takes lower case, a lone b or y doubled.
            letter, count = text.lower(), 1
            while index < len(pieces) and pieces[index][0] == "syntax" and pieces[index][1].lower() == letter:
                count, index = count + 1, index + 1
            out.append(letter * (2 if letter in "by" and count == 1 else count))
            previous = letter
        elif kind == "syntax":
            out.append(text)
        elif kind == "literal":
            out.append("\\" + text if _folded(text) in _KEPT_ESCAPED else text)
        elif kind == "currency":
            out.append("$")
        elif kind == "general":
            out.append("General")
        elif kind == "ampm":
            # AM/PM takes the case Excel writes it in; a/p keeps its own.
            out.append("AM/PM" if text.lower() == "am/pm" else text)
        elif kind == "quoted":
            out.append("$" if text == '"$"' else text)
        elif kind == "escaped":
            out.append(_escaped(text[1], after, digits))
        elif kind == "bracket":
            out.append(_bracket(text))
        else:
            out.append(text)
    return "".join(out)


def _escaped(char: str, after: str, digits: bool) -> str:
    """An escaped character as NumberFormat spells it: bare unless bare it would mean something.

    Measured over every printable character (scripts/measure_format_escapes.py).
    A comma stays escaped straight after a digit placeholder, where bare it
    would scale the number, and a dot in a section with digits, where it
    would be the decimal point.
    """
    if _folded(char) in _KEPT_ESCAPED:
        return "\\" + char
    if (char == "," and after in ("0", "#", "?")) or (char == "." and digits):
        return "\\" + char
    return char


@lru_cache(maxsize=1024)
def from_file(code: str) -> str:
    """A code as a file spells it, as Range.NumberFormat reads it; kept, as a rule's is read for every cell it
    shows."""
    return normalized(_rewritten(code, lambda kind, text: "\\$" if kind == "char" and text == "$" else text),
                      strict=False)


def to_file(code: str) -> str:
    """A code as Range.NumberFormat spells it, as the file spells it: every character shown as it is escaped,
    and the currency symbol quoted."""
    out: list[str] = []
    for kind, text in _pieces(code):
        if kind == "literal":
            out.append("\\" + text)
        elif kind == "currency":
            out.append('"$"')
        else:
            out.append(text)
    return "".join(out)
