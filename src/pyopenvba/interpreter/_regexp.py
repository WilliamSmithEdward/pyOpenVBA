"""VBScript.RegExp, which CreateObject makes without starting the VBScript engine.

Measured in live Excel (scripts/measure_scripting_objects.py,
tests/fixtures/scripting_objects.json). The dialect is JavaScript's of its
day, which Python's ``re`` runs once a pattern is put in its terms:

- ``\\d``, ``\\w``, ``\\s`` and ``\\b`` are ASCII's: no Arabic digit is a
  digit, no accented letter a word character and a no-break space no
  space; IgnoreCase folds every letter, an accented capital with its small
  letter.
- ``^`` and ``$`` are the text's ends, or every line's under MultiLine, a
  carriage return ending a line as a line feed does; ``.`` is anything but
  a line feed.
- ``{`` is a quantifier only as ``{n}``, ``{n,}`` or ``{n,m}``, and is text
  otherwise: ``a{,2}`` matches itself. An escape of a letter that means
  nothing is the letter.
- Groups are ``(...)``, ``(?:...)`` and the lookaheads ``(?=...)`` and
  ``(?!...)``; a lookbehind is a syntax error, 5017. A ``(`` left open is
  error 5020, a ``[`` 5019, a quantifier with nothing to repeat 5018, a
  range backwards 5021. A pattern is checked where it is used, not where
  it is set.
- Replace writes ``$1`` to ``$99`` as the groups, a group that took no
  part as nothing and one past the last as the text written, ``$&`` the
  match, ``$``` what comes before it, ``$'`` what after and ``$$`` a dollar.
  Without Global, Replace and Execute stop at the first match. An empty
  match is found at every place, as JavaScript finds one.
- A match's FirstIndex counts from 0; a sub-match that took no part is
  Empty. Collections count from 0, and an item past their end is error 5.
  Null where a text is wanted is error 13, and a number is its text.
"""

from __future__ import annotations

import re
from collections.abc import Iterator

from pyopenvba.interpreter._objects import VBAObject, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAInt, error, to_bool, to_integer, to_text

#: VBScript's errors for a pattern it cannot read, by number.
_ERRORS = {5017: "Syntax error in regular expression", 5018: "Unexpected quantifier",
           5019: "Expected ']' in regular expression", 5020: "Expected ')' in regular expression",
           5021: "Invalid range in character set"}
#: The ASCII classes the shorthands stand for, inside a character class.
_INSIDE = {"d": "0-9", "w": "A-Za-z0-9_", "s": " \\t\\n\\r\\f\\v"}
#: The escapes that mean the same in Python's patterns.
_KEPT = set("tnrfv0123456789") | {"x", "u"}
#: What ends a line for ^ and $ under MultiLine.
_LINE_END = "\\n\\r\\u2028\\u2029"
_QUANTIFIER = re.compile(r"\{\d+(?:,\d*)?\}")


class _PatternError(Exception):
    def __init__(self, number: int) -> None:
        super().__init__(_ERRORS[number])
        self.number = number


def translated(pattern: str, *, multiline: bool) -> str:
    """A VBScript pattern in Python's terms (see the module's notes)."""
    out: list[str] = []
    index, depth = 0, 0
    # Whether what was just written can take a quantifier.
    repeatable = False
    while index < len(pattern):
        char = pattern[index]
        if char == "\\":
            if index + 1 >= len(pattern):
                raise _PatternError(5017)
            escaped = pattern[index + 1]
            index += 2
            if escaped in "dwsDWS":
                # ASCII's classes, whatever the rest of the pattern does with case.
                out.append(f"(?a:\\{escaped})")
            elif escaped in "bB":
                out.append(f"(?a:\\{escaped})")
                repeatable = False
                continue
            elif escaped == "c" and index < len(pattern) and pattern[index].isalpha():
                out.append(re.escape(chr(ord(pattern[index]) % 32)))
                index += 1
            elif escaped in _KEPT:
                out.append("\\" + escaped)
            else:
                out.append(re.escape(escaped))
            repeatable = True
            continue
        if char == "[":
            index, written = _class(pattern, index)
            out.append(written)
            repeatable = True
            continue
        if char == "(":
            if pattern.startswith("(?", index):
                kind = pattern[index + 2:index + 3]
                if kind not in (":", "=", "!"):
                    raise _PatternError(5017)
                out.append(pattern[index:index + 3])
                index += 3
            else:
                out.append("(")
                index += 1
            depth += 1
            repeatable = False
            continue
        if char == ")":
            if not depth:
                raise _PatternError(5017)
            depth -= 1
            out.append(")")
            index += 1
            repeatable = True
            continue
        if char in "*+?":
            if not repeatable:
                raise _PatternError(5018)
            out.append(char)
            index += 1
            if index < len(pattern) and pattern[index] == "?":
                out.append("?")
                index += 1
            repeatable = False
            continue
        if char == "{":
            found = _QUANTIFIER.match(pattern, index)
            if found is None:
                out.append("\\{")
                index += 1
                repeatable = True
                continue
            if not repeatable:
                raise _PatternError(5018)
            out.append(found.group(0))
            index = found.end()
            if index < len(pattern) and pattern[index] == "?":
                out.append("?")
                index += 1
            repeatable = False
            continue
        if char == "^":
            out.append(f"(?:(?<=[{_LINE_END}])|\\A)" if multiline else "\\A")
            index += 1
            repeatable = False
            continue
        if char == "$":
            out.append(f"(?=[{_LINE_END}]|\\Z)" if multiline else "\\Z")
            index += 1
            repeatable = False
            continue
        if char == "|":
            out.append("|")
            index += 1
            repeatable = False
            continue
        out.append(char if char == "." else re.escape(char))
        index += 1
        repeatable = True
    if depth:
        raise _PatternError(5020)
    return "".join(out)


def _class(pattern: str, start: int) -> tuple[int, str]:
    """A character class from its [, as Python writes it, and where the pattern goes on after it."""
    index = start + 1
    negated = index < len(pattern) and pattern[index] == "^"
    if negated:
        index += 1
    if index < len(pattern) and pattern[index] == "]":
        # [] matches nothing and [^] anything, as JavaScript reads them.
        return index + 1, "[\\s\\S]" if negated else "(?!)"
    parts: list[str] = []
    while index < len(pattern) and pattern[index] != "]":
        char = pattern[index]
        if char == "\\":
            if index + 1 >= len(pattern):
                raise _PatternError(5019)
            escaped = pattern[index + 1]
            index += 2
            if escaped in _INSIDE:
                parts.append(_INSIDE[escaped])
            elif escaped in "DWS":
                parts.append("\\" + escaped)
            elif escaped == "b":
                parts.append("\\x08")
            elif escaped == "c" and index < len(pattern) and pattern[index].isalpha():
                parts.append(re.escape(chr(ord(pattern[index]) % 32)))
                index += 1
            elif escaped in _KEPT:
                parts.append("\\" + escaped)
            else:
                parts.append(re.escape(escaped))
            continue
        # Python reads [, &&, ~~, || and -- in a set as set operations to come; here they are characters.
        parts.append("\\" + char if char in "[&~|" or (char == "-" and parts and parts[-1] == "-") else char)
        index += 1
    if index >= len(pattern):
        raise _PatternError(5019)
    return index + 1, "[" + ("^" if negated else "") + "".join(parts) + "]"


def _text(value: object) -> str:
    if value is NULL:
        raise error(13, "Type mismatch")
    return to_text(value)


class RegExp(VBAObject):
    """VBScript.RegExp."""

    vba_type_name = "RegExp"

    def __init__(self) -> None:
        self._pattern = ""
        self._global = False
        self._ignore_case = False
        self._multiline = False

    def _compiled(self) -> re.Pattern[str]:
        try:
            written = translated(self._pattern, multiline=self._multiline)
        except _PatternError as exc:
            raise error(exc.number, str(exc)) from None
        try:
            return re.compile(written, re.IGNORECASE if self._ignore_case else 0)
        except re.error as exc:
            message = str(exc)
            number = 5021 if "bad character range" in message else 5018 if "repeat" in message else 5017
            raise error(number, _ERRORS[number]) from None

    def _matches(self, text: str) -> list[re.Match[str]]:
        compiled = self._compiled()
        if self._global:
            return list(compiled.finditer(text))
        found = compiled.search(text)
        return [] if found is None else [found]

    @member
    def Pattern(self) -> object:
        return self._pattern

    @setter("Pattern")
    def _set_pattern(self, value: object) -> None:
        self._pattern = _text(value)

    @member
    def Global(self) -> object:
        return self._global

    @setter("Global")
    def _set_global(self, value: object) -> None:
        self._global = to_bool(value)

    @member
    def IgnoreCase(self) -> object:
        return self._ignore_case

    @setter("IgnoreCase")
    def _set_ignore_case(self, value: object) -> None:
        self._ignore_case = to_bool(value)

    @member
    def MultiLine(self) -> object:
        return self._multiline

    @setter("MultiLine")
    def _set_multiline(self, value: object) -> None:
        self._multiline = to_bool(value)

    @method
    def Test(self, sourceString: object = MISSING) -> object:
        if sourceString is MISSING:
            raise error(449)
        return self._compiled().search(_text(sourceString)) is not None

    @method
    def Execute(self, sourceString: object = MISSING) -> object:
        if sourceString is MISSING:
            raise error(449)
        return MatchCollection([Match(found) for found in self._matches(_text(sourceString))])

    @method
    def Replace(self, sourceString: object = MISSING, replaceVar: object = MISSING) -> object:
        if sourceString is MISSING or replaceVar is MISSING:
            raise error(449)
        text, replacement = _text(sourceString), _text(replaceVar)
        out: list[str] = []
        last = 0
        for found in self._matches(text):
            out.append(text[last:found.start()])
            out.append(_replaced(found, replacement, text))
            last = found.end()
        out.append(text[last:])
        return "".join(out)


def _replaced(found: re.Match[str], replacement: str, text: str) -> str:
    """What a match becomes under a replacement's $ codes."""
    groups = found.re.groups
    out: list[str] = []
    index = 0
    while index < len(replacement):
        char = replacement[index]
        following = replacement[index + 1:index + 2]
        if char != "$" or not following:
            out.append(char)
            index += 1
            continue
        if following == "$":
            out.append("$")
            index += 2
        elif following == "&":
            out.append(found.group(0))
            index += 2
        elif following == "`":
            out.append(text[:found.start()])
            index += 2
        elif following == "'":
            out.append(text[found.end():])
            index += 2
        elif following.isdigit():
            two = replacement[index + 1:index + 3]
            if len(two) == 2 and two.isdigit() and 1 <= int(two) <= groups:
                number, index = int(two), index + 3
            elif 1 <= int(following) <= groups:
                number, index = int(following), index + 2
            else:
                out.append("$")
                index += 1
                continue
            out.append(found.group(number) or "")
        else:
            out.append("$")
            index += 1
    return "".join(out)


class Match(VBAObject):
    """One match Execute found."""

    vba_type_name = "Match"

    def __init__(self, found: re.Match[str]) -> None:
        self._found = found

    @member(default=True)
    def Value(self) -> object:
        return self._found.group(0)

    @member
    def FirstIndex(self) -> object:
        return VBAInt(self._found.start(), "Long")

    @member
    def Length(self) -> object:
        return VBAInt(self._found.end() - self._found.start(), "Long")

    @member
    def SubMatches(self) -> object:
        return SubMatches([EMPTY if group is None else group for group in self._found.groups()])


class _Items(VBAObject):
    """A collection counted from 0, an item past its end being error 5."""

    def __init__(self, items: list[object]) -> None:
        self._items = items

    @member(default=True)
    def Item(self, index: object = MISSING) -> object:
        if index is MISSING:
            raise error(449)
        position = int(to_integer(index, "Long"))
        if not 0 <= position < len(self._items):
            raise error(5, "Invalid procedure call or argument")
        return self._items[position]

    @member
    def Count(self) -> object:
        return VBAInt(len(self._items), "Long")

    def vba_iterate(self) -> Iterator[object]:
        yield from list(self._items)


class MatchCollection(_Items):
    """The matches Execute found."""

    vba_type_name = "MatchCollection"


class SubMatches(_Items):
    """What a match's groups took."""

    vba_type_name = "SubMatches"
