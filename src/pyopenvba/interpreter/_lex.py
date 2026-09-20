"""The VBA tokenizer.

VBA is a line-oriented language: a statement ends at the end of a line
unless a trailing underscore carries it on, and a colon starts another
statement on the same line.  The tokens carry the line they came from so
that a compile error or a run-time error can name it the way the VBA IDE
does.

Three details are worth stating because they are where a naive tokenizer
goes wrong:

* ``#`` opens a date literal and also ends a Double, and it opens a file
  number in ``Print #1``.  A ``#`` is a date only when the rest of the
  line closes it and what is between reads as a date.
* ``&`` is both concatenation and the Long type suffix.  It is a suffix
  only where an expression cannot continue, so ``1&`` is a Long and
  ``1 & 2`` is "12".
* ``Rem`` is a comment, but only where a statement may begin: ``Rem`` is
  a perfectly good variable name in the middle of an expression.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace
from typing import Final

from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter._values import parse_date_text

#: The characters that may end a name or a number to declare its type.
TYPE_SUFFIXES: Final = "%&!#@$"

#: What each suffix declares, for a name as well as for a literal.
SUFFIX_TYPES: Final[dict[str, str]] = {
    "%": "Integer",
    "&": "Long",
    "!": "Single",
    "#": "Double",
    "@": "Currency",
    "$": "String",
}

#: Every character above ASCII, which VBA allows in a name.  Built
#: with chr() so this source file itself stays plain ASCII.
_ABOVE_ASCII = f"{chr(0x80)}-{chr(0xFFFF)}"

_IDENT = re.compile(f"[A-Za-z_{_ABOVE_ASCII}][A-Za-z0-9_{_ABOVE_ASCII}]*")
_HEX = re.compile(r"&[hH][0-9A-Fa-f]+&?")
_OCTAL = re.compile(r"&[oO][0-7]+&?")
_DECIMAL = re.compile(r"(?:\d+\.\d*|\.\d+|\d+)(?:[eEdD][-+]?\d+)?")
_CONTINUES = re.compile(f"[A-Za-z_\"'{_ABOVE_ASCII}0-9.([]|&[hHoO]")
_DATE = re.compile(r"#([^#\r\n]{1,64})#")
_BRACKET = re.compile(r"\[([^\]\r\n]*)\]")

#: Longest first, so that <= is never read as < then =.
_OPERATORS: Final = (
    ":=",
    "<=",
    ">=",
    "<>",
    "=",
    "<",
    ">",
    "+",
    "-",
    "*",
    "/",
    "\\",
    "^",
    "&",
    "(",
    ")",
    ",",
    ".",
    "!",
    "#",
    ";",
    "?",
    "@",
    "{",
    "}",
)


@dataclass(frozen=True)
class Token:
    """One token, with the line it was read from (counting from 1)."""

    kind: str
    text: str
    line: int
    #: For a number, the type its suffix or shape declares.
    declared: str = ""
    #: Whether whitespace came before it.  This is the only thing that
    #: separates ``Debug.Print .Count`` inside a With block, where the
    #: dot opens an argument, from ``Debug.Print.Count``, where it would
    #: carry on the member chain.
    spaced: bool = False

    @property
    def lower(self) -> str:
        return self.text.lower()

    def __repr__(self) -> str:
        return f"{self.kind}({self.text!r}@{self.line})"


def _ends_statement(token: Token | None) -> bool:
    """Whether a statement may begin at the token after this one."""
    if token is None:
        return True
    if token.kind == "eos":
        return True
    return token.kind == "ident" and token.lower in ("then", "else")


def _expression_may_follow(text: str, at: int) -> bool:
    """Whether what comes after position ``at`` could continue an expression.

    Used to tell the Long suffix in ``1&`` from the concatenation in
    ``1 & x``: an operand may follow the operator, so an ampersand with
    an operand behind it is an operator.
    """
    rest = text[at:]
    stripped = rest.lstrip(" \t")
    if not stripped:
        return False
    return bool(re.match(_CONTINUES, stripped))


def tokenize(source: str, *, module: str = "") -> list[Token]:
    """Every token in ``source``, ending with one of kind ``eof``."""
    tokens: list[Token] = []
    position = 0
    line = 1
    length = len(source)
    spaced = False

    def fail(message: str) -> VBACompileError:
        where = f"{module} line {line}" if module else f"line {line}"
        return VBACompileError(message, where=where)

    def add(token: Token) -> None:
        nonlocal spaced
        tokens.append(replace(token, spaced=spaced))
        spaced = False

    while position < length:
        char = source[position]

        if char in " \t":
            position += 1
            spaced = True
            continue

        if char == "_" and not _IDENT.match(source, position + 1):
            # A continuation: the rest of the line has to be empty.
            rest = source[position + 1 :]
            skipped = len(rest) - len(rest.lstrip(" \t"))
            after = position + 1 + skipped
            if after >= length or source[after] in "\r\n":
                while after < length and source[after] in "\r":
                    after += 1
                if after < length and source[after] == "\n":
                    after += 1
                line += 1
                position = after
                continue
            raise fail("a line continuation has to be the last thing on its line")

        if char in "\r\n":
            if char == "\r" and position + 1 < length and source[position + 1] == "\n":
                position += 1
            position += 1
            add(Token("eos", "\n", line))
            line += 1
            continue

        if char == "'":
            while position < length and source[position] not in "\r\n":
                position += 1
            continue

        if char == '"':
            end = position + 1
            pieces: list[str] = []
            while True:
                if end >= length or source[end] in "\r\n":
                    raise fail("a string literal is not closed before the end of the line")
                if source[end] == '"':
                    if end + 1 < length and source[end + 1] == '"':
                        pieces.append('"')
                        end += 2
                        continue
                    end += 1
                    break
                pieces.append(source[end])
                end += 1
            add(Token("string", "".join(pieces), line))
            position = end
            continue

        if char == "#":
            match = _DATE.match(source, position)
            if match and parse_date_text(match.group(1)) is not None:
                add(Token("date", match.group(1), line))
                position = match.end()
                continue

        if char == "[":
            match = _BRACKET.match(source, position)
            if match:
                add(Token("bracket", match.group(1), line))
                position = match.end()
                continue

        if char == "&" and position + 1 < length and source[position + 1] in "hHoO":
            match = _HEX.match(source, position) or _OCTAL.match(source, position)
            if match:
                text = match.group(0).rstrip("&")
                base = 16 if text[1] in "hH" else 8
                digits = text[2:]
                value = int(digits, base)
                # The width comes from how many digits were written, not
                # from the value: &HFFFF is -1 as an Integer and
                # &HFFFFFFFF is -1 as a Long.
                narrow = len(digits) <= (4 if base == 16 else 6)
                if narrow and value > 0x7FFF:
                    value -= 0x10000
                elif not narrow and value > 0x7FFFFFFF:
                    value -= 0x100000000
                kind = "Integer" if narrow and -32768 <= value <= 32767 else "Long"
                add(Token("number", str(value), line, kind))
                position = match.end()
                continue

        if char.isdigit() or (char == "." and position + 1 < length and source[position + 1].isdigit()):
            match = _DECIMAL.match(source, position)
            if not match:
                raise fail(f"cannot read a number at {source[position:position + 12]!r}")
            text = match.group(0).replace("d", "e").replace("D", "e")
            position = match.end()
            declared = ""
            if position < length and source[position] in TYPE_SUFFIXES:
                suffix = source[position]
                if suffix != "&" or not _expression_may_follow(source, position + 1):
                    declared = SUFFIX_TYPES[suffix]
                    position += 1
            if not declared:
                if "." in text or "e" in text.lower():
                    declared = "Double"
                else:
                    whole = int(text)
                    if -32768 <= whole <= 32767:
                        declared = "Integer"
                    elif -2147483648 <= whole <= 2147483647:
                        declared = "Long"
                    else:
                        # Too wide for a Long, so VBA reads it as a
                        # Double: CStr(1000000000000000) is "1E+15".
                        declared = "Double"
            add(Token("number", text, line, declared))
            continue

        match = _IDENT.match(source, position)
        if match:
            name = match.group(0)
            position = match.end()
            if name.lower() == "rem" and _ends_statement(tokens[-1] if tokens else None):
                while position < length and source[position] not in "\r\n":
                    position += 1
                continue
            declared = ""
            if position < length and source[position] in TYPE_SUFFIXES:
                suffix = source[position]
                if suffix != "&" or not _expression_may_follow(source, position + 1):
                    declared = SUFFIX_TYPES[suffix]
                    name += suffix
                    position += 1
            add(Token("ident", name, line, declared))
            continue

        if char == ":" and not source.startswith(":=", position):
            add(Token("eos", ":", line))
            position += 1
            continue

        for operator in _OPERATORS:
            if source.startswith(operator, position):
                add(Token("op", operator, line))
                position += len(operator)
                break
        else:
            raise fail(f"cannot read {char!r}")

    add(Token("eos", "\n", line))
    add(Token("eof", "", line))
    return tokens


def strip_suffix(name: str) -> str:
    """A name without its type-declaration character."""
    return name[:-1] if name and name[-1] in TYPE_SUFFIXES else name
