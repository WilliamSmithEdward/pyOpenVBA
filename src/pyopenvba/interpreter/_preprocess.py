"""Conditional compilation: what the VBA compiler sees before it parses.

``#If`` is resolved before parsing, as VBA resolves it, so a module that
declares a 64-bit API one way and a 32-bit API the other parses as one
of the two rather than as both at once.  Excluded lines are blanked and
not removed, so every line number a later error reports still matches
the source the reader has in front of them.

The constants are the ones the host defines (``VBA7``, ``Win64``,
``Win32``, ``Mac``) plus whatever ``#Const`` declares.  An undeclared
name is Empty, which is false, as it is in VBA.
"""

from __future__ import annotations

import re
from typing import Any, Final

from pyopenvba.exceptions import VBACompileError

#: What 64-bit desktop Office defines.  Win32 is true there too: it
#: means "the Windows API", not "a 32-bit process".
DEFAULT_CONSTANTS: Final[dict[str, object]] = {
    "VBA6": True,
    "VBA7": True,
    "Win16": False,
    "Win32": True,
    "Win64": True,
    "Mac": False,
    "MacOffice16": False,
}

_DIRECTIVE = re.compile(r"^[ \t]*#[ \t]*(?P<word>[A-Za-z]+)(?P<rest>.*)$")
_CONST = re.compile(r"^[ \t]*(?P<name>[A-Za-z_][A-Za-z0-9_]*)[ \t]*=(?P<value>.*)$")
_TOKEN = re.compile(
    r"""(?P<ws>\s+)
      | (?P<string>"(?:[^"]|"")*")
      | (?P<number>\d+(?:\.\d*)?)
      | (?P<name>[A-Za-z_][A-Za-z0-9_]*)
      | (?P<op><>|<=|>=|=|<|>|\(|\))""",
    re.VERBOSE,
)


def strip_directives(source: str, *, module: str = "", constants: dict[str, object] | None = None) -> str:
    """``source`` with every branch the constants exclude blanked out."""
    if "#" not in source:
        return source
    defined: dict[str, object] = dict(DEFAULT_CONSTANTS)
    if constants:
        defined.update(constants)
    lines = source.splitlines(keepends=True)
    out: list[str] = []
    # Each open #If is (taking_now, taken_already).
    stack: list[tuple[bool, bool]] = []
    for number, line in enumerate(lines, start=1):
        match = _DIRECTIVE.match(line)
        word = match.group("word").lower() if match else ""
        rest = match.group("rest") if match else ""
        live = all(taking for taking, _ in stack)
        if word == "const" and live:
            named = _CONST.match(rest)
            if not named:
                raise VBACompileError("#Const has to name a constant", where=_where(module, number))
            defined[named.group("name")] = _evaluate(named.group("value"), defined, module, number)
            out.append(_blank(line))
            continue
        if word == "if":
            condition = _truth(_evaluate(_without_then(rest, module, number), defined, module, number))
            stack.append((condition and live, condition and live))
            out.append(_blank(line))
            continue
        if word == "elseif":
            if not stack:
                raise VBACompileError("#ElseIf without #If", where=_where(module, number))
            _, taken = stack[-1]
            outer = all(taking for taking, _ in stack[:-1])
            condition = (not taken) and outer and _truth(
                _evaluate(_without_then(rest, module, number), defined, module, number)
            )
            stack[-1] = (condition, taken or condition)
            out.append(_blank(line))
            continue
        if word == "else":
            if not stack:
                raise VBACompileError("#Else without #If", where=_where(module, number))
            _, taken = stack[-1]
            outer = all(taking for taking, _ in stack[:-1])
            stack[-1] = ((not taken) and outer, True)
            out.append(_blank(line))
            continue
        if word == "end" and rest.strip().lower() == "if":
            if not stack:
                raise VBACompileError("#End If without #If", where=_where(module, number))
            stack.pop()
            out.append(_blank(line))
            continue
        out.append(line if live else _blank(line))
    if stack:
        raise VBACompileError("a #If is never closed by #End If", where=_where(module, len(lines)))
    return "".join(out)


def _where(module: str, line: int) -> str:
    return f"{module} line {line}" if module else f"line {line}"


def _blank(line: str) -> str:
    """The line's ending and nothing else, so line numbers do not shift."""
    stripped = line.rstrip("\r\n")
    return line[len(stripped) :]


def _without_then(text: str, module: str, line: int) -> str:
    body = text.strip()
    if body.lower().endswith("then"):
        return body[:-4]
    raise VBACompileError("#If has to be followed by Then", where=_where(module, line))


def _truth(value: object) -> bool:
    if value is None or value is False:
        return False
    if isinstance(value, str):
        return value != ""
    return bool(value)


def _evaluate(text: str, constants: dict[str, object], module: str, line: int) -> object:
    """A directive expression, which is constants, literals and operators."""
    tokens: list[tuple[str, str]] = []
    position = 0
    while position < len(text):
        match = _TOKEN.match(text, position)
        if not match:
            raise VBACompileError(
                f"cannot read the directive at {text[position:position + 12]!r}", where=_where(module, line)
            )
        position = match.end()
        kind = match.lastgroup or ""
        if kind != "ws":
            tokens.append((kind, match.group(0)))
    reader = _Reader(tokens, constants, module, line)
    value = reader.expression()
    if not reader.done:
        raise VBACompileError("the directive has more after its expression", where=_where(module, line))
    return value


class _Reader:
    """Just enough of an expression reader for a compilation directive."""

    def __init__(self, tokens: list[tuple[str, str]], constants: dict[str, object], module: str, line: int) -> None:
        self.tokens = tokens
        self.at = 0
        self.constants = constants
        self.module = module
        self.line = line

    @property
    def done(self) -> bool:
        return self.at >= len(self.tokens)

    def _peek(self) -> tuple[str, str]:
        return self.tokens[self.at] if self.at < len(self.tokens) else ("", "")

    def expression(self) -> object:
        return self._logical()

    def _logical(self) -> object:
        left = self._comparison()
        while True:
            kind, text = self._peek()
            word = text.lower()
            if kind != "name" or word not in ("and", "or", "xor"):
                return left
            self.at += 1
            right = self._comparison()
            if word == "and":
                left = _truth(left) and _truth(right)
            elif word == "or":
                left = _truth(left) or _truth(right)
            else:
                left = _truth(left) != _truth(right)

    def _comparison(self) -> object:
        left = self._unary()
        kind, text = self._peek()
        if kind != "op" or text not in ("=", "<>", "<", ">", "<=", ">="):
            return left
        self.at += 1
        right = self._unary()
        first, second = _pair(left, right)
        if text == "=":
            return first == second
        if text == "<>":
            return first != second
        order = (first > second) - (first < second)
        if text == "<":
            return order < 0
        if text == ">":
            return order > 0
        if text == "<=":
            return order <= 0
        return order >= 0

    def _unary(self) -> object:
        kind, text = self._peek()
        if kind == "name" and text.lower() == "not":
            self.at += 1
            return not _truth(self._unary())
        return self._primary()

    def _primary(self) -> object:
        if self.done:
            raise VBACompileError("the directive ends where a value was expected", where=_where(self.module, self.line))
        kind, text = self.tokens[self.at]
        self.at += 1
        if kind == "number":
            return float(text) if "." in text else int(text)
        if kind == "string":
            return text[1:-1].replace('""', '"')
        if kind == "op" and text == "(":
            value = self.expression()
            if self._peek()[1] != ")":
                raise VBACompileError("a directive's bracket is not closed", where=_where(self.module, self.line))
            self.at += 1
            return value
        if kind == "name":
            lower = text.lower()
            if lower == "true":
                return True
            if lower == "false":
                return False
            for name, value in self.constants.items():
                if name.lower() == lower:
                    return value
            return None
        raise VBACompileError(f"unexpected {text!r} in a directive", where=_where(self.module, self.line))


def _pair(left: object, right: object) -> tuple[Any, Any]:
    """The two sides of a directive comparison, made comparable."""
    if isinstance(left, str) or isinstance(right, str):
        return str(left), str(right)
    return (0 if left is None else left), (0 if right is None else right)
