"""The ``Formulas/Section1.m`` document: one shared member per query.

A section document is M source, and the queries in a workbook are its
members::

    section Section1;

    [ Description = "What it does" ]
    shared #"Order Lines" = let
        Source = ...
    in
        Source;

Editing goes through the source text rather than around it: a member
keeps the span it occupies, and a change splices that span, so every
other byte of the document survives untouched.  That is what makes a
read-and-write round trip byte for byte.

Excel's own layout, measured by having it add queries to workbooks and
reading the bytes back: CRLF throughout, one blank line between members,
and no newline after the last one.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from pyopenvba.exceptions import PowerQueryError

#: What Excel writes between members, and at the end of the header.
NEWLINE = "\r\n"
BLANK = NEWLINE + NEWLINE
#: The section name Excel gives every workbook.
DEFAULT_SECTION = "Section1"

_IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
_KEYWORDS = frozenset(
    """and as each else error false if in is let meta not otherwise or section
    shared then true try type #binary #date #datetime #datetimezone #duration
    #infinity #nan #sections #shared #table #time""".split()
)


def quote_name(name: str) -> str:
    """A member name as the document spells it: bare when it is a plain
    identifier, ``#"..."`` otherwise."""
    if _IDENTIFIER.fullmatch(name) and name not in _KEYWORDS:
        return name
    return '#"' + name.replace('"', '""') + '"'


def unquote_name(text: str) -> str:
    """The name a spelling stands for."""
    if text.startswith('#"') and text.endswith('"') and len(text) >= 3:
        return text[2:-1].replace('""', '"')
    return text


class _Scanner:
    """A minimal M lexer: enough to know where a token ends.

    It has to be exact about text literals, quoted identifiers and both
    comment forms, because a ``;`` or a ``,`` inside any of them is not a
    separator -- a query whose step is ``Txt = "a, b, c"`` has one step,
    not three.
    """

    def __init__(self, text: str) -> None:
        self.text = text
        self.at = 0

    def skip_trivia(self) -> None:
        """Move past whitespace and comments."""
        while self.at < len(self.text):
            char = self.text[self.at]
            if char.isspace():
                self.at += 1
            elif self.text.startswith("//", self.at):
                end = self.text.find("\n", self.at)
                self.at = len(self.text) if end < 0 else end + 1
            elif self.text.startswith("/*", self.at):
                end = self.text.find("*/", self.at + 2)
                if end < 0:
                    raise PowerQueryError("a block comment in the section is never closed")
                self.at = end + 2
            else:
                return

    def _skip_quoted(self) -> None:
        """Move past a text literal or a quoted identifier, doubled quotes
        included."""
        self.at += 2 if self.text.startswith('#"', self.at) else 1
        while True:
            end = self.text.find('"', self.at)
            if end < 0:
                raise PowerQueryError("a text literal in the section is never closed")
            self.at = end + 1
            if not self.text.startswith('"', self.at):
                return
            self.at += 1

    def step(self) -> None:
        """Move past one token, or one balanced bracket group."""
        self.skip_trivia()
        if self.at >= len(self.text):
            return
        char = self.text[self.at]
        if char == '"' or self.text.startswith('#"', self.at):
            self._skip_quoted()
            return
        if char in "([{":
            self.skip_group()
            return
        if char.isalpha() or char == "_" or (char == "#" and _identifier_continues(self.text, self.at + 1)):
            self.at += 1
            while _identifier_continues(self.text, self.at):
                self.at += 1
            return
        if char.isdigit():
            self.at += 1
            while self.at < len(self.text) and (self.text[self.at].isalnum() or self.text[self.at] == "."):
                self.at += 1
            return
        self.at += 1

    def skip_group(self) -> None:
        """Move past a balanced ``(...)``, ``[...]`` or ``{...}``."""
        closers = {"(": ")", "[": "]", "{": "}"}
        opener = self.text[self.at]
        if opener not in closers:
            raise PowerQueryError(f"expected a bracket at {self.at}, found {opener!r}")
        depth = 0
        while self.at < len(self.text):
            char = self.text[self.at]
            if char == '"' or self.text.startswith('#"', self.at):
                self._skip_quoted()
                continue
            if self.text.startswith("//", self.at) or self.text.startswith("/*", self.at):
                self.skip_trivia()
                continue
            if char in closers:
                depth += 1
            elif char in ")]}":
                depth -= 1
                if depth == 0:
                    self.at += 1
                    return
            self.at += 1
        raise PowerQueryError("a bracket group in the section is never closed")

    def find_top_level(self, wanted: str, stop: int | None = None) -> int:
        """Where the next `wanted` character sits at bracket depth zero, or
        -1.  Brackets, comments and quoted text are stepped over whole."""
        limit = len(self.text) if stop is None else stop
        while self.at < limit:
            self.skip_trivia()
            if self.at >= limit:
                break
            char = self.text[self.at]
            if char == wanted:
                return self.at
            before = self.at
            self.step()
            if self.at <= before:  # pragma: no cover - the scanner always advances
                raise PowerQueryError("the section scanner stopped moving")
        return -1


@dataclass
class Member:
    """One ``shared Name = expression;`` and the span it occupies."""

    name: str
    #: The whole member, attributes included, without its separator.
    start: int
    end: int
    #: The expression, between the ``=`` and the ``;``.
    formula_start: int
    formula_end: int
    #: The name as written, so a rename knows what to replace.
    name_start: int
    name_end: int
    #: The attribute record ahead of the member, when it has one.
    attributes: str | None
    shared: bool


class Section:
    """A parsed section document that can be edited and written back."""

    def __init__(self, text: str) -> None:
        self._text = text
        self._name = DEFAULT_SECTION
        self._members: list[Member] = []
        self._parse()

    # -- reading ------------------------------------------------------------

    @property
    def text(self) -> str:
        return self._text

    @property
    def name(self) -> str:
        return self._name

    @property
    def members(self) -> list[Member]:
        return list(self._members)

    def names(self) -> list[str]:
        return [member.name for member in self._members]

    def member(self, name: str) -> Member:
        for member in self._members:
            if member.name == name:
                return member
        raise PowerQueryError(f"this section has no member named {name!r}")

    def has(self, name: str) -> bool:
        return any(member.name == name for member in self._members)

    def formula(self, name: str) -> str:
        member = self.member(name)
        return self._text[member.formula_start : member.formula_end]

    def description(self, name: str) -> str | None:
        """The ``Description`` an attribute record carries, if any."""
        attributes = self.member(name).attributes
        if not attributes:
            return None
        return _attribute_text(attributes, "Description")

    def steps(self, name: str) -> list[str]:
        """The top-level ``let`` bindings of this member's expression.

        Excel records one metadata item per binding, and none at all when
        the expression is not a ``let`` -- a function whose body is a
        ``let`` included (measured).
        """
        return let_steps(self.formula(name))

    # -- editing ------------------------------------------------------------

    def set_formula(self, name: str, formula: str) -> None:
        member = self.member(name)
        self._splice(member.formula_start, member.formula_end, formula)

    def set_description(self, name: str, description: str | None) -> None:
        """Add, change or drop the attribute record ahead of a member."""
        member = self.member(name)
        record = None if description is None else f'[ Description = "{_escape_m(description)}" ]'
        if member.attributes is None:
            if record is None:
                return
            self._splice(member.start, member.start, record + NEWLINE)
            return
        attribute_end = member.start + len(member.attributes)
        if record is None:
            trailing = attribute_end
            while trailing < len(self._text) and self._text[trailing] in "\r\n":
                trailing += 1
            self._splice(member.start, trailing, "")
            return
        self._splice(member.start, attribute_end, record)

    def rename(self, old: str, new: str, *, update_references: bool = False) -> None:
        """Rename a member, and optionally the references to it.

        A reference is an identifier token that names the member, so a
        match inside a text literal, a comment or a record's field name is
        left where it is.
        """
        if old == new:
            return
        if self.has(new):
            raise PowerQueryError(f"this section already has a member named {new!r}")
        member = self.member(old)
        self._splice(member.name_start, member.name_end, quote_name(new))
        if update_references:
            self._replace_all(rename_references(self._text, old, new, skip=self.member(new)))

    def add(self, name: str, formula: str, description: str | None = None, *, shared: bool = True) -> None:
        """Append a member, laid out the way Excel lays one out."""
        if self.has(name):
            raise PowerQueryError(f"this section already has a member named {name!r}")
        record = "" if description is None else f'[ Description = "{_escape_m(description)}" ]' + NEWLINE
        head = "shared " if shared else ""
        body = f"{record}{head}{quote_name(name)} = {formula};"
        text = self._text.rstrip("\r\n")
        self._replace_all(text + BLANK + body)

    def remove(self, name: str) -> None:
        member = self.member(name)
        start, end = member.start, member.end
        # Take the blank line that follows, or the one ahead of the last member.
        while end < len(self._text) and self._text[end] in "\r\n":
            end += 1
        if end >= len(self._text):
            while start > 0 and self._text[start - 1] in "\r\n":
                start -= 1
        self._splice(start, end, "")

    # -- writing ------------------------------------------------------------

    def _splice(self, start: int, end: int, replacement: str) -> None:
        self._replace_all(self._text[:start] + replacement + self._text[end:])

    def _replace_all(self, text: str) -> None:
        self._text = text
        self._parse()

    # -- parsing ------------------------------------------------------------

    def _parse(self) -> None:
        text = self._text
        scanner = _Scanner(text)
        scanner.skip_trivia()
        # An attribute record may sit ahead of the section keyword.
        if scanner.at < len(text) and text[scanner.at] == "[":
            scanner.skip_group()
            scanner.skip_trivia()
        if not text.startswith("section", scanner.at):
            raise PowerQueryError("this is not a section document: it does not start with 'section'")
        scanner.at += len("section")
        scanner.skip_trivia()
        name_start = scanner.at
        scanner.step()
        self._name = unquote_name(text[name_start : scanner.at].strip())
        end = _Scanner(text)
        end.at = scanner.at
        semicolon = end.find_top_level(";")
        if semicolon < 0:
            raise PowerQueryError("the section header has no ';'")
        at = semicolon + 1
        self._members = []
        while True:
            probe = _Scanner(text)
            probe.at = at
            probe.skip_trivia()
            if probe.at >= len(text):
                break
            self._members.append(self._parse_member(text, probe.at))
            at = self._members[-1].end

    def _parse_member(self, text: str, start: int) -> Member:
        scanner = _Scanner(text)
        scanner.at = start
        attributes: str | None = None
        if text[scanner.at] == "[":
            group_start = scanner.at
            scanner.skip_group()
            attributes = text[group_start : scanner.at]
            scanner.skip_trivia()
        shared = False
        if text.startswith("shared", scanner.at) and not _identifier_continues(text, scanner.at + 6):
            shared = True
            scanner.at += len("shared")
            scanner.skip_trivia()
        name_start = scanner.at
        scanner.step()
        name_end = scanner.at
        name = unquote_name(text[name_start:name_end])
        if not name:
            raise PowerQueryError(f"a section member at {start} has no name")
        equals = _Scanner(text)
        equals.at = name_end
        equals.skip_trivia()
        if equals.at >= len(text) or text[equals.at] != "=":
            raise PowerQueryError(f"the section member {name!r} has no '='")
        formula_start = equals.at + 1
        finder = _Scanner(text)
        finder.at = formula_start
        semicolon = finder.find_top_level(";")
        if semicolon < 0:
            raise PowerQueryError(f"the section member {name!r} has no ';'")
        formula = text[formula_start:semicolon]
        lead = len(formula) - len(formula.lstrip())
        return Member(
            name=name,
            start=start,
            end=semicolon + 1,
            formula_start=formula_start + lead,
            formula_end=semicolon - (len(formula) - len(formula.rstrip())),
            name_start=name_start,
            name_end=name_end,
            attributes=attributes,
            shared=shared,
        )


def _identifier_continues(text: str, at: int) -> bool:
    return at < len(text) and (text[at].isalnum() or text[at] == "_")


def _escape_m(value: str) -> str:
    return value.replace('"', '""')


def _attribute_text(record: str, field: str) -> str | None:
    """The text value a record's field holds, or None."""
    inner = record.strip()
    if not (inner.startswith("[") and inner.endswith("]")):
        return None
    scanner = _Scanner(inner[1:-1])
    body = scanner.text
    at = 0
    while at < len(body):
        finder = _Scanner(body)
        finder.at = at
        finder.skip_trivia()
        key_start = finder.at
        finder.step()
        key = unquote_name(body[key_start : finder.at])
        finder.skip_trivia()
        if finder.at < len(body) and body[finder.at] == "=":
            value_start = finder.at + 1
            comma = _Scanner(body)
            comma.at = value_start
            end = comma.find_top_level(",")
            end = len(body) if end < 0 else end
            if key == field:
                value = body[value_start:end].strip()
                if value.startswith('"') and value.endswith('"'):
                    return value[1:-1].replace('""', '"')
                return value
            at = end + 1
            continue
        break
    return None


def _keyword_at(text: str, at: int, word: str) -> bool:
    return text.startswith(word, at) and not _identifier_continues(text, at + len(word))


def _binding_end(text: str, at: int) -> tuple[int, str]:
    """Where one ``let`` binding ends: the comma that starts the next one,
    or the ``in`` that closes the ``let``.

    A binding's own expression may be another ``let``, whose commas and
    whose ``in`` belong to it -- so the scan counts ``let`` against ``in``
    as well as stepping over brackets, comments and quoted text.
    """
    depth = 0
    scanner = _Scanner(text)
    scanner.at = at
    while scanner.at < len(text):
        scanner.skip_trivia()
        if scanner.at >= len(text):
            break
        if text[scanner.at] == "," and depth == 0:
            return scanner.at, ","
        if _keyword_at(text, scanner.at, "let"):
            depth += 1
            scanner.at += len("let")
            continue
        if _keyword_at(text, scanner.at, "in"):
            if depth == 0:
                return scanner.at, "in"
            depth -= 1
            scanner.at += len("in")
            continue
        before = scanner.at
        scanner.step()
        if scanner.at <= before:  # pragma: no cover - the scanner always advances
            break
    return len(text), ""


def let_steps(formula: str) -> list[str]:
    """The names of a ``let`` expression's top-level bindings.

    An expression that is not a ``let`` has none, which is what Excel
    records: a function, even one whose body is a ``let``, gets no step
    items at all.  Bindings of a nested ``let`` belong to that ``let``
    and are not steps of the query (measured against Excel).
    """
    scanner = _Scanner(formula)
    scanner.skip_trivia()
    if not _keyword_at(formula, scanner.at, "let"):
        return []
    at = scanner.at + len("let")
    steps: list[str] = []
    while True:
        cursor = _Scanner(formula)
        cursor.at = at
        cursor.skip_trivia()
        if cursor.at >= len(formula):
            break
        name_start = cursor.at
        cursor.step()
        name = unquote_name(formula[name_start : cursor.at])
        cursor.skip_trivia()
        if cursor.at >= len(formula) or formula[cursor.at] != "=":
            break
        steps.append(name)
        end, terminator = _binding_end(formula, cursor.at + 1)
        if terminator != ",":
            break
        at = end + 1
    return steps


def is_function_expression(formula: str) -> bool:
    """Whether the expression is written as a function.

    Excel records ``ResultType = Function`` from what a query *evaluates*
    to, not from how it is written: it marks ``each _ + 1``, ``(x) => x``
    and ``let F = (x) => x in F`` alike (measured).  The last of those
    cannot be told from the text without running the query, so this
    answers the syntactic question -- a function literal, with or without
    a return type, or an ``each`` -- and leaves the rest to Excel, which
    rewrites the entry on its next refresh.
    """
    scanner = _Scanner(formula)
    scanner.skip_trivia()
    if _keyword_at(formula, scanner.at, "each"):
        return True
    if scanner.at >= len(formula) or formula[scanner.at] != "(":
        return False
    scanner.skip_group()
    scanner.skip_trivia()
    if _keyword_at(formula, scanner.at, "as"):
        scanner.at += len("as")
        scanner.step()
        scanner.skip_trivia()
    return formula.startswith("=>", scanner.at)


def _tokens(text: str) -> list[tuple[int, int, str]]:
    """Every identifier token in the text, with where it sits.

    Text literals and comments never yield a token, and an identifier in
    a record's field-name position is marked ``field`` rather than
    ``name`` -- ``[Total]`` names a column, not a query.
    """
    found: list[tuple[int, int, str]] = []

    def walk(start: int, stop: int, *, in_record: bool) -> None:
        scanner = _Scanner(text)
        scanner.at = start
        key_position = in_record
        while scanner.at < stop:
            scanner.skip_trivia()
            if scanner.at >= stop:
                return
            char = text[scanner.at]
            if char == '"':
                scanner.step()
                key_position = False
                continue
            if char in "([{":
                opener = char
                group_start = scanner.at
                scanner.skip_group()
                walk(group_start + 1, scanner.at - 1, in_record=opener == "[")
                key_position = False
                continue
            if char == ",":
                scanner.at += 1
                key_position = in_record
                continue
            before = scanner.at
            if char.isalpha() or char == "_" or text.startswith('#"', scanner.at):
                scanner.step()
                found.append((before, scanner.at, "field" if key_position else "name"))
                key_position = False
                continue
            scanner.step()
            if scanner.at <= before:  # pragma: no cover - the scanner always advances
                return
            key_position = False

    walk(0, len(text), in_record=False)
    return found


def rename_references(text: str, old: str, new: str, *, skip: Member | None = None) -> str:
    """Every reference to `old` in the document, renamed.

    `skip` is a member whose own name has already been changed, so the
    declaration is not rewritten twice.
    """
    spelling = quote_name(new)
    out: list[str] = []
    at = 0
    for start, end, kind in _tokens(text):
        if kind != "name" or unquote_name(text[start:end]) != old:
            continue
        if skip is not None and start == skip.name_start and end == skip.name_end:
            continue
        out.append(text[at:start])
        out.append(spelling)
        at = end
    out.append(text[at:])
    return "".join(out)


def new_section(name: str = DEFAULT_SECTION) -> Section:
    """An empty section document, laid out the way Excel lays one out."""
    return Section(f"section {quote_name(name)};")
