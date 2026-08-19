"""A VBA -> VBA7 p-code compiler for statement bodies.

Validated against Microsoft's own compiler: every statement below is
re-emitted byte-for-byte identically to what Access produced for the same
source (16/16 statements in the mixed reference module, 11/11 in the
loop+branch module). See ``README.md``.

Scope: expressions with full VBA precedence, assignment, If/ElseIf/Else,
Do While/Loop, For/Next. Control flow needs no jump fixups -- VBA p-code
is *structured*: IfBlock/ElseBlock/EndIfBlock and For/NextVar are plain
markers that the runtime pairs up itself.

Names resolve against the project identifier table:
``name_operand = 524 + 2*index`` (measured, ref.accdb 2026-08).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "docs"
                       / "research" / "pcode"))
from pcode_asm import encode_instruction

OP = {"Xor": 2, "Or": 3, "And": 4, "Eq": 5, "Ne": 6, "Le": 7, "Ge": 8,
      "Lt": 9, "Gt": 10, "Add": 11, "Sub": 12, "Mod": 13, "IDiv": 14,
      "Mul": 15, "Div": 16, "Concat": 17, "Not": 21, "UMi": 22, "Ld": 32,
      "St": 39, "DoWhile": 98, "ElseBlock": 100, "ElseIfBlock": 101,
      "EndFunc": 105, "EndIfBlock": 107, "ExitDo": 120, "ExitFor": 121,
      "For": 146, "IfBlock": 156, "LitDI2": 172, "LitDI4": 173,
      "LitStr": 185, "Loop": 188, "NextVar": 203, "EndForVariable": 257,
      "StartForVariable": 258}


def _ins(op, operands=(), payload=None):
    return SimpleNamespace(raw_word=op, operands=tuple(operands),
                           payload=payload)


class CompileError(Exception):
    pass


# --- tokenizer ---------------------------------------------------------
_TOK = re.compile(
    r"(?P<ws>\s+)"
    r"|(?P<num>\d+(?:\.\d+)?)"
    r"|(?P<str>\"(?:[^\"]|\"\")*\")"
    r"|(?P<name>[A-Za-z_][A-Za-z0-9_]*)"
    r"|(?P<op><>|<=|>=|[-+*/\\^&=<>(),])"
)


def tokenize(text):
    out, i = [], 0
    while i < len(text):
        m = _TOK.match(text, i)
        if not m:
            raise CompileError("bad token at " + repr(text[i:i + 12]))
        i = m.end()
        if m.lastgroup != "ws":
            out.append((m.lastgroup, m.group()))
    return out


# --- expression parser (precedence climbing) ---------------------------
_CMP = {"=": "Eq", "<>": "Ne", "<": "Lt", ">": "Gt", "<=": "Le", ">=": "Ge"}


class Parser:
    def __init__(self, toks, names):
        self.t, self.i, self.names = toks, 0, names

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        v = self.peek()
        self.i += 1
        return v

    def accept_op(self, *vals):
        k, v = self.peek()
        if k == "op" and v in vals:
            self.i += 1
            return v
        return None

    def accept_kw(self, *kws):
        k, v = self.peek()
        if k == "name" and v.lower() in [w.lower() for w in kws]:
            self.i += 1
            return v
        return None

    def expr(self):
        return self.p_or()

    def p_or(self):
        out = self.p_and()
        while True:
            w = self.accept_kw("Or", "Xor")
            if not w:
                return out
            key = "Or" if w.lower() == "or" else "Xor"
            out = out + self.p_and() + [_ins(OP[key])]

    def p_and(self):
        out = self.p_not()
        while self.accept_kw("And"):
            out = out + self.p_not() + [_ins(OP["And"])]
        return out

    def p_not(self):
        if self.accept_kw("Not"):
            return [*self.p_not(), _ins(OP["Not"])]
        return self.p_cmp()

    def p_cmp(self):
        out = self.p_concat()
        while True:
            o = self.accept_op(*_CMP)
            if not o:
                return out
            out = out + self.p_concat() + [_ins(OP[_CMP[o]])]

    def p_concat(self):
        out = self.p_add()
        while self.accept_op("&"):
            out = out + self.p_add() + [_ins(OP["Concat"])]
        return out

    def p_add(self):
        out = self.p_mul()
        while True:
            o = self.accept_op("+", "-")
            if not o:
                return out
            out = out + self.p_mul() + [_ins(OP["Add" if o == "+" else "Sub"])]

    def p_mul(self):
        out = self.p_idiv()
        while True:
            o = self.accept_op("*", "/")
            if not o:
                return out
            out = out + self.p_idiv() + [_ins(OP["Mul" if o == "*" else "Div"])]

    def p_idiv(self):
        out = self.p_mod()
        while self.accept_op("\\"):
            out = out + self.p_mod() + [_ins(OP["IDiv"])]
        return out

    def p_mod(self):
        out = self.p_unary()
        while self.accept_kw("Mod"):
            out = out + self.p_unary() + [_ins(OP["Mod"])]
        return out

    def p_unary(self):
        if self.accept_op("-"):
            return [*self.p_unary(), _ins(OP["UMi"])]
        return self.p_atom()

    def p_atom(self):
        k, v = self.take()
        if k == "op" and v == "(":
            e = self.expr()
            if not self.accept_op(")"):
                raise CompileError("missing )")
            return e
        if k == "num":
            n = int(v)
            if 0 <= n <= 0x7FFF:
                return [_ins(OP["LitDI2"], (("0x", n),))]
            return [_ins(OP["LitDI4"],
                         (("0x", n & 0xFFFF), ("0x", (n >> 16) & 0xFFFF)))]
        if k == "str":
            s = v[1:-1].replace('""', '"').encode("latin-1")
            return [_ins(OP["LitStr"], (), s)]
        if k == "name":
            return [_ins(OP["Ld"], (("name", _res(self.names, v)),))]
        raise CompileError("unexpected token " + repr(v))


def _res(names, n):
    if n.lower() not in names:
        raise CompileError("unknown identifier " + repr(n)
                           + " (not in project identifier table)")
    return names[n.lower()]


# --- statement compiler ------------------------------------------------
def compile_line(text, names):
    """Compile one source line to p-code bytes, or None if it emits none."""
    s = text.strip()
    if not s or s.startswith("'"):
        return None
    low = s.lower()

    def P(t):
        return Parser(tokenize(t), names)

    def enc(seq):
        return b"".join(encode_instruction(x) for x in seq)

    if low == "else":
        return enc([_ins(OP["ElseBlock"])])
    if low == "end if":
        return enc([_ins(OP["EndIfBlock"])])
    if low == "loop":
        return enc([_ins(OP["Loop"])])
    if low == "exit do":
        return enc([_ins(OP["ExitDo"])])
    if low == "exit for":
        return enc([_ins(OP["ExitFor"])])
    if low.startswith("elseif ") and low.endswith(" then"):
        return enc([*P(s[7:-5]).expr(), _ins(OP["ElseIfBlock"])])
    if low.startswith("if ") and low.endswith(" then"):
        return enc([*P(s[3:-5]).expr(), _ins(OP["IfBlock"])])
    if low.startswith("do while "):
        return enc([*P(s[9:]).expr(), _ins(OP["DoWhile"])])

    m = re.match(r"(?i)^for\s+([A-Za-z_]\w*)\s*=\s*(.+?)\s+to\s+(.+)$", s)
    if m:
        v, a, b = m.groups()
        ref = [_ins(OP["StartForVariable"]),
               _ins(OP["Ld"], (("name", _res(names, v)),)),
               _ins(OP["EndForVariable"])]
        return enc([*ref, *P(a).expr(), *P(b).expr(), _ins(OP["For"])])

    m = re.match(r"(?i)^next\s+([A-Za-z_]\w*)$", s)
    if m:
        return enc([_ins(OP["StartForVariable"]),
                    _ins(OP["Ld"], (("name", _res(names, m.group(1))),)),
                    _ins(OP["EndForVariable"]), _ins(OP["NextVar"])])

    m = re.match(r"^([A-Za-z_]\w*)\s*=\s*(.+)$", s)
    if m:
        target, rhs = m.groups()
        return enc([*P(rhs).expr(),
                    _ins(OP["St"], (("name", _res(names, target)),))])

    raise CompileError("unsupported statement: " + repr(s))


# Words the statement grammar consumes itself; everything else that
# lexes as a name is an identifier the program references.
_KEYWORDS = frozenset(["if", "then", "else", "elseif", "end", "do", "while", "loop", "for", "to", "next", "step", "exit", "not", "and", "or", "xor", "mod"])


# A comment line is stored as text in the same region as the p-code,
# tagged E3 and pointed at by a line record of kind 0x09:
#
#     E3 00 <u16 indent> <u16 text length> <text>
#
# The leading apostrophe is dropped and everything after it kept
# verbatim. The indent is the source line's leading-space count -- the
# same value a code line carries in its line record's byte 3.
_COMMENT_TAG = b"\xe3\x00"


def comment_record(text: str, code_page: int = 1252) -> bytes:
    """Encode one comment line for the module's text region."""
    indent = len(text) - len(text.lstrip())
    body = text.strip()
    if body.startswith("'"):
        body = body[1:]
    encoded = body.encode(f"cp{code_page}", errors="replace")
    record = (_COMMENT_TAG + indent.to_bytes(2, "little")
              + len(encoded).to_bytes(2, "little") + encoded)
    return record + b"\x00" * (len(record) & 1)      # padded to even length


def is_comment(text: str) -> bool:
    return text.strip().startswith("'")


def referenced_names(text):
    """Identifier-like tokens in one statement, minus VBA keywords.

    A comment introduces no identifiers, and its free text does not lex,
    so skip it rather than trying.
    """
    if is_comment(text):
        return []
    out = []
    for kind, value in tokenize(text.strip()):
        if kind == "name" and value.lower() not in _KEYWORDS:
            out.append(value)
    return out


def name_table(path):
    """Map lowercased identifier name -> p-code name operand.

    Positional records are addressed as ``524 + 2*index`` and records
    carrying their own slot as ``2*slot + 2``. Use each record's own
    ``index`` rather than its position in the tuple: slotted records take
    no position, so enumerating would shift every name after one.
    """
    from pyopenvba.access_read import AccessReader

    table: dict[str, int] = {}
    for record in AccessReader(path).identifiers():
        operand = (2 * record.slot + 2 if record.slot is not None
                   else 524 + 2 * record.index)
        table[record.name.lower()] = operand
    return table
