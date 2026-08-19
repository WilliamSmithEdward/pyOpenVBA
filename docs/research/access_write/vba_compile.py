"""A VBA -> VBA7 p-code compiler for statement bodies.

Validated against Microsoft's own compiler: every statement below is
re-emitted byte-for-byte identically to what Access produced for the same
source. ``construct_matrix.bas`` is the corpus that decides what "below"
means -- build it with ``build_matrix.ps1`` and check it with
``verify_compiler.py``.

Covered: expressions with full VBA precedence including ``^``; integer,
floating-point and string literals; ``True``/``False``/``Null``/
``Empty``/``Nothing``; assignment to a variable, a member, an array
element or an indexed member; ``Set``; calls in statement and expression
position, on a name or a member; ``If``/``ElseIf``/``Else``;
``Do While``/``Do Until``/``Loop``/``Loop While``/``Loop Until``;
``While``/``Wend``; ``For``/``For ... Step``/``Next``; ``Exit`` in its
four forms; and ``Select Case`` with plain, list, ``To`` and ``Is``
clauses.

Not covered, and refused rather than approximated: ``With``,
``Dim``/``Const``/``ReDim``/``Erase``, ``On Error``, line labels,
single-line ``If ... Then <statement>``, date literals, and calls to
built-ins that live in the pre-populated slots below 261 rather than in
the project identifier table.

Refusal matters more than coverage here. A compiler that drops what it
does not understand emits valid p-code for the wrong program; three such
faults (a discarded ``^`` operand, an array assignment compiled as a
call, and a member call with its arguments and object transposed) were
found only by diffing against Access, so the grammar now fails on any
input it cannot fully consume.

Names resolve against the project identifier table:
``name_operand = 524 + 2*index`` (measured, ref.accdb 2026-08).
"""
from __future__ import annotations

import re
import struct
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "docs"
                       / "research" / "pcode"))
from pcode_asm import encode_instruction

from pyopenvba.vba_pcode import OPCODES_VBA7

OP = {"Xor": 2, "Or": 3, "And": 4, "Eq": 5, "Ne": 6, "Le": 7, "Ge": 8,
      "Lt": 9, "Gt": 10, "Add": 11, "Sub": 12, "Mod": 13, "IDiv": 14,
      "Mul": 15, "Div": 16, "Concat": 17, "Pwr": 19, "Not": 21, "UMi": 22,
      "Ld": 32, "MemLd": 33, "ArgsLd": 36, "ArgsMemLd": 37, "St": 39,
      "MemSt": 40, "ArgsSt": 43, "ArgsMemSt": 44, "Set": 46,
      "ArgsCall": 65, "ArgsMemCall": 66,
      "Case": 75, "CaseTo": 76, "CaseGt": 77, "CaseLt": 78, "CaseGe": 79,
      "CaseLe": 80, "CaseNe": 81, "CaseEq": 82, "CaseElse": 83,
      "CaseDone": 84,
      "DoUntil": 97, "DoWhile": 98, "ElseBlock": 100, "ElseIfBlock": 101,
      "EndFunc": 105, "EndIfBlock": 107, "EndSelect": 110, "EndSub": 111,
      "ExitDo": 120, "ExitFor": 121, "ExitFunc": 122, "ExitSub": 124,
      "For": 146, "ForStep": 149, "IfBlock": 156, "LitDI2": 172,
      "LitDI4": 173, "LitNothing": 178, "LitR8": 183, "LitStr": 185,
      "LitVarSpecial": 186, "Loop": 188, "LoopUntil": 189, "LoopWhile": 190,
      "NextVar": 203, "SelectCase": 237, "SetStmt": 240, "Wend": 246,
      "While": 247, "EndForVariable": 257, "StartForVariable": 258}

# Every entry must agree with the canonical VBA7 table. Two of these were
# transcribed with their numbers swapped (LitNothing/LitR8), which the
# differential gate caught only because a probe happened to use both.
_WRONG = {n: v for n, v in OP.items() if OPCODES_VBA7[v][0] != n}
assert not _WRONG, f"opcode numbers disagree with OPCODES_VBA7: {_WRONG}"

# `LitVarSpecial` names a constant in its op_type: measured from
# `v = True` (~1), `v = False` (~0), `v = Null` (~2) and `v = Empty` (~3).
VAR_SPECIAL = {"false": 0, "true": 1, "null": 2, "empty": 3}


# An instruction word is the opcode in its low 10 bits and an op_type in
# the rest. Statement calls -- the ones whose result is discarded -- carry
# op_type 16, measured from `MsgBox "hi"` (0x4041) and `DoCmd.Beep`
# (0x4042) against the plain 0x0041 / 0x0042 an expression call uses.
STATEMENT_CALL_OP_TYPE = 16


def _ins(op, operands=(), payload=None, op_type=0):
    return SimpleNamespace(raw_word=op | (op_type << 10),
                           operands=tuple(operands), payload=payload)


class CompileError(Exception):
    pass


# --- tokenizer ---------------------------------------------------------
_TOK = re.compile(
    r"(?P<ws>\s+)"
    r"|(?P<num>\d+(?:\.\d+)?)"
    r"|(?P<str>\"(?:[^\"]|\"\")*\")"
    r"|(?P<name>[A-Za-z_][A-Za-z0-9_]*)"
    r"|(?P<op><>|<=|>=|[-+*/\\^&=<>(),.])"
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
_CASE_IS = {"=": "CaseEq", "<>": "CaseNe", "<": "CaseLt",
            ">": "CaseGt", "<=": "CaseLe", ">=": "CaseGe"}


class Parser:
    def __init__(self, toks, names):
        self.t, self.i, self.names = toks, 0, names

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        v = self.peek()
        self.i += 1
        return v

    def rest(self):
        """The tokens not yet consumed, for a useful refusal message."""
        return " ".join(v for _, v in self.t[self.i:])

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
        return self.p_pow()

    def p_pow(self):
        out = self.p_atom()
        if self.accept_op("^"):
            # Right-associative: the exponent may itself be a power.
            return out + self.p_pow() + [_ins(OP["Pwr"])]
        return out

    def p_atom(self):
        k, v = self.take()
        if k == "op" and v == "(":
            e = self.expr()
            if not self.accept_op(")"):
                raise CompileError("missing )")
            return e
        if k == "num":
            if "." in v:
                # LitR8 carries the raw little-endian double as four words.
                raw = struct.pack("<d", float(v))
                return [_ins(OP["LitR8"],
                             tuple(("0x", w) for w in
                                   struct.unpack("<4H", raw)))]
            n = int(v)
            if 0 <= n <= 0x7FFF:
                return [_ins(OP["LitDI2"], (("0x", n),))]
            return [_ins(OP["LitDI4"],
                         (("0x", n & 0xFFFF), ("0x", (n >> 16) & 0xFFFF)))]
        if k == "str":
            s = v[1:-1].replace('""', '"').encode("latin-1")
            return [_ins(OP["LitStr"], (), s)]
        if k == "name":
            if v.lower() in VAR_SPECIAL:
                return [_ins(OP["LitVarSpecial"],
                             op_type=VAR_SPECIAL[v.lower()])]
            if v.lower() == "nothing":
                return [_ins(OP["LitNothing"])]
            nxt_kind, nxt = self.peek()
            if nxt_kind == "op" and nxt == ".":
                # obj.Member -- push the object, then read the member.
                self.take()
                member = self.take()
                if member[0] != "name":
                    raise CompileError("expected a member name after '.'")
                had_parens = self.peek() == ("op", "(")
                args = self.arg_list()
                obj = _ins(OP["Ld"], (("name", _res(self.names, v)),))
                if not had_parens:
                    return [obj, _ins(OP["MemLd"],
                                      (("name", _res(self.names, member[1])),))]
                return [*args, obj,
                        _ins(OP["ArgsMemLd"],
                             (("name", _res(self.names, member[1])),
                              ("0x", len(self.last_arg_count))))]
            if nxt_kind == "op" and nxt == "(":
                # Foo(a, b) -- push the arguments, then call by name.
                args = self.arg_list()
                return [*args, _ins(OP["ArgsLd"],
                                    (("name", _res(self.names, v)),
                                     ("0x", len(self.last_arg_count))))]
            return [_ins(OP["Ld"], (("name", _res(self.names, v)),))]
        raise CompileError("unexpected token " + repr(v))

    def arg_list(self):
        """Parse ``(a, b)`` if present; record how many arguments it held."""
        self.last_arg_count = []
        if not self.accept_op("("):
            return []
        out = []
        if self.accept_op(")"):
            return out
        while True:
            out += self.expr()
            self.last_arg_count.append(1)
            if self.accept_op(","):
                continue
            if self.accept_op(")"):
                return out
            raise CompileError("expected ',' or ')' in argument list")


def _res(names, n):
    if n.lower() not in names:
        raise CompileError("unknown identifier " + repr(n)
                           + " (not in project identifier table)")
    return names[n.lower()]


# --- statement compiler ------------------------------------------------
# An assignment target is a name, a member of one, or either with a
# subscript or argument list: `x`, `o.P`, `a(1)`, `o.Item("k")`.
_TARGET = re.compile(r"(?i)^([A-Za-z_]\w*)(?:\.([A-Za-z_]\w*))?(\(.*\))?$")


def _is_assignable(text):
    return _TARGET.match(text.strip()) is not None


def _split_args(text):
    """Split an argument list on its top-level commas."""
    out, depth, start, quoted = [], 0, 0, False
    for i, ch in enumerate(text):
        if ch == '"':
            quoted = not quoted
        elif quoted:
            continue
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            out.append(text[start:i])
            start = i + 1
    out.append(text[start:])
    return [piece.strip() for piece in out]


def _store(target, names, expr):
    """Resolve an assignment target to (opcode, operands, preceding code).

    The preceding code runs after the value and before the store opcode,
    because Access pushes the value first, then any subscript or argument
    list, and pushes the object last of all.
    """
    m = _TARGET.match(target.strip())
    if m is None:
        raise CompileError(f"not an assignable target: {target!r}")
    obj, member, call = m.groups()
    args, count = [], 0
    if call and call[1:-1].strip():
        for piece in _split_args(call[1:-1]):
            args += expr(piece)
            count += 1
    if member:
        holder = _ins(OP["Ld"], (("name", _res(names, obj)),))
        if call:
            return ("ArgsMemSt",
                    (("name", _res(names, member)), ("0x", count)),
                    [*args, holder])
        return ("MemSt", (("name", _res(names, member)),), [holder])
    if call:
        return ("ArgsSt", (("name", _res(names, obj)), ("0x", count)), args)
    return ("St", (("name", _res(names, obj)),), [])


def compile_line(text, names):
    """Compile one source line to p-code bytes, or None if it emits none."""
    s = text.strip()
    if not s or s.startswith("'"):
        return None
    low = s.lower()

    def P(t):
        return Parser(tokenize(t), names)

    def E(t):
        """Parse a whole expression, refusing to drop a trailing tail.

        Silently ignoring unconsumed tokens is how `x = a ^ b` once
        compiled to `Ld(a) | St(x)`: valid p-code for the wrong program.
        Anything this grammar does not understand must fail, not shrink.
        """
        parser = P(t)
        out = parser.expr()
        if parser.peek()[0] is not None:
            raise CompileError(
                f"unparsed input {parser.rest()!r} in expression {t!r}")
        return out

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
    if low == "exit sub":
        return enc([_ins(OP["ExitSub"])])
    if low == "exit function":
        return enc([_ins(OP["ExitFunc"])])
    if low == "wend":
        return enc([_ins(OP["Wend"])])
    if low == "end select":
        return enc([_ins(OP["EndSelect"])])
    if low == "case else":
        return enc([_ins(OP["CaseElse"])])
    if low.startswith("select case "):
        return enc([*E(s[12:]), _ins(OP["SelectCase"])])
    if low.startswith("case "):
        # Each clause pushes its operands then its own comparison opcode;
        # one CaseDone closes the whole list.
        out = []
        for piece in _split_args(s[5:]):
            m = re.match(r"(?i)^is\s*(<>|<=|>=|=|<|>)\s*(.+)$", piece)
            if m:
                out += [*E(m.group(2)), _ins(OP[_CASE_IS[m.group(1)]])]
                continue
            m = re.match(r"(?i)^(.+?)\s+to\s+(.+)$", piece)
            if m:
                out += [*E(m.group(1)), *E(m.group(2)), _ins(OP["CaseTo"])]
                continue
            out += [*E(piece), _ins(OP["Case"])]
        return enc([*out, _ins(OP["CaseDone"])])
    if low.startswith("elseif ") and low.endswith(" then"):
        return enc([*E(s[7:-5]), _ins(OP["ElseIfBlock"])])
    if low.startswith("if ") and low.endswith(" then"):
        return enc([*E(s[3:-5]), _ins(OP["IfBlock"])])
    if low.startswith("do while "):
        return enc([*E(s[9:]), _ins(OP["DoWhile"])])
    if low.startswith("do until "):
        return enc([*E(s[9:]), _ins(OP["DoUntil"])])
    if low.startswith("loop while "):
        return enc([*E(s[11:]), _ins(OP["LoopWhile"])])
    if low.startswith("loop until "):
        return enc([*E(s[11:]), _ins(OP["LoopUntil"])])
    if low.startswith("while "):
        return enc([*E(s[6:]), _ins(OP["While"])])

    m = re.match(r"(?i)^for\s+([A-Za-z_]\w*)\s*=\s*(.+?)\s+to\s+(.+)$", s)
    if m:
        v, a, b = m.groups()
        step = None
        m2 = re.match(r"(?i)^(.+?)\s+step\s+(.+)$", b)
        if m2:
            b, step = m2.groups()
        ref = [_ins(OP["StartForVariable"]),
               _ins(OP["Ld"], (("name", _res(names, v)),)),
               _ins(OP["EndForVariable"])]
        if step is None:
            return enc([*ref, *E(a), *E(b), _ins(OP["For"])])
        return enc([*ref, *E(a), *E(b), *E(step), _ins(OP["ForStep"])])

    m = re.match(r"(?i)^next\s+([A-Za-z_]\w*)$", s)
    if m:
        return enc([_ins(OP["StartForVariable"]),
                    _ins(OP["Ld"], (("name", _res(names, m.group(1))),)),
                    _ins(OP["EndForVariable"]), _ins(OP["NextVar"])])

    # Assignment. Access pushes the value first, then any subscript or
    # argument list, then the object, and closes with the store opcode --
    # so the target is compiled after the value in every form.
    m = re.match(r"(?i)^(set\s+)?(.+?)\s*=\s*(.+)$", s)
    if m and _is_assignable(m.group(2)):
        is_set, target, rhs = m.group(1), m.group(2).strip(), m.group(3)
        value = E(rhs)
        store = _store(target, names, E)
        if is_set:
            # `Set` brackets the whole statement rather than replacing the
            # store: SetStmt, the value, then Set(name).
            if store[0] != "St":
                raise CompileError(
                    f"Set target must be a plain variable, not {target!r}")
            return enc([_ins(OP["SetStmt"]), *value,
                        _ins(OP["Set"], store[1])])
        return enc([*value, *store[2], _ins(OP[store[0]], store[1])])

    # A statement call discards its result: `MsgBox "hi"`, `DoCmd.Beep`,
    # `Helper 7`. Parentheses are optional in this form, so the arguments
    # are whatever follows the name.
    m = re.match(r"(?i)^([A-Za-z_]\w*)(?:\.([A-Za-z_]\w*))?\s*(.*)$", s)
    if m:
        obj, member, rest = m.groups()
        rest = rest.strip()
        if rest.startswith("(") and rest.endswith(")"):
            rest = rest[1:-1].strip()
        parser = P(rest) if rest else None
        args, count = [], 0
        if parser is not None:
            while True:
                args += parser.expr()
                count += 1
                if not parser.accept_op(","):
                    break
            if parser.peek()[0] is not None:
                raise CompileError(f"unparsed input in statement: {s!r}")
        if member:
            # Arguments first, object last -- the same order every other
            # member form uses. A zero-argument call such as `DoCmd.Beep`
            # cannot tell the two orders apart, which is why this was
            # wrong for as long as the probes had no arguments.
            return enc([*args, _ins(OP["Ld"], (("name", _res(names, obj)),)),
                        _ins(OP["ArgsMemCall"],
                             (("name", _res(names, member)), ("0x", count)),
                             op_type=STATEMENT_CALL_OP_TYPE)])
        return enc([*args, _ins(OP["ArgsCall"],
                                (("name", _res(names, obj)), ("0x", count)),
                                op_type=STATEMENT_CALL_OP_TYPE)])

    raise CompileError("unsupported statement: " + repr(s))


# Words the statement grammar consumes itself; everything else that
# lexes as a name is an identifier the program references.
_KEYWORDS = frozenset([
    "if", "then", "else", "elseif", "end", "do", "while", "wend", "until",
    "loop", "for", "to", "next", "step", "exit", "not", "and", "or",
    "xor", "mod", "set", "select", "case", "is", "sub", "function",
    # The Variant constants compile to LitVarSpecial/LitNothing, so they
    # are literals here and never reach the identifier table.
    "true", "false", "null", "empty", "nothing",
])


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
