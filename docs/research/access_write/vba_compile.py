"""A VBA -> VBA7 p-code compiler for statement bodies.

Validated against Microsoft's own compiler: every statement below is
re-emitted byte-for-byte identically to what Access produced for the same
source. ``construct_matrix.bas`` and ``builtins_probe.bas`` are the
corpora that decide what "below" means -- build them with
``build_matrix.py`` and check them with ``verify_compiler.py``.

Covered: expressions with full VBA precedence, including ``^`` and the
``Paren`` marker Access records for explicit grouping; integer,
floating-point, string and date literals; ``True``/``False``/``Null``/
``Empty``/``Nothing``; assignment to a variable, a member, an array
element or an indexed member; ``Set``; calls in statement and expression
position, on a name, a member, or implicitly inside ``With``;
``If``/``ElseIf``/``Else`` in block and single-line form; ``Select Case``
with plain, list, ``To`` and ``Is`` clauses; ``Do``/``Do While``/
``Do Until``/``Loop``/``Loop While``/``Loop Until``; ``While``/``Wend``;
``For``/``For ... Step``/``Next``; ``Exit`` in four forms; ``With``,
``Erase``, ``On Error GoTo``, line labels, ``Resume Next``, and
``Debug.Print``.

VBA's built-ins are reached four different ways and all four are handled;
see ``BUILTIN_OPCODE`` and its neighbours.

Not covered, and refused rather than approximated: ``Dim``, ``Const``,
``ReDim``, user-defined ``Type``, and ``New`` -- every one of which needs
a record in the pre-0xCAFE header that this code cannot yet grow.

Refusal matters more than coverage here. A compiler that drops what it
does not understand emits valid p-code for the wrong program, and four
such faults have turned up this way: a discarded ``^`` operand, an array
assignment compiled as a call, a member call with its arguments and
object transposed, and a nested call overwriting the outer argument
count. None was visible without diffing against Access, so the grammar
now fails on any input it cannot fully consume.

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
      "Dim": 93, "VarDefn": 245,
      "Paren": 29,
      "Debug": 91, "Do": 95, "LitDate": 170, "PrintItemNL": 217,
      "PrintObj": 220,
      "ArgsMemCallWith": 67, "BoSImplicit": 71, "EndIf": 106,
      "EndWith": 113, "Erase": 114, "If": 155, "Label": 163,
      "OnError": 204, "Resume": 232, "With": 248, "StartWithExpr": 260,
      "FnAbs": 23, "FnLen": 27, "ArgsArray": 68, "Coerce": 88,
      "FnInStr": 132, "FnLBound": 138, "FnUBound": 145,
      "FnFix": 24, "FnInt": 25, "FnSgn": 26, "FnStrComp": 141,
      "NextVar": 203, "SelectCase": 237, "SetStmt": 240, "Wend": 246,
      "While": 247, "EndForVariable": 257, "StartForVariable": 258}

# Every entry must agree with the canonical VBA7 table. Two of these were
# transcribed with their numbers swapped (LitNothing/LitR8), which the
# differential gate caught only because a probe happened to use both.
_WRONG = {n: v for n, v in OP.items() if OPCODES_VBA7[v][0] != n}
assert not _WRONG, f"opcode numbers disagree with OPCODES_VBA7: {_WRONG}"

# A VBA built-in is reached in one of three ways, and which one is not
# guessable -- all three were read off a probe module that calls eighty of
# them (`BI.bas`). Most simply become project identifiers, exactly like a
# user-written name, and need nothing here.
#
# 1. A dedicated opcode, with the arguments already on the stack.
BUILTIN_OPCODE = {"len": "FnLen", "abs": "FnAbs", "int": "FnInt",
                  "fix": "FnFix", "sgn": "FnSgn", "instr": "FnInStr",
                  "strcomp": "FnStrComp"}

# 2. A pre-populated operand slot, called through the ordinary ArgsLd.
#    The operand is the usual ``2*slot + 2``.
BUILTIN_SLOT = {"left": 109, "mid": 124, "string": 173, "format": 85,
                "curdir": 37, "freefile": 87}

# 3. `Array(...)` has its own opcode but still names slot 8.
ARRAY_SLOT = 8

# `UBound`/`LBound` carry the dimension as an operand rather than an
# argument count, defaulting to 0 for the first dimension.
BOUND_OPCODE = {"ubound": "FnUBound", "lbound": "FnLBound"}

# Some names are pre-interned in the slot table and resolve to a slot
# rather than to a project identifier, even when used as an ordinary
# variable. They are VBA's own vocabulary -- `Option Base`, the `Dir`
# and `Name` functions, `Line Input`, `GoTo` -- which is why a procedure
# innocently called `Go` binds to slot 92 while `Zebra` gets a project
# identifier.
#
# Harvested by compiling probes that assign to ~180 candidate names and
# reading back which landed below slot 261. That is a sample, not the
# whole table: 261 slots exist and only these are mapped, so an unmapped
# reserved name still resolves to a fresh project identifier, which is
# wrong but detectable -- Access renumbers it on its next compile.
RESERVED_SLOT = {"b": 11, "base": 12, "cdec": 25, "curdir": 37,
                 "date": 44, "dir": 62, "f": 81, "format": 85,
                 "freefile": 87, "go": 92, "left": 109, "line": 115,
                 "mid": 124, "name": 130, "rgb": 159, "strcomp": 172,
                 "string": 173, "text": 177}

# A few built-ins are values rather than calls and appear bare, loaded
# straight from their slot: Access rewrites `Date()` to `Date` and emits
# `Ld(slot 44)`. `Time` and `Now` are ordinary project identifiers.
BUILTIN_VALUE_SLOT = {"date": 44}

# The conversion functions share one opcode and differ only in op_type.
COERCE = {"cvar": 0, "cint": 2, "clng": 3, "cdbl": 5, "cdate": 7,
          "cstr": 8, "cbool": 11}

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


def _date_serial(text: str) -> float:
    """`#1/2/2003#` -> the VBA serial, days since 1899-12-30."""
    import datetime

    for pattern in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            when = datetime.datetime.strptime(text.strip(), pattern).date()
        except ValueError:
            continue
        return float((when - datetime.date(1899, 12, 30)).days)
    raise CompileError(f"unsupported date literal: #{text}#")


# --- tokenizer ---------------------------------------------------------
_TOK = re.compile(
    r"(?P<ws>\s+)"
    r"|(?P<date>#[^#]+#)"
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
            # Grouping parentheses are not free: Access records them with
            # a Paren marker after the grouped expression, so `(a + b) * a`
            # and `a + b * a` differ in more than precedence.
            e = self.expr()
            if not self.accept_op(")"):
                raise CompileError("missing )")
            return [*e, _ins(OP["Paren"])]
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
        if k == "date":
            # A VBA date literal is a double: days since 1899-12-30.
            raw = struct.pack("<d", _date_serial(v[1:-1]))
            return [_ins(OP["LitDate"],
                         tuple(("0x", w) for w in struct.unpack("<4H", raw)))]
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
                args, count = self.arg_list()
                obj = _ins(OP["Ld"], (("name", _res(self.names, v)),))
                if not had_parens:
                    return [obj, _ins(OP["MemLd"],
                                      (("name", _res(self.names, member[1])),))]
                return [*args, obj,
                        _ins(OP["ArgsMemLd"],
                             (("name", _res(self.names, member[1])),
                              ("0x", count)))]
            if nxt_kind == "op" and nxt == "(":
                # Foo(a, b) -- push the arguments, then call by name.
                args, count = self.arg_list()
                low = v.lower()
                if low in BUILTIN_OPCODE:
                    return [*args, _ins(OP[BUILTIN_OPCODE[low]])]
                if low in COERCE:
                    return [*args, _ins(OP["Coerce"], op_type=COERCE[low])]
                if low in BOUND_OPCODE:
                    # A dimension is an operand, not a stack argument, so
                    # only the array itself is pushed. Anything but the
                    # first dimension is refused rather than mis-encoded.
                    if count > 1:
                        raise CompileError(
                            f"{v}(array, dimension) is not supported; only "
                            "the first dimension is understood")
                    return [*args, _ins(OP[BOUND_OPCODE[low]], (("0x", 0),))]
                if low == "array":
                    return [*args, _ins(OP["ArgsArray"],
                                        (("name", 2 * ARRAY_SLOT + 2),
                                         ("0x", count)))]
                if low in BUILTIN_SLOT:
                    return [*args, _ins(OP["ArgsLd"],
                                        (("name", 2 * BUILTIN_SLOT[low] + 2),
                                         ("0x", count)))]
                return [*args, _ins(OP["ArgsLd"],
                                    (("name", _res(self.names, v)),
                                     ("0x", count)))]
            if v.lower() in BUILTIN_VALUE_SLOT:
                return [_ins(OP["Ld"],
                             (("name", 2 * BUILTIN_VALUE_SLOT[v.lower()] + 2),))]
            return [_ins(OP["Ld"], (("name", _res(self.names, v)),))]
        raise CompileError("unexpected token " + repr(v))

    def arg_list(self):
        """Parse ``(a, b)`` if present, returning ``(code, count)``.

        The count is returned rather than stashed on the parser: a nested
        call such as ``DateAdd("d", 1, Now())`` parses its own argument
        list while the outer one is still open, and shared state let the
        inner count overwrite the outer.
        """
        if not self.accept_op("("):
            return [], 0
        out, count = [], 0
        if self.accept_op(")"):
            return out, 0
        while True:
            out += self.expr()
            count += 1
            if self.accept_op(","):
                continue
            if self.accept_op(")"):
                return out, count
            raise CompileError("expected ',' or ')' in argument list")


def _res(names, n):
    low = n.lower()
    if low in names:
        return names[low]
    if low in RESERVED_SLOT:
        return 2 * RESERVED_SLOT[low] + 2
    raise CompileError("unknown identifier " + repr(n)
                       + " (not in project identifier table)")


# A declaration is eight bytes of p-code and a 24-byte header record.
# The p-code half is the same whatever the type -- everything that
# distinguishes a Long from a String lives in the record, which
# `accdb_write.add_declaration` writes. `VarDefn` carries op_type 1 and a
# u32 `var_`, the record's offset from the declaration base.
DECLARATION = re.compile(
    r"(?i)^dim\s+([A-Za-z_]\w*)\s*(?:as\s+([A-Za-z_]\w*))?$")

# Anything that reserves storage. Only the plain scalar `Dim` above is
# modelled; the rest reshape the header in their own way and are listed
# here so they can be *detected* and refused rather than silently
# miscounted. An array or fixed string carries a descriptor after its
# record, `Static` shifts the whole region, and `Const` sets 0x40 in the
# record's type field -- all measured, none implemented.
DECLARES_STORAGE = re.compile(
    r"(?i)^(?:public\s+|private\s+)?(dim|static|const|redim)\b")


def is_declaration(text):
    """``(name, type)`` for a modelled `Dim` line, or None.

    Only ``Dim x`` and ``Dim x As <scalar>`` qualify. No type means
    Variant.
    """
    m = DECLARATION.match(text.strip())
    if not m:
        return None
    return m.group(1), (m.group(2) or "Variant")


def declares_storage(text):
    """True for any line that reserves storage, modelled or not.

    Used to refuse a rewrite whose module contains a declaration form the
    record model does not cover: miscounting those puts a new record on
    top of an array descriptor, and Access crashes on the result.
    """
    return bool(DECLARES_STORAGE.match(text.strip()))


def declaration_pcode(var_):
    """The p-code for one `Dim`, given its record's ``var_`` offset."""
    return b"".join(encode_instruction(x) for x in
                    (_ins(OP["Dim"]), _ins(OP["VarDefn"], (("var_", var_),),
                                           op_type=1)))


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
    if low == "do":
        return enc([_ins(OP["Do"])])
    if low.startswith("debug.print "):
        # Debug.Print is special-cased by the compiler rather than being a
        # member call: object, then the value, then the newline marker.
        return enc([_ins(OP["Debug"]), _ins(OP["PrintObj"]),
                    *E(s[12:]), _ins(OP["PrintItemNL"])])
    if low == "wend":
        return enc([_ins(OP["Wend"])])
    if low == "end select":
        return enc([_ins(OP["EndSelect"])])
    if low == "end with":
        return enc([_ins(OP["EndWith"])])
    if low == "resume next":
        # `Resume~1` with a null name operand, measured; plain `Resume`
        # and `Resume <label>` are different shapes and stay unsupported.
        return enc([_ins(OP["Resume"], (("name", 0),), op_type=1)])
    if low.startswith("with "):
        return enc([_ins(OP["StartWithExpr"]), *E(s[5:]), _ins(OP["With"])])
    if low.startswith("erase "):
        return enc([*E(s[6:]), _ins(OP["Erase"], (("0x", 1),))])

    m = re.match(r"(?i)^on\s+error\s+goto\s+([A-Za-z_]\w*)$", s)
    if m:
        return enc([_ins(OP["OnError"], (("name", _res(names, m.group(1))),))])

    # A line label is a bare name followed by a colon.
    m = re.match(r"^([A-Za-z_]\w*):$", s)
    if m:
        return enc([_ins(OP["Label"], (("name", _res(names, m.group(1))),))])

    # A one-line If wraps its consequent in the same p-code line.
    m = re.match(r"(?i)^if\s+(.+?)\s+then\s+(.+)$", s)
    if m:
        inner = compile_line(m.group(2), names)
        if inner is None:
            raise CompileError(
                f"not a statement after Then: {m.group(2)!r}")
        return enc([*E(m.group(1)), _ins(OP["If"]),
                    _ins(OP["BoSImplicit"])]) + inner + enc(
                        [_ins(OP["EndIf"])])

    # `.Member args` inside a With block: the object is implicit.
    m = re.match(r"^\.([A-Za-z_]\w*)\s*(.*)$", s)
    if m:
        member, rest = m.group(1), m.group(2).strip()
        if rest.startswith("(") and rest.endswith(")"):
            rest = rest[1:-1].strip()
        args, count = [], 0
        if rest:
            for piece in _split_args(rest):
                args += E(piece)
                count += 1
        return enc([*args, _ins(OP["ArgsMemCallWith"],
                                (("name", _res(names, member)), ("0x", count)),
                                op_type=STATEMENT_CALL_OP_TYPE)])
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
    "with", "erase", "on", "error", "goto", "resume",
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
    if is_declaration(text):
        # `Dim x As Long` introduces `x`, which the caller adds itself;
        # the type name and `Dim`/`As` are grammar, not identifiers.
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
