"""The VBA parser: tokens in, a :class:`~pyopenvba.interpreter._ast.Module` out.

A module is parsed whole when it is loaded, the way the VBA compiler
does it, so bad source is refused before anything runs rather than when
execution happens to reach it.

Statements this understands but does not implement -- the file I/O verbs
and the handful of others listed in ``_UNSUPPORTED`` -- parse into an
:class:`~pyopenvba.interpreter._ast.Unsupported` node.  They are a gap in
pyOpenVBA, not bad VBA, so they are kept apart from a compile error and
reported only if execution reaches them.
"""

from __future__ import annotations

from typing import Final, NoReturn

from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._lex import Token, strip_suffix, tokenize
from pyopenvba.interpreter._values import (
    EMPTY,
    NOTHING,
    NULL,
    PRINT_CONTINUE,
    PRINT_ZONE,
    VBAInt,
    VBASingle,
    parse_date_text,
)

# --- the words the grammar reserves ------------------------------------------------

#: Words that may open a statement.  Anything else opening one is a call.
STATEMENT_KEYWORDS: Final = frozenset(
    """
    dim redim static private public global friend const set let if for do while select with
    exit goto gosub return on resume erase error stop end option type enum declare event
    raiseevent implements attribute open close print write input line get put seek lock
    unlock name kill mkdir rmdir chdir chdrive filecopy setattr reset width randomize beep
    sendkeys appactivate lset rset mid midb wend loop next case else elseif then call
    property sub function deftype
    """.split()
)

#: Statements that are real VBA and that this does not run.
_UNSUPPORTED: Final[dict[str, str]] = {
    "open": "file I/O",
    "close": "file I/O",
    "print": "file I/O",
    "write": "file I/O",
    "input": "file I/O",
    "get": "file I/O",
    "put": "file I/O",
    "seek": "file I/O",
    "lock": "file I/O",
    "unlock": "file I/O",
    "reset": "file I/O",
    "width": "file I/O",
    "kill": "file system access",
    "mkdir": "file system access",
    "rmdir": "file system access",
    "chdir": "file system access",
    "chdrive": "file system access",
    "filecopy": "file system access",
    "setattr": "file system access",
    "name": "file system access",
    "sendkeys": "sending keystrokes to a window",
    "appactivate": "activating another application's window",
    "beep": "the PC speaker",
    "lset": "LSet",
    "rset": "RSet",
}

#: The types a declaration may name, beyond the classes a host provides.
BUILTIN_TYPES: Final = frozenset(
    """
    boolean byte integer long longlong longptr single double currency decimal date string
    object variant any
    """.split()
)

_ASSIGN_OPS: Final = frozenset({"=", "<>", "<", ">", "<=", ">="})

#: Binary operators, tightest binding first.
_PRECEDENCE: Final[dict[str, int]] = {
    "^": 12,
    "*": 10,
    "/": 10,
    "\\": 9,
    "mod": 8,
    "+": 7,
    "-": 7,
    "&": 6,
    "=": 5,
    "<>": 5,
    "<": 5,
    ">": 5,
    "<=": 5,
    ">=": 5,
    "is": 5,
    "like": 5,
    "and": 3,
    "or": 2,
    "xor": 1,
    "eqv": 0,
    "imp": -1,
}

_RIGHT_ASSOCIATIVE: Final = frozenset({"^"})


class Parser:
    """A recursive-descent parser over one module's tokens."""

    def __init__(self, tokens: list[Token], *, module: str = "", kind: str = "standard") -> None:
        self.tokens = tokens
        self.at = 0
        self.module_name = module
        self.module_kind = kind
        self._pending_next: list[str] = []
        #: Greater than zero while parsing the statements a single-line
        #: If carries, where the line's end closes the If rather than
        #: the statement.
        self._inline = 0
        #: How many With blocks are open, which decides whether a
        #: spaced leading dot opens an argument.
        self._with_depth = 0

    # --- token helpers -------------------------------------------------------------

    @property
    def token(self) -> Token:
        return self.tokens[self.at]

    def peek(self, ahead: int = 1) -> Token:
        index = min(self.at + ahead, len(self.tokens) - 1)
        return self.tokens[index]

    def advance(self) -> Token:
        token = self.tokens[self.at]
        if token.kind != "eof":
            self.at += 1
        return token

    def at_end(self) -> bool:
        return self.token.kind == "eof"

    def at_word(self, *words: str) -> bool:
        return self.token.kind == "ident" and self.token.lower in words

    def at_op(self, *ops: str) -> bool:
        return self.token.kind == "op" and self.token.text in ops

    def accept_word(self, *words: str) -> bool:
        if self.at_word(*words):
            self.advance()
            return True
        return False

    def accept_op(self, *ops: str) -> bool:
        if self.at_op(*ops):
            self.advance()
            return True
        return False

    def expect_word(self, word: str) -> Token:
        if not self.at_word(word):
            self.fail(f"expected {word!r} but found {self.token.text!r}")
        return self.advance()

    def expect_op(self, op: str) -> Token:
        if not self.at_op(op):
            self.fail(f"expected {op!r} but found {self.token.text!r}")
        return self.advance()

    def expect_name(self) -> str:
        if self.token.kind != "ident":
            self.fail(f"expected a name but found {self.token.text!r}")
        return self.advance().text

    def at_eos(self) -> bool:
        return self.token.kind in ("eos", "eof")

    def end_statement(self) -> None:
        if self.token.kind == "eos":
            # Inside a single-line If, the newline belongs to the If: it
            # is what ends the whole construct, not this statement.
            if self._inline and self.token.text == "\n":
                return
            self.advance()
            return
        if self.token.kind == "eof":
            return
        if self._inline and self.at_word("else"):
            return
        self.fail(f"unexpected {self.token.text!r} at the end of a statement")

    def skip_blank(self) -> None:
        while self.token.kind == "eos":
            self.advance()

    def fail(self, message: str) -> NoReturn:
        where = f"{self.module_name} line {self.token.line}" if self.module_name else f"line {self.token.line}"
        raise VBACompileError(message, where=where)

    # --- module --------------------------------------------------------------------

    def parse_module(self, source: str = "") -> A.Module:
        module = A.Module(name=self.module_name, kind=self.module_kind, source=source)
        self.skip_blank()
        while not self.at_end():
            self._parse_module_member(module)
            self.skip_blank()
        if not module.name:
            module.name = module.attributes.get("VB_Name", "")
        return module

    def _parse_module_member(self, module: A.Module) -> None:
        line = self.token.line
        scope = "public"
        static = False
        while self.at_word("public", "private", "global", "friend", "static"):
            word = self.advance().lower
            if word == "static":
                static = True
            elif word == "friend":
                scope = "public"
            elif word == "global":
                scope = "public"
            else:
                scope = word

        if self.at_word("sub", "function", "property"):
            module.procedures.append(self._parse_procedure(scope=scope, static=static, line=line))
            return
        if self.at_word("declare"):
            module.declares.append(self._parse_declare(line))
            return
        if self.at_word("type"):
            module.types.append(self._parse_type(scope, line))
            return
        if self.at_word("enum"):
            module.enums.append(self._parse_enum(scope, line))
            return
        if self.at_word("event"):
            module.events.append(self._parse_event(line))
            return
        if self.at_word("const"):
            self.advance()
            module.constants.append(self._parse_const(scope, line))
            return
        if self.at_word("implements"):
            self.advance()
            module.implements.append(self.expect_name())
            self.end_statement()
            return
        if self.at_word("attribute"):
            self.advance()
            name = self.expect_name()
            while self.accept_op("."):
                name += "." + self.expect_name()
            self.expect_op("=")
            value = self.advance()
            while not self.at_eos():
                self.advance()
            module.attributes[name] = value.text
            self.end_statement()
            return
        if self.at_word("option"):
            self.advance()
            words: list[str] = []
            while not self.at_eos():
                words.append(self.advance().text)
            module.options.append(" ".join(words))
            self.end_statement()
            return
        if self.at_word("dim", "static"):
            self.advance()
            module.variables.append(self._parse_dim(scope="private" if scope == "public" else scope, line=line))
            return
        if self.at_word("deftype") or self._at_def_type():
            # DefInt A-Z and friends: recognised, and their effect on
            # untyped names is not modelled.
            while not self.at_eos():
                self.advance()
            self.end_statement()
            return
        if self.token.kind == "ident":
            module.variables.append(self._parse_dim(scope=scope, line=line, static=static))
            return
        self.fail(f"unexpected {self.token.text!r} outside a procedure")

    def _at_def_type(self) -> bool:
        if self.token.kind != "ident":
            return False
        lower = self.token.lower
        return lower.startswith("def") and lower[3:] in (
            "int",
            "lng",
            "sng",
            "dbl",
            "cur",
            "str",
            "byte",
            "bool",
            "date",
            "obj",
            "var",
            "dec",
            "lnglng",
        )

    # --- declarations ---------------------------------------------------------------

    def _parse_dim(self, *, scope: str, line: int, static: bool = False) -> A.Dim:
        decls: list[A.VarDecl] = []
        while True:
            decls.append(self._parse_var_decl())
            if not self.accept_op(","):
                break
        self.end_statement()
        return A.Dim(line=line, decls=decls, scope=scope, static=static)

    def _parse_var_decl(self) -> A.VarDecl:
        if self.token.kind != "ident":
            self.fail(f"expected a variable name but found {self.token.text!r}")
        token = self.advance()
        name = strip_suffix(token.text)
        declared = token.declared or "Variant"
        bounds: list[tuple[A.Expr | None, A.Expr]] = []
        is_array = False
        if self.at_op("("):
            is_array = True
            self.advance()
            while not self.at_op(")"):
                first = self._parse_expression()
                if self.accept_word("to"):
                    bounds.append((first, self._parse_expression()))
                else:
                    bounds.append((None, first))
                if not self.accept_op(","):
                    break
            self.expect_op(")")
        as_new = False
        if self.accept_word("as"):
            if self.accept_word("new"):
                as_new = True
            declared = self._parse_type_name()
        return A.VarDecl(name=name, declared=declared, bounds=bounds, is_array=is_array, as_new=as_new)

    def _parse_type_name(self) -> str:
        if self.token.kind != "ident":
            self.fail(f"expected a type name but found {self.token.text!r}")
        name = self.advance().text
        while self.accept_op("."):
            name += "." + self.expect_name()
        if self.at_op("*"):
            # Fixed-length string: "String * 10".
            self.advance()
            self.advance()
        return name

    def _parse_const(self, scope: str, line: int) -> A.ConstDecl:
        names: list[tuple[str, str, A.Expr]] = []
        while True:
            token = self.advance()
            if token.kind != "ident":
                self.fail(f"expected a constant name but found {token.text!r}")
            declared = token.declared or "Variant"
            if self.accept_word("as"):
                declared = self._parse_type_name()
            self.expect_op("=")
            names.append((strip_suffix(token.text), declared, self._parse_expression()))
            if not self.accept_op(","):
                break
        self.end_statement()
        return A.ConstDecl(line=line, names=names, scope=scope)

    def _parse_type(self, scope: str, line: int) -> A.TypeDef:
        self.expect_word("type")
        defined = A.TypeDef(name=self.expect_name(), scope=scope, line=line)
        self.end_statement()
        self.skip_blank()
        while not (self.at_word("end") and self.peek().kind == "ident" and self.peek().lower == "type"):
            if self.at_end():
                self.fail("a Type block is not closed")
            field = self._parse_var_decl()
            defined.fields.append(
                A.TypeField(
                    name=field.name,
                    declared=field.declared,
                    bounds=field.bounds,
                    is_array=field.is_array,
                )
            )
            self.end_statement()
            self.skip_blank()
        self.advance()
        self.advance()
        self.end_statement()
        return defined

    def _parse_enum(self, scope: str, line: int) -> A.EnumDef:
        self.expect_word("enum")
        defined = A.EnumDef(name=self.expect_name(), scope=scope, line=line)
        self.end_statement()
        self.skip_blank()
        while not (self.at_word("end") and self.peek().kind == "ident" and self.peek().lower == "enum"):
            if self.at_end():
                self.fail("an Enum block is not closed")
            name = self.expect_name()
            value = self._parse_expression() if self.accept_op("=") else None
            defined.members.append((strip_suffix(name), value))
            self.end_statement()
            self.skip_blank()
        self.advance()
        self.advance()
        self.end_statement()
        return defined

    def _parse_event(self, line: int) -> A.EventDef:
        self.expect_word("event")
        name = self.expect_name()
        params = self._parse_params()
        self.end_statement()
        return A.EventDef(name=name, params=params, line=line)

    def _parse_declare(self, line: int) -> A.DeclareDef:
        self.expect_word("declare")
        self.accept_word("ptrsafe")
        is_function = self.at_word("function")
        if not self.accept_word("function") and not self.accept_word("sub"):
            self.fail("Declare has to name a Sub or a Function")
        name = self.expect_name()
        library = ""
        alias = ""
        if self.accept_word("lib") and self.token.kind == "string":
            library = self.advance().text
        if self.accept_word("alias") and self.token.kind == "string":
            alias = self.advance().text
        params = self._parse_params()
        returns = "Variant"
        if self.accept_word("as"):
            returns = self._parse_type_name()
        self.end_statement()
        return A.DeclareDef(
            name=strip_suffix(name),
            library=library,
            alias=alias,
            is_function=is_function,
            params=params,
            returns=returns,
            line=line,
        )

    def _parse_params(self) -> list[A.Param]:
        params: list[A.Param] = []
        if not self.accept_op("("):
            return params
        while not self.at_op(")"):
            param = A.Param()
            while self.at_word("optional", "byval", "byref", "paramarray"):
                word = self.advance().lower
                if word == "optional":
                    param.optional = True
                elif word == "byval":
                    param.by_val = True
                elif word == "paramarray":
                    param.param_array = True
            token = self.advance()
            if token.kind != "ident":
                self.fail(f"expected a parameter name but found {token.text!r}")
            param.name = strip_suffix(token.text)
            param.declared = token.declared or "Variant"
            if self.at_op("("):
                self.advance()
                self.expect_op(")")
                param.is_array = True
            if self.accept_word("as"):
                param.declared = self._parse_type_name()
            if self.accept_op("="):
                param.default = self._parse_expression()
                param.optional = True
            params.append(param)
            if not self.accept_op(","):
                break
        self.expect_op(")")
        return params

    def _parse_procedure(self, *, scope: str, static: bool, line: int) -> A.Procedure:
        word = self.advance().lower
        kind = word
        if word == "property":
            if not self.at_word("get", "let", "set"):
                self.fail("Property has to be followed by Get, Let or Set")
            kind = self.advance().lower
        name_token = self.advance()
        if name_token.kind != "ident":
            self.fail(f"expected a procedure name but found {name_token.text!r}")
        procedure = A.Procedure(
            kind=kind,
            name=strip_suffix(name_token.text),
            scope=scope,
            static=static,
            line=line,
            returns=name_token.declared or "Variant",
        )
        procedure.params = self._parse_params()
        if self.accept_word("as"):
            procedure.returns = self._parse_type_name()
        self.end_statement()
        closing = "sub" if kind == "sub" else ("function" if kind == "function" else "property")
        procedure.body = self._parse_block(ends=(closing,))
        self.expect_word("end")
        self.expect_word(closing)
        self.end_statement()
        for statement in procedure.body:
            if isinstance(statement, A.Unsupported) and statement.text.lower().startswith("attribute "):
                continue
        return procedure

    # --- statements ------------------------------------------------------------------

    def _parse_block(self, *, ends: tuple[str, ...]) -> list[A.Stmt]:
        """Statements up to, but not including, whichever ender comes first."""
        body: list[A.Stmt] = []
        while True:
            self.skip_blank()
            if self.at_end():
                if ends:
                    self.fail(f"a block opened earlier is never closed by End {ends[0].capitalize()}")
                return body
            if self._at_block_end(ends):
                return body
            if ends and self.at_word("end") and self.peek().kind == "ident":
                closing = self.peek().lower
                if closing in ("sub", "function", "property") and closing not in ends:
                    opener = ends[0].capitalize()
                    self.fail(
                        f"End {self.peek().text} is here, and the {opener} block above it "
                        f"has no End {opener}"
                    )
            body.append(self._parse_statement())

    def _at_block_end(self, ends: tuple[str, ...]) -> bool:
        if self._pending_next and "for" in ends:
            # A shared ``Next j, i`` has already closed this loop.
            return True
        if self.token.kind != "ident":
            return False
        lower = self.token.lower
        if lower == "end":
            nxt = self.peek()
            return nxt.kind == "ident" and nxt.lower in ends
        if lower in ("else", "elseif") and "if" in ends:
            return True
        if lower == "case" and "select" in ends:
            return True
        if lower == "next" and "for" in ends:
            return True
        if lower == "loop" and "do" in ends:
            return True
        if lower == "wend" and "while" in ends:
            return True
        return False

    def _parse_statement(self) -> A.Stmt:
        line = self.token.line
        token = self.token

        if token.kind == "number" and self.peek().kind not in ("eos", "eof"):
            # An old-style line number, which is a label.
            self.advance()
            return A.Label(line=line, name=token.text)

        if token.kind == "ident":
            lower = token.lower
            if self.peek().kind == "eos" and self.peek().text == ":" and lower not in STATEMENT_KEYWORDS:
                self.advance()
                self.advance()
                return A.Label(line=line, name=token.text)
            handler = getattr(self, f"_stmt_{lower}", None)
            if handler is not None and lower in STATEMENT_KEYWORDS:
                return handler(line)
            if lower in _UNSUPPORTED and self._opens_a_statement():
                return self._unsupported(line, _UNSUPPORTED[lower])
            if lower in ("public", "private", "friend", "global"):
                self.advance()
                return self._parse_statement()

        return self._parse_call_or_assign(line)

    def _opens_a_statement(self) -> bool:
        """Whether a word like Name or Close is a statement here.

        Every one of them is also an ordinary name: ``Name`` is a
        property on half the object model, and ``Close`` is a method.
        What separates them is what comes next, since the statement
        forms are all followed by a value and the others by an
        operator.
        """
        nxt = self.peek()
        if nxt.kind in ("string", "number", "date"):
            return True
        if nxt.kind == "op":
            return nxt.text == "#"
        if nxt.kind == "ident":
            return True
        return False

    def _unsupported(self, line: int, reason: str) -> A.Unsupported:
        words: list[str] = []
        while not self.at_eos():
            words.append(self.advance().text)
        self.end_statement()
        return A.Unsupported(line=line, text=" ".join(words), reason=reason)

    # Declarations inside a procedure.

    def _stmt_dim(self, line: int) -> A.Stmt:
        self.advance()
        return self._parse_dim(scope="local", line=line)

    def _stmt_static(self, line: int) -> A.Stmt:
        self.advance()
        if self.at_word("sub", "function", "property"):
            self.fail("a procedure cannot be declared inside another procedure")
        return self._parse_dim(scope="local", line=line, static=True)

    def _stmt_const(self, line: int) -> A.Stmt:
        self.advance()
        return self._parse_const("local", line)

    def _stmt_redim(self, line: int) -> A.Stmt:
        self.advance()
        preserve = self.accept_word("preserve")
        decls: list[A.VarDecl] = []
        while True:
            decls.append(self._parse_var_decl())
            if not self.accept_op(","):
                break
        self.end_statement()
        return A.ReDim(line=line, preserve=preserve, decls=decls)

    def _stmt_erase(self, line: int) -> A.Stmt:
        self.advance()
        names: list[A.Expr] = []
        while True:
            names.append(self._parse_expression())
            if not self.accept_op(","):
                break
        self.end_statement()
        return A.Erase(line=line, names=names)

    # Assignment.

    def _stmt_set(self, line: int) -> A.Stmt:
        self.advance()
        target = self._parse_target()
        self.expect_op("=")
        value = self._parse_expression()
        self.end_statement()
        return A.Assign(line=line, target=target, value=value, kind="set")

    def _stmt_let(self, line: int) -> A.Stmt:
        self.advance()
        target = self._parse_target()
        self.expect_op("=")
        value = self._parse_expression()
        self.end_statement()
        return A.Assign(line=line, target=target, value=value, kind="let")

    def _parse_target(self) -> A.Expr:
        """The left of an assignment, where ``=`` ends it rather than compares."""
        return self._parse_postfix(self._parse_primary())

    def _stmt_mid(self, line: int) -> A.Stmt:
        return self._mid_assignment(line, bytes_wide=False)

    def _stmt_midb(self, line: int) -> A.Stmt:
        return self._mid_assignment(line, bytes_wide=True)

    def _mid_assignment(self, line: int, *, bytes_wide: bool) -> A.Stmt:
        start_at = self.at
        self.advance()
        if not self.at_op("("):
            self.at = start_at
            return self._parse_call_or_assign(line)
        self.advance()
        target = self._parse_expression()
        self.expect_op(",")
        start = self._parse_expression()
        length = self._parse_expression() if self.accept_op(",") else None
        self.expect_op(")")
        if not self.at_op("="):
            self.at = start_at
            return self._parse_call_or_assign(line)
        self.advance()
        value = self._parse_expression()
        self.end_statement()
        return A.MidAssign(
            line=line, target=target, start=start, length=length, value=value, bytes_wide=bytes_wide
        )

    # Control flow.

    def _stmt_if(self, line: int) -> A.Stmt:
        self.advance()
        condition = self._parse_expression()
        if not self.accept_word("then"):
            self.fail("If has to be followed by Then")
        if not self.at_eos():
            return self._single_line_if(line, condition)
        self.end_statement()
        statement = A.If(line=line, branches=[A.IfBranch(condition=condition, body=self._parse_block(ends=("if",)))])
        while self.at_word("elseif"):
            self.advance()
            branch_condition = self._parse_expression()
            if not self.accept_word("then"):
                self.fail("ElseIf has to be followed by Then")
            if not self.at_eos():
                # ElseIf x Then stmt -- the rest of the chain is on one line.
                inline = self._single_line_if(self.token.line, branch_condition)
                statement.branches.extend(inline.branches)
                statement.otherwise = inline.otherwise
                return statement
            self.end_statement()
            statement.branches.append(
                A.IfBranch(condition=branch_condition, body=self._parse_block(ends=("if",)))
            )
        if self.at_word("else"):
            self.advance()
            self.end_statement()
            statement.otherwise = self._parse_block(ends=("if",))
        self.expect_word("end")
        self.expect_word("if")
        self.end_statement()
        return statement

    def _single_line_if(self, line: int, condition: A.Expr) -> A.If:
        """``If x Then a: b Else c``, which ends where its line does."""
        body: list[A.Stmt] = []
        otherwise: list[A.Stmt] = []
        target = body
        self._inline += 1
        try:
            while not self.at_eos():
                if self.at_word("else"):
                    self.advance()
                    target = otherwise
                    continue
                target.append(self._parse_statement())
                if self.token.kind == "eos" and self.token.text == ":":
                    self.advance()
        finally:
            self._inline -= 1
        self.end_statement()
        return A.If(line=line, branches=[A.IfBranch(condition=condition, body=body)], otherwise=otherwise)

    def _stmt_for(self, line: int) -> A.Stmt:
        self.advance()
        if self.at_word("each"):
            self.advance()
            target = self._parse_target()
            if not self.accept_word("in"):
                self.fail("For Each has to be followed by In")
            collection = self._parse_expression()
            self.end_statement()
            body = self._parse_block(ends=("for",))
            self._close_next(target)
            return A.ForEach(line=line, target=target, collection=collection, body=body)
        target = self._parse_target()
        self.expect_op("=")
        start = self._parse_expression()
        if not self.accept_word("to"):
            self.fail("For has to be followed by To")
        limit = self._parse_expression()
        step = self._parse_expression() if self.accept_word("step") else None
        self.end_statement()
        body = self._parse_block(ends=("for",))
        self._close_next(target)
        return A.For(line=line, target=target, start=start, limit=limit, step=step, body=body)

    def _close_next(self, target: A.Expr) -> None:
        """Consume the Next that closes this loop, sharing a Next i, j."""
        if self._pending_next:
            name = self._pending_next.pop(0)
            if name and isinstance(target, A.Name) and name.lower() != target.name.lower():
                self.fail(f"Next {name} does not match For {target.name}")
            return
        self.expect_word("next")
        names: list[str] = []
        while self.token.kind == "ident":
            names.append(self.advance().text)
            if not self.accept_op(","):
                break
        self.end_statement()
        if names:
            first = names[0]
            if isinstance(target, A.Name) and strip_suffix(first).lower() != target.name.lower():
                self.fail(f"Next {first} does not match For {target.name}")
            self._pending_next = [strip_suffix(name) for name in names[1:]]

    def _stmt_do(self, line: int) -> A.Stmt:
        self.advance()
        test = ""
        condition: A.Expr | None = None
        if self.at_word("while", "until"):
            test = self.advance().lower
            condition = self._parse_expression()
        self.end_statement()
        body = self._parse_block(ends=("do",))
        self.expect_word("loop")
        at_end = False
        if self.at_word("while", "until"):
            if test:
                self.fail("a Do loop cannot be tested at both ends")
            test = self.advance().lower
            condition = self._parse_expression()
            at_end = True
        self.end_statement()
        return A.DoLoop(line=line, body=body, condition=condition, test=test, at_end=at_end)

    def _stmt_while(self, line: int) -> A.Stmt:
        self.advance()
        condition = self._parse_expression()
        self.end_statement()
        body = self._parse_block(ends=("while",))
        self.expect_word("wend")
        self.end_statement()
        return A.WhileLoop(line=line, condition=condition, body=body)

    def _stmt_with(self, line: int) -> A.Stmt:
        self.advance()
        subject = self._parse_expression()
        self.end_statement()
        self._with_depth += 1
        try:
            body = self._parse_block(ends=("with",))
        finally:
            self._with_depth -= 1
        self.expect_word("end")
        self.expect_word("with")
        self.end_statement()
        return A.With(line=line, subject=subject, body=body)

    def _stmt_select(self, line: int) -> A.Stmt:
        self.advance()
        if not self.accept_word("case"):
            self.fail("Select has to be followed by Case")
        subject = self._parse_expression()
        self.end_statement()
        self.skip_blank()
        statement = A.SelectCase(line=line, subject=subject)
        while self.at_word("case"):
            self.advance()
            if self.accept_word("else"):
                self.end_statement()
                statement.otherwise = self._parse_block(ends=("select",))
                break
            clauses = [self._parse_case_clause()]
            while self.accept_op(","):
                clauses.append(self._parse_case_clause())
            self.end_statement()
            statement.blocks.append(A.CaseBlock(clauses=clauses, body=self._parse_block(ends=("select",))))
        self.expect_word("end")
        self.expect_word("select")
        self.end_statement()
        return statement

    def _parse_case_clause(self) -> A.CaseClause:
        if self.at_word("is"):
            self.advance()
            if self.token.kind != "op" or self.token.text not in _ASSIGN_OPS:
                self.fail("Case Is has to be followed by a comparison")
            op = self.advance().text
            return A.CaseClause(kind="compare", op=op, value=self._parse_expression())
        if self.token.kind == "op" and self.token.text in ("<", ">", "<=", ">=", "<>", "="):
            op = self.advance().text
            return A.CaseClause(kind="compare", op=op, value=self._parse_expression())
        first = self._parse_expression()
        if self.accept_word("to"):
            return A.CaseClause(kind="range", value=first, upper=self._parse_expression())
        return A.CaseClause(kind="value", value=first)

    # Jumps and error handling.

    def _stmt_exit(self, line: int) -> A.Stmt:
        self.advance()
        if self.token.kind != "ident":
            self.fail("Exit has to name what it leaves")
        what = self.advance().lower
        if what not in ("sub", "function", "property", "for", "do"):
            self.fail(f"Exit {what} is not a thing VBA can leave")
        self.end_statement()
        return A.ExitStmt(line=line, what=what)

    def _stmt_goto(self, line: int) -> A.Stmt:
        self.advance()
        label = self.advance().text
        self.end_statement()
        return A.GoTo(line=line, label=label)

    def _stmt_gosub(self, line: int) -> A.Stmt:
        self.advance()
        label = self.advance().text
        self.end_statement()
        return A.GoTo(line=line, label=label, gosub=True)

    def _stmt_return(self, line: int) -> A.Stmt:
        self.advance()
        self.end_statement()
        return A.ReturnStmt(line=line)

    def _stmt_on(self, line: int) -> A.Stmt:
        self.advance()
        if self.accept_word("error"):
            if self.accept_word("resume"):
                if not self.accept_word("next"):
                    self.fail("On Error Resume has to be followed by Next")
                self.end_statement()
                return A.OnError(line=line, mode="resume_next")
            if not self.accept_word("goto"):
                self.fail("On Error has to be followed by GoTo or Resume Next")
            label = self.advance()
            self.end_statement()
            if label.kind == "number" and label.text == "0":
                return A.OnError(line=line, mode="clear")
            return A.OnError(line=line, mode="goto", label=label.text)
        value = self._parse_expression()
        gosub = False
        if self.accept_word("gosub"):
            gosub = True
        elif not self.accept_word("goto"):
            self.fail("On has to be followed by GoTo or GoSub")
        labels: list[str] = []
        while True:
            labels.append(self.advance().text)
            if not self.accept_op(","):
                break
        self.end_statement()
        return A.OnGoto(line=line, value=value, labels=labels, gosub=gosub)

    def _stmt_resume(self, line: int) -> A.Stmt:
        self.advance()
        if self.at_eos():
            self.end_statement()
            return A.Resume(line=line, mode="same")
        if self.accept_word("next"):
            self.end_statement()
            return A.Resume(line=line, mode="next")
        label = self.advance()
        self.end_statement()
        if label.kind == "number" and label.text == "0":
            return A.Resume(line=line, mode="same")
        return A.Resume(line=line, mode="label", label=label.text)

    def _stmt_error(self, line: int) -> A.Stmt:
        self.advance()
        if self.at_eos():
            self.fail("Error has to name a number")
        number = self._parse_expression()
        self.end_statement()
        return A.ErrorStmt(line=line, number=number)

    def _stmt_raiseevent(self, line: int) -> A.Stmt:
        self.advance()
        name = self.expect_name()
        args: list[A.Argument] = []
        if self.accept_op("("):
            args = self._parse_arguments(closing=True)
        self.end_statement()
        return A.RaiseEvent(line=line, name=name, args=args)

    def _stmt_stop(self, line: int) -> A.Stmt:
        self.advance()
        self.end_statement()
        return A.Stop(line=line)

    def _stmt_end(self, line: int) -> A.Stmt:
        self.advance()
        if self.token.kind == "ident":
            self.fail(f"End {self.token.text} does not close anything that is open")
        self.end_statement()
        return A.EndStmt(line=line)

    def _stmt_option(self, line: int) -> A.Stmt:
        self.advance()
        words: list[str] = []
        while not self.at_eos():
            words.append(self.advance().text)
        self.end_statement()
        return A.OptionStmt(line=line, text=" ".join(words))

    def _stmt_call(self, line: int) -> A.Stmt:
        self.advance()
        callee = self._parse_expression()
        args: list[A.Argument] = []
        if isinstance(callee, A.Index):
            args = callee.args
            callee = callee.target if callee.target is not None else callee
        self.end_statement()
        return A.CallStmt(line=line, callee=callee, args=args, explicit=True)

    def _stmt_line(self, line: int) -> A.Stmt:
        # "Line Input #1, s" is file I/O; "Line" alone is a drawing method
        # on a Form, which this does not have either.
        return self._unsupported(line, "file I/O")

    def _stmt_randomize(self, line: int) -> A.Stmt:
        self.advance()
        seed = None if self.at_eos() else self._parse_expression()
        self.end_statement()
        return A.CallStmt(
            line=line,
            callee=A.Name(line=line, name="Randomize"),
            args=[] if seed is None else [A.Argument(value=seed)],
        )

    def _stmt_implements(self, line: int) -> A.Stmt:
        return self._unsupported(line, "Implements")

    def _stmt_declare(self, line: int) -> A.Stmt:
        self.fail("Declare belongs outside a procedure")

    def _stmt_attribute(self, line: int) -> A.Stmt:
        words: list[str] = []
        while not self.at_eos():
            words.append(self.advance().text)
        self.end_statement()
        return A.Unsupported(line=line, text=" ".join(words), reason="a procedure attribute")

    # Calls and assignments.

    def _parse_call_or_assign(self, line: int) -> A.Stmt:
        target = self._parse_postfix(self._parse_primary(), statement=True)
        if self.at_op("="):
            self.advance()
            value = self._parse_expression()
            self.end_statement()
            return A.Assign(line=line, target=target, value=value, kind="let")
        args: list[A.Argument] = []
        if _is_print(target) and not (self._inline and self.at_word("else")):
            args = self._parse_print_arguments()
        elif isinstance(target, A.Index) and self.at_eos():
            args = target.args
            inner = target.target
            if inner is not None:
                target = inner
        elif not self.at_eos():
            args = self._parse_arguments(closing=False)
        self.end_statement()
        return A.CallStmt(line=line, callee=target, args=args)

    def _parse_print_arguments(self) -> list[A.Argument]:
        """Print's own list, where ``;`` joins and ``,`` moves a zone on."""
        args: list[A.Argument] = []
        line = self.token.line
        while not self.at_eos():
            if self._inline and self.at_word("else"):
                break
            if self.at_op(";"):
                self.advance()
                if self.at_eos():
                    args.append(A.Argument(value=A.Literal(line=line, value=PRINT_CONTINUE)))
                continue
            if self.at_op(","):
                self.advance()
                args.append(A.Argument(value=A.Literal(line=line, value=PRINT_ZONE)))
                if self.at_eos():
                    args.append(A.Argument(value=A.Literal(line=line, value=PRINT_CONTINUE)))
                continue
            args.append(A.Argument(value=self._parse_expression()))
        return args

    # --- expressions ------------------------------------------------------------------

    def expression(self) -> A.Expr:
        """One whole expression, for a caller parsing nothing else."""
        return self._parse_expression()

    def _parse_expression(self, minimum: int = -2) -> A.Expr:
        left = self._parse_unary()
        while True:
            token = self.token
            op = ""
            if token.kind == "op" and token.text in _PRECEDENCE:
                op = token.text
            elif token.kind == "ident" and token.lower in _PRECEDENCE:
                op = token.lower
            if not op:
                return left
            precedence = _PRECEDENCE[op]
            if precedence < minimum:
                return left
            line = self.advance().line
            if op == "is" and self.at_word("not"):
                self.advance()
                right = self._parse_expression(precedence + 1)
                left = A.Unary(line=line, op="not", operand=A.Binary(line=line, op="is", left=left, right=right))
                continue
            nxt = precedence if op in _RIGHT_ASSOCIATIVE else precedence + 1
            right = self._parse_expression(nxt)
            left = A.Binary(line=line, op=op, left=left, right=right)

    def _parse_unary(self) -> A.Expr:
        token = self.token
        if token.kind == "op" and token.text == "-":
            self.advance()
            return A.Unary(line=token.line, op="-", operand=self._parse_expression(_PRECEDENCE["^"]))
        if token.kind == "op" and token.text == "+":
            self.advance()
            return self._parse_unary()
        if token.kind == "ident" and token.lower == "not":
            self.advance()
            return A.Unary(line=token.line, op="not", operand=self._parse_expression(_PRECEDENCE["and"] + 1))
        if token.kind == "ident" and token.lower == "typeof":
            self.advance()
            value = self._parse_expression(_PRECEDENCE["is"] + 1)
            if not self.accept_word("is"):
                self.fail("TypeOf has to be followed by Is")
            return A.TypeOfIs(line=token.line, value=value, type_name=self._parse_type_name())
        if token.kind == "ident" and token.lower == "addressof":
            self.advance()
            return A.AddressOf(line=token.line, name=self.expect_name())
        if token.kind == "ident" and token.lower == "new":
            self.advance()
            return A.NewExpr(line=token.line, type_name=self._parse_type_name())
        return self._parse_postfix(self._parse_primary())

    def _parse_primary(self) -> A.Expr:
        token = self.token
        line = token.line
        if token.kind == "number":
            self.advance()
            return A.Literal(line=line, value=_number_value(token))
        if token.kind == "string":
            self.advance()
            return A.Literal(line=line, value=token.text)
        if token.kind == "date":
            self.advance()
            parsed = parse_date_text(token.text)
            if parsed is None:  # pragma: no cover - the lexer checked already
                self.fail(f"#{token.text}# is not a date")
            return A.Literal(line=line, value=parsed)
        if token.kind == "bracket":
            self.advance()
            return A.Bracket(line=line, text=token.text)
        if token.kind == "op" and token.text == "(":
            self.advance()
            inner = self._parse_expression()
            self.expect_op(")")
            return inner
        if token.kind == "op" and token.text == ".":
            self.advance()
            return A.Member(line=line, target=None, name=self.expect_name())
        if token.kind == "op" and token.text == "!":
            self.advance()
            return A.Bang(line=line, target=None, name=self.expect_name())
        if token.kind == "ident":
            lower = token.lower
            if lower == "nothing":
                self.advance()
                return A.Literal(line=line, value=NOTHING)
            if lower == "null":
                self.advance()
                return A.Literal(line=line, value=NULL)
            if lower == "empty":
                self.advance()
                return A.Literal(line=line, value=EMPTY)
            if lower == "true":
                self.advance()
                return A.Literal(line=line, value=True)
            if lower == "false":
                self.advance()
                return A.Literal(line=line, value=False)
            self.advance()
            return A.Name(line=line, name=token.text)
        self.fail(f"unexpected {token.text!r} where a value was expected")

    def _parse_postfix(self, target: A.Expr, *, statement: bool = False) -> A.Expr:
        while True:
            if self.at_op("."):
                if statement and self._with_depth and self.token.spaced:
                    # Inside a With block, ``Debug.Print .Count`` passes
                    # the block's .Count as an argument.  Without the
                    # space it would carry on the member chain, and the
                    # space is the only thing that separates them.
                    return target
                self.advance()
                if self.token.kind != "ident":
                    self.fail(f"expected a member name after '.' but found {self.token.text!r}")
                target = A.Member(line=self.token.line, target=target, name=self.advance().text)
                continue
            if self.at_op("!"):
                self.advance()
                if self.token.kind != "ident":
                    self.fail(f"expected a name after '!' but found {self.token.text!r}")
                target = A.Bang(line=self.token.line, target=target, name=self.advance().text)
                continue
            if self.at_op("("):
                line = self.token.line
                self.advance()
                args = self._parse_arguments(closing=True)
                target = A.Index(line=line, target=target, args=args)
                continue
            return target

    def _parse_arguments(self, *, closing: bool) -> list[A.Argument]:
        args: list[A.Argument] = []
        if closing and self.at_op(")"):
            self.advance()
            return args
        while True:
            if self.at_op(","):
                args.append(A.Argument())
                self.advance()
                continue
            if closing and self.at_op(")"):
                args.append(A.Argument())
                break
            if not closing and (self.at_eos() or (self._inline and self.at_word("else"))):
                args.append(A.Argument())
                break
            name = ""
            if self.token.kind == "ident" and self.peek().kind == "op" and self.peek().text == ":=":
                name = self.advance().text
                self.advance()
            args.append(A.Argument(value=self._parse_expression(), name=name))
            if self.accept_op(","):
                continue
            break
        if closing:
            self.expect_op(")")
        return args


def _is_print(target: A.Expr) -> bool:
    """Whether this call is a Print, whose argument list has its own shape."""
    return isinstance(target, A.Member) and target.name.lower() in ("print", "write")


def _number_value(token: Token) -> object:
    text = token.text
    declared = token.declared
    if declared in ("Integer", "Long"):
        return VBAInt(int(float(text)), declared)
    if declared == "Single":
        return VBASingle(float(text))
    if declared == "Currency":
        from pyopenvba.interpreter._values import VBACurrency

        return VBACurrency(text)
    if declared == "String":
        return text
    return float(text)


def parse_module(source: str, *, name: str = "", kind: str = "standard") -> A.Module:
    """Parse one module's source, raising VBACompileError if it is not VBA."""
    from pyopenvba.interpreter._preprocess import strip_directives

    prepared = strip_directives(source, module=name)
    return Parser(tokenize(prepared, module=name), module=name, kind=kind).parse_module(source)


def parse_expression(source: str, *, name: str = "") -> A.Expr:
    """Parse a single expression, for an immediate-window style evaluation."""
    parser = Parser(tokenize(source, module=name), module=name)
    expression = parser.expression()
    if not parser.at_eos():
        parser.fail(f"unexpected {parser.token.text!r} after the expression")
    return expression
