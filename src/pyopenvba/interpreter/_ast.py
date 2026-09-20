"""The shape of a parsed VBA module.

Every node carries the line it started on, because an error raised deep
in an expression is reported against the statement the reader can see.

The tree keeps what VBA's compiler keeps and no more.  An expression
that is a call and an expression that is an array subscript look the
same in source (``Total(3)``), so they parse to the same node and the
interpreter tells them apart from what the name turns out to be, as VBA
does.
"""

from __future__ import annotations

from dataclasses import dataclass, field


# --- expressions ------------------------------------------------------------------


@dataclass(slots=True)
class Expr:
    line: int = 0


@dataclass(slots=True)
class Literal(Expr):
    value: object = None


@dataclass(slots=True)
class Name(Expr):
    """A bare identifier: a variable, a constant, or a call with no arguments."""

    name: str = ""


@dataclass(slots=True)
class Member(Expr):
    """``target.name``, or ``.name`` inside a With block when target is None."""

    target: Expr | None = None
    name: str = ""


@dataclass(slots=True)
class Bang(Expr):
    """``target!name``: the default member, called with the name as a string."""

    target: Expr | None = None
    name: str = ""


@dataclass(slots=True)
class Index(Expr):
    """``target(args)``: a call, an array subscript or a default member."""

    target: Expr | None = None
    args: list[Argument] = field(default_factory=lambda: [])


@dataclass(slots=True)
class Bracket(Expr):
    """``[A1]``, Excel's shorthand for Evaluate."""

    text: str = ""


@dataclass(slots=True)
class Unary(Expr):
    op: str = ""
    operand: Expr | None = None


@dataclass(slots=True)
class Binary(Expr):
    op: str = ""
    left: Expr | None = None
    right: Expr | None = None


@dataclass(slots=True)
class NewExpr(Expr):
    type_name: str = ""


@dataclass(slots=True)
class TypeOfIs(Expr):
    value: Expr | None = None
    type_name: str = ""
    negated: bool = False


@dataclass(slots=True)
class AddressOf(Expr):
    name: str = ""


@dataclass(slots=True)
class Argument:
    """One argument in a call: positional, named, or left out entirely."""

    value: Expr | None = None
    name: str = ""

    @property
    def omitted(self) -> bool:
        return self.value is None


# --- statements -------------------------------------------------------------------


@dataclass(slots=True)
class Stmt:
    line: int = 0


@dataclass(slots=True)
class VarDecl:
    name: str = ""
    declared: str = "Variant"
    bounds: list[tuple[Expr | None, Expr]] = field(default_factory=lambda: [])
    is_array: bool = False
    as_new: bool = False


@dataclass(slots=True)
class Dim(Stmt):
    decls: list[VarDecl] = field(default_factory=lambda: [])
    scope: str = "dim"
    static: bool = False


@dataclass(slots=True)
class ConstDecl(Stmt):
    names: list[tuple[str, str, Expr]] = field(default_factory=lambda: [])
    scope: str = "private"


@dataclass(slots=True)
class Assign(Stmt):
    target: Expr | None = None
    value: Expr | None = None
    #: "let" for a value, "set" for a reference, "lset"/"rset" for the
    #: two padding assignments.
    kind: str = "let"


@dataclass(slots=True)
class MidAssign(Stmt):
    """``Mid(s, start, length) = text``, which overwrites in place."""

    target: Expr | None = None
    start: Expr | None = None
    length: Expr | None = None
    value: Expr | None = None
    bytes_wide: bool = False


@dataclass(slots=True)
class CallStmt(Stmt):
    callee: Expr | None = None
    args: list[Argument] = field(default_factory=lambda: [])
    explicit: bool = False


@dataclass(slots=True)
class IfBranch:
    condition: Expr | None = None
    body: list[Stmt] = field(default_factory=lambda: [])


@dataclass(slots=True)
class If(Stmt):
    branches: list[IfBranch] = field(default_factory=lambda: [])
    otherwise: list[Stmt] = field(default_factory=lambda: [])


@dataclass(slots=True)
class For(Stmt):
    target: Expr | None = None
    start: Expr | None = None
    limit: Expr | None = None
    step: Expr | None = None
    body: list[Stmt] = field(default_factory=lambda: [])


@dataclass(slots=True)
class ForEach(Stmt):
    target: Expr | None = None
    collection: Expr | None = None
    body: list[Stmt] = field(default_factory=lambda: [])


@dataclass(slots=True)
class DoLoop(Stmt):
    body: list[Stmt] = field(default_factory=lambda: [])
    condition: Expr | None = None
    #: "while" or "until"; empty for Do ... Loop with no test.
    test: str = ""
    #: True when the test is written on the Loop rather than the Do.
    at_end: bool = False


@dataclass(slots=True)
class WhileLoop(Stmt):
    condition: Expr | None = None
    body: list[Stmt] = field(default_factory=lambda: [])


@dataclass(slots=True)
class With(Stmt):
    subject: Expr | None = None
    body: list[Stmt] = field(default_factory=lambda: [])


@dataclass(slots=True)
class CaseClause:
    #: "value", "range" or "compare".
    kind: str = "value"
    value: Expr | None = None
    upper: Expr | None = None
    op: str = ""


@dataclass(slots=True)
class CaseBlock:
    clauses: list[CaseClause] = field(default_factory=lambda: [])
    body: list[Stmt] = field(default_factory=lambda: [])


@dataclass(slots=True)
class SelectCase(Stmt):
    subject: Expr | None = None
    blocks: list[CaseBlock] = field(default_factory=lambda: [])
    otherwise: list[Stmt] | None = None


@dataclass(slots=True)
class ExitStmt(Stmt):
    what: str = ""


@dataclass(slots=True)
class GoTo(Stmt):
    label: str = ""
    gosub: bool = False


@dataclass(slots=True)
class ReturnStmt(Stmt):
    pass


@dataclass(slots=True)
class Label(Stmt):
    name: str = ""


@dataclass(slots=True)
class OnError(Stmt):
    #: "goto", "resume_next" or "clear" (On Error GoTo 0).
    mode: str = "clear"
    label: str = ""


@dataclass(slots=True)
class OnGoto(Stmt):
    """``On expr GoTo l1, l2`` and its GoSub form."""

    value: Expr | None = None
    labels: list[str] = field(default_factory=lambda: [])
    gosub: bool = False


@dataclass(slots=True)
class Resume(Stmt):
    #: "same", "next" or "label".
    mode: str = "same"
    label: str = ""


@dataclass(slots=True)
class ReDim(Stmt):
    preserve: bool = False
    decls: list[VarDecl] = field(default_factory=lambda: [])


@dataclass(slots=True)
class Erase(Stmt):
    names: list[Expr] = field(default_factory=lambda: [])


@dataclass(slots=True)
class ErrorStmt(Stmt):
    number: Expr | None = None


@dataclass(slots=True)
class RaiseEvent(Stmt):
    name: str = ""
    args: list[Argument] = field(default_factory=lambda: [])


@dataclass(slots=True)
class Stop(Stmt):
    pass


@dataclass(slots=True)
class EndStmt(Stmt):
    pass


@dataclass(slots=True)
class OptionStmt(Stmt):
    text: str = ""


@dataclass(slots=True)
class Unsupported(Stmt):
    """A statement this recognises as VBA and does not implement.

    Parsed rather than refused so that a module containing one still
    loads, and so that the error names the statement when execution
    actually reaches it.
    """

    text: str = ""
    reason: str = ""


# --- declarations -----------------------------------------------------------------


@dataclass(slots=True)
class Param:
    name: str = ""
    declared: str = "Variant"
    by_val: bool = False
    optional: bool = False
    default: Expr | None = None
    param_array: bool = False
    is_array: bool = False


@dataclass(slots=True)
class Procedure:
    #: "sub", "function", "get", "let" or "set".
    kind: str = "sub"
    name: str = ""
    params: list[Param] = field(default_factory=lambda: [])
    body: list[Stmt] = field(default_factory=lambda: [])
    returns: str = "Variant"
    scope: str = "public"
    static: bool = False
    line: int = 0
    attributes: dict[str, str] = field(default_factory=lambda: {})

    @property
    def is_property(self) -> bool:
        return self.kind in ("get", "let", "set")


@dataclass(slots=True)
class TypeField:
    name: str = ""
    declared: str = "Variant"
    bounds: list[tuple[Expr | None, Expr]] = field(default_factory=lambda: [])
    is_array: bool = False


@dataclass(slots=True)
class TypeDef:
    name: str = ""
    fields: list[TypeField] = field(default_factory=lambda: [])
    scope: str = "public"
    line: int = 0


@dataclass(slots=True)
class EnumDef:
    name: str = ""
    members: list[tuple[str, Expr | None]] = field(default_factory=lambda: [])
    scope: str = "public"
    line: int = 0


@dataclass(slots=True)
class DeclareDef:
    name: str = ""
    library: str = ""
    alias: str = ""
    is_function: bool = False
    params: list[Param] = field(default_factory=lambda: [])
    returns: str = "Variant"
    line: int = 0


@dataclass(slots=True)
class EventDef:
    name: str = ""
    params: list[Param] = field(default_factory=lambda: [])
    line: int = 0


@dataclass(slots=True)
class Module:
    """A parsed module: its declarations and its procedures."""

    name: str = ""
    kind: str = "standard"
    attributes: dict[str, str] = field(default_factory=lambda: {})
    options: list[str] = field(default_factory=lambda: [])
    variables: list[Dim] = field(default_factory=lambda: [])
    constants: list[ConstDecl] = field(default_factory=lambda: [])
    types: list[TypeDef] = field(default_factory=lambda: [])
    enums: list[EnumDef] = field(default_factory=lambda: [])
    declares: list[DeclareDef] = field(default_factory=lambda: [])
    events: list[EventDef] = field(default_factory=lambda: [])
    procedures: list[Procedure] = field(default_factory=lambda: [])
    implements: list[str] = field(default_factory=lambda: [])
    source: str = ""

    @property
    def option_explicit(self) -> bool:
        return any(option.lower().startswith("explicit") for option in self.options)

    @property
    def option_compare_text(self) -> bool:
        return any(option.lower().replace(" ", "") == "comparetext" for option in self.options)

    @property
    def option_base(self) -> int:
        for option in self.options:
            if option.lower().startswith("base"):
                return int(option.split()[1])
        return 0
