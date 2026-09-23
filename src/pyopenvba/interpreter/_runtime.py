"""The interpreter: a parsed module, executed against a host.

Execution is a walk of the tree with one frame per call.  Three things
are worth stating, because they are where a VBA interpreter usually
gets the semantics wrong:

* ``On Error GoTo`` does not unwind.  The handler runs where the error
  happened, so ``Resume Next`` can hand control back to the statement
  after it even when that statement is inside a loop, which is how VBA
  behaves and what a loop-and-skip macro depends on.
* ``ByRef`` is the default.  An argument that is a plain variable is
  passed as a slot the callee can write through; anything else, a
  literal, an expression, or a parenthesised variable, is a copy.
* An object used where a value is wanted gives up its default member,
  so ``Range("A1") + 1`` adds to the cell's value without anyone
  writing ``.Value``.
"""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field
from decimal import Decimal
from functools import lru_cache
from typing import Any, Callable, Final

from pyopenvba.exceptions import VBACompileError, VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._parse import parse_module
from pyopenvba.interpreter._values import (
    EMPTY,
    ERR_MEMBER_NOT_FOUND,
    ERR_OBJECT_VARIABLE_NOT_SET,
    ERR_SUBSCRIPT_OUT_OF_RANGE,
    ERR_TYPE_MISMATCH,
    MISSING,
    STORAGE_WIDTH,
    NOTHING,
    NULL,
    PRINT_CONTINUE,
    PRINT_ZONE,
    VBAArray,
    VBAInt,
    add,
    coerce,
    compare,
    concat,
    default_for,
    divide,
    error,
    int_divide,
    is_object,
    like_match,
    logical_and,
    logical_eqv,
    logical_imp,
    logical_not,
    logical_or,
    logical_xor,
    modulo,
    multiply,
    negate,
    power,
    subtract,
    to_bool,
    to_integer,
    to_text,
    type_name,
)

#: How deep a call chain may go before this reports VBA's error 28.
#: VBA's own limit is however much stack it has, so any number here is a
#: choice; this one is deep enough for ordinary recursion and shallow
#: enough to report the error rather than exhaust Python's stack.
MAX_DEPTH: Final = 160


# --- control-flow signals -----------------------------------------------------------


class _Signal(Exception):
    """Not an error: one of VBA's jumps, travelling up the Python stack."""


class _ExitSignal(_Signal):
    def __init__(self, what: str) -> None:
        super().__init__(what)
        self.what = what


class _GotoSignal(_Signal):
    def __init__(self, label: str) -> None:
        super().__init__(label)
        self.label = label


class _ResumeSignal(_Signal):
    def __init__(self, mode: str, label: str = "") -> None:
        super().__init__(mode)
        self.mode = mode
        self.label = label


class _EndSignal(_Signal):
    """The End statement: stop everything, without an error."""


# --- storage ------------------------------------------------------------------------


@dataclass(slots=True)
class Slot:
    """One variable.  Passing the slot itself is what ByRef means."""

    value: object = EMPTY
    declared: str = "Variant"

    def get(self) -> object:
        return self.value

    def set(self, value: object) -> None:
        self.value = coerce(value, self.declared)


@dataclass(slots=True)
class Frame:
    """One call in progress."""

    procedure: A.Procedure
    module: ModuleRuntime
    locals: dict[str, Slot] = field(default_factory=lambda: {})
    me: object = None
    with_stack: list[object] = field(default_factory=lambda: [])
    handler: str = "none"
    handler_label: str = ""
    in_handler: bool = False
    args_named: dict[str, Slot] = field(default_factory=lambda: {})

    def slot(self, name: str) -> Slot | None:
        return self.locals.get(name.lower())


class ModuleRuntime:
    """A parsed module together with the state it holds while running."""

    def __init__(self, parsed: A.Module, interpreter: Interpreter) -> None:
        self.parsed = parsed
        self.interpreter = interpreter
        self.name = parsed.name
        self.variables: dict[str, Slot] = {}
        self.constants: dict[str, object] = {}
        self.statics: dict[str, dict[str, Slot]] = {}
        self.procedures: dict[str, list[A.Procedure]] = {}
        self.types: dict[str, A.TypeDef] = {}
        self.enums: dict[str, dict[str, object]] = {}
        self.enum_members: dict[str, object] = {}
        self.declares: dict[str, A.DeclareDef] = {}
        #: A document module's one instance, bound to the object it is the code of (see bind_document).
        self.document: UserClassInstance | None = None
        self._initialised = False
        for procedure in parsed.procedures:
            self.procedures.setdefault(procedure.name.lower(), []).append(procedure)
        for defined in parsed.types:
            self.types[defined.name.lower()] = defined

    @property
    def is_class(self) -> bool:
        return self.parsed.kind in ("class", "document", "form")

    def initialise(self) -> None:
        """Give every module-level declaration its starting value."""
        if self._initialised:
            return
        self._initialised = True
        interpreter = self.interpreter
        for enum in self.parsed.enums:
            values: dict[str, object] = {}
            running = VBAInt(0, "Long")
            for name, expression in enum.members:
                if expression is not None:
                    running = VBAInt(int(to_integer(interpreter.evaluate_constant(expression, self), "Long")), "Long")
                values[name.lower()] = running
                self.enum_members[name.lower()] = running
                running = VBAInt(int(running) + 1, "Long")
            self.enums[enum.name.lower()] = values
        for declaration in self.parsed.constants:
            for name, declared, expression in declaration.names:
                self.constants[name.lower()] = coerce(
                    interpreter.evaluate_constant(expression, self), declared
                )
        for group in self.parsed.variables:
            for declaration in group.decls:
                self.variables[declaration.name.lower()] = interpreter.make_slot(declaration, self)
        for declare in self.parsed.declares:
            self.declares[declare.name.lower()] = declare

    def procedure(self, name: str, kind: str = "") -> A.Procedure | None:
        found = self.procedures.get(name.lower())
        if not found:
            return None
        if kind:
            for candidate in found:
                if candidate.kind == kind:
                    return candidate
            return None
        for candidate in found:
            if candidate.kind in ("sub", "function", "get"):
                return candidate
        return found[0]


# --- the objects the language itself provides ---------------------------------------


class ErrObject(VBAObject):
    """VBA's ``Err``: what the last error was, and how to raise one."""

    vba_type_name = "ErrObject"

    def __init__(self) -> None:
        self.number = 0
        self.description = ""
        self.source = ""
        self.help_file = ""
        self.help_context = 0

    @member
    def Number(self) -> object:
        return VBAInt(self.number, "Long")

    @setter("Number")
    def _set_number(self, value: object) -> None:
        self.number = int(to_integer(value, "Long"))
        if self.number == 0:
            self.description = ""
            self.source = ""

    @member
    def Description(self) -> object:
        return self.description

    @setter("Description")
    def _set_description(self, value: object) -> None:
        self.description = to_text(value)

    @member
    def Source(self) -> object:
        return self.source

    @setter("Source")
    def _set_source(self, value: object) -> None:
        self.source = to_text(value)

    @member
    def HelpFile(self) -> object:
        return self.help_file

    @member
    def HelpContext(self) -> object:
        return VBAInt(self.help_context, "Long")

    @method
    def Clear(self) -> object:
        self.reset()
        return EMPTY

    def reset(self) -> None:
        """No error, as VBA leaves Err on entering a procedure, on an On Error statement, on Resume and on leaving
        an error handler (tests/fixtures/vba_semantics/err_lifetime.json)."""
        self.number = 0
        self.description = ""
        self.source = ""

    @method
    def Raise(
        self,
        Number: object = MISSING,
        Source: object = MISSING,
        Description: object = MISSING,
        HelpFile: object = MISSING,
        HelpContext: object = MISSING,
    ) -> object:
        if Number is MISSING:
            raise error(449)
        number = int(to_integer(Number, "Long"))
        text = to_text(Description) if Description is not MISSING else ""
        raise error(number, text, source=to_text(Source) if Source is not MISSING else "")

    def capture(self, failure: VBARuntimeError) -> None:
        self.number = failure.number
        self.description = failure.description
        self.source = failure.source

    def describe(self, indent: str = "") -> str:
        if not self.number:
            return f"{indent}Err: no error"
        return f"{indent}Err {self.number}: {self.description}"


class DebugObject(VBAObject):
    """``Debug``: Print goes to the console this interpreter collects."""

    vba_type_name = "Debug"

    def __init__(self, interpreter: Interpreter) -> None:
        self.interpreter = interpreter

    @method
    def Print(self, *args: object) -> object:
        self.interpreter.write_print(_print_text(args), holding=bool(args) and args[-1] is PRINT_CONTINUE)
        return EMPTY

    @method
    def Assert(self, Condition: object = MISSING) -> object:
        if Condition is not MISSING and not to_bool(Condition):
            self.interpreter.assertions.append(self.interpreter.current_where())
        return EMPTY


#: How wide Print's comma zones are.
PRINT_ZONE_WIDTH: Final = 14


def _print_text(args: Sequence[object]) -> str:
    """Print's own formatting.

    A number is written with a leading space where it is not negative
    and a trailing one always, a comma moves to the next 14-column
    zone, and a trailing semicolon means the next Print continues this
    line.
    """
    out = ""
    for value in args:
        if value is PRINT_ZONE:
            out += " " * (PRINT_ZONE_WIDTH - len(out) % PRINT_ZONE_WIDTH)
            continue
        if value is PRINT_CONTINUE:
            continue
        if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
            out += to_text(value)
            continue
        text = to_text(value)
        out += (text if text.startswith("-") else f" {text}") + " "
    return out


class Collection(VBACollection):
    """VBA's intrinsic Collection, keys and all."""

    vba_type_name = "Collection"

    def __init__(self) -> None:
        self._items: list[object] = []
        self._keys: list[str | None] = []

    def vba_items(self) -> list[object]:
        return self._items

    @method
    def Add(
        self,
        Item: object = MISSING,
        Key: object = MISSING,
        Before: object = MISSING,
        After: object = MISSING,
    ) -> object:
        if Item is MISSING:
            raise error(449)
        key = to_text(Key) if Key is not MISSING else None
        if key is not None and any(existing is not None and existing.lower() == key.lower() for existing in self._keys):
            raise error(457)
        at = len(self._items)
        if Before is not MISSING:
            at = self._position(Before) - 1
        elif After is not MISSING:
            at = self._position(After)
        self._items.insert(at, Item)
        self._keys.insert(at, key)
        return EMPTY

    @method
    def Remove(self, Index: object = MISSING) -> object:
        at = self._position(Index)
        self._items.pop(at - 1)
        self._keys.pop(at - 1)
        return EMPTY

    def _position(self, index: object) -> int:
        if isinstance(index, str):
            for at, key in enumerate(self._keys, start=1):
                if key is not None and key.lower() == index.lower():
                    return at
            raise error(5)
        at = int(to_integer(index, "Long"))
        if not 1 <= at <= len(self._items):
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        return at

    def vba_lookup(self, index: object, items: list[object]) -> object:
        return self._items[self._position(index) - 1]

    @method(default=True)
    def Item(self, Index: object = MISSING) -> object:
        """A method in VBA's Collection, where a host's collections have a property: CallByName(c, "Item",
        VbMethod, 2) reaches it (tests/fixtures/excel_model/probes.txt)."""
        if Index is MISSING:
            raise error(449)
        return self.vba_lookup(Index, self._items)

    def describe(self, indent: str = "") -> str:
        return f"{indent}Collection ({len(self._items)} items)"


class UserTypeValue(VBAObject):
    """An instance of a ``Type ... End Type``."""

    def __init__(self, defined: A.TypeDef, interpreter: Interpreter, module: ModuleRuntime) -> None:
        self.defined = defined
        self.vba_type_name = defined.name
        self.fields: dict[str, Slot] = {}
        for field_ in defined.fields:
            declaration = A.VarDecl(
                name=field_.name,
                declared=field_.declared,
                bounds=field_.bounds,
                is_array=field_.is_array,
            )
            self.fields[field_.name.lower()] = interpreter.make_slot(declaration, module)

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        slot = self.fields.get(name.lower())
        if slot is None:
            raise error(ERR_MEMBER_NOT_FOUND, f"{self.vba_type_name} has no member named {name}")
        value = slot.get()
        if args and isinstance(value, VBAArray):
            return value.get([int(to_integer(one, "Long")) for one in args])
        return value

    def vba_set(
        self,
        name: str,
        value: object,
        args: Sequence[object] = (),
        named: dict[str, object] | None = None,
        *,
        by_ref: bool = False,
    ) -> None:
        slot = self.fields.get(name.lower())
        if slot is None:
            raise error(ERR_MEMBER_NOT_FOUND, f"{self.vba_type_name} has no member named {name}")
        if args and isinstance(slot.value, VBAArray):
            slot.value.set([int(to_integer(one, "Long")) for one in args], value)
            return
        slot.set(value)

    def vba_member(self, name: str) -> Any:
        return None

    def describe(self, indent: str = "") -> str:
        inner = ", ".join(f"{name}={_short(slot.get())}" for name, slot in self.fields.items())
        return f"{indent}{self.vba_type_name}({inner})"


class UserClassInstance(VBAObject):
    """An instance of a class module."""

    def __init__(self, module: ModuleRuntime, interpreter: Interpreter) -> None:
        self.module = module
        self.interpreter = interpreter
        self.vba_type_name = module.name
        #: For a document module, the object it is the code of -- a worksheet, the workbook -- which Me is.
        self.host: VBAObject | None = None
        self.variables: dict[str, Slot] = {}
        for group in module.parsed.variables:
            for declaration in group.decls:
                self.variables[declaration.name.lower()] = interpreter.make_slot(declaration, module)

    def vba_member(self, name: str) -> Any:
        return None

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        procedure = self.module.procedure(name, "get") or self.module.procedure(name)
        if procedure is not None and procedure.kind in ("get", "function", "sub"):
            return self.interpreter.call(procedure, self.module, list(args), named or {}, me=self)
        slot = self.variables.get(name.lower())
        if slot is not None:
            value = slot.get()
            if args and isinstance(value, VBAArray):
                return value.get([int(to_integer(one, "Long")) for one in args])
            return value
        constant = self.module.constants.get(name.lower())
        if constant is not None:
            return constant
        raise error(ERR_MEMBER_NOT_FOUND, f"{self.vba_type_name} has no member named {name}")

    def vba_set(
        self,
        name: str,
        value: object,
        args: Sequence[object] = (),
        named: dict[str, object] | None = None,
        *,
        by_ref: bool = False,
    ) -> None:
        kind = "set" if by_ref else "let"
        procedure = self.module.procedure(name, kind)
        if procedure is None and by_ref:
            procedure = self.module.procedure(name, "let")
        if procedure is not None:
            self.interpreter.call(procedure, self.module, [*args, value], named or {}, me=self)
            return
        slot = self.variables.get(name.lower())
        if slot is None:
            raise error(ERR_MEMBER_NOT_FOUND, f"{self.vba_type_name} has no member named {name}")
        if args and isinstance(slot.value, VBAArray):
            slot.value.set([int(to_integer(one, "Long")) for one in args], value)
            return
        slot.set(value)

    def vba_value(self) -> object:
        procedure = self.module.procedure("Value", "get")
        if procedure is not None:
            return self.interpreter.call(procedure, self.module, [], {}, me=self)
        raise error(ERR_MEMBER_NOT_FOUND, f"{self.vba_type_name} has no default member")

    # --- a document module: a sheet's or the workbook's own code --------------------------

    def public(self, name: str) -> bool:
        """Whether code outside the module reaches ``name``: a Public procedure or module-level variable."""
        key = name.lower()
        if any(procedure.scope != "private" for procedure in self.module.procedures.get(key, [])):
            return True
        return key in self.variables and any(
            group.scope in ("public", "global") and any(declaration.name.lower() == key for declaration in group.decls)
            for group in self.module.parsed.variables)

    def raise_event(self, name: str, args: list[object]) -> None:
        self.interpreter.raise_event(self, name, args)

    def describe(self, indent: str = "") -> str:
        inner = ", ".join(f"{name}={_short(slot.get())}" for name, slot in self.variables.items())
        return f"{indent}{self.vba_type_name}({inner})"


def _short(value: object) -> str:
    if isinstance(value, str):
        return f'"{value}"'
    if isinstance(value, VBAObject):
        return f"<{value.vba_type_name}>"
    if value is EMPTY:
        return "Empty"
    try:
        return to_text(value)
    except VBARuntimeError:
        return type_name(value)


# --- assignment targets --------------------------------------------------------------


@dataclass(slots=True)
class Assignable:
    """Somewhere a value can be put: a slot, an element, or a member."""

    put: Callable[[object, bool], None]
    take: Callable[[], object]


# --- the interpreter -------------------------------------------------------------------


class Interpreter:
    """Runs VBA against a host object model."""

    def __init__(self, host: HostBridge | None = None) -> None:
        self.modules: dict[str, ModuleRuntime] = {}
        self.host = host if host is not None else HostBridge()
        self.err = ErrObject()
        self.debug = DebugObject(self)
        self.console: list[str] = []
        self.assertions: list[str] = []
        self.frames: list[Frame] = []
        self.option_compare_text = False
        #: Every MsgBox and InputBox the run put up, in order.
        self.dialogs: list[tuple[str, str, str, int]] = []
        #: What the next dialog answers with, oldest first.
        self.answers: list[object] = []
        self.frozen_now: Any = None
        self._rnd_state = 0x50000
        self._line_open = False
        self._intrinsics = _load_intrinsics()

    # --- loading ---------------------------------------------------------------------

    def add_module(self, source: str, *, name: str = "", kind: str = "standard") -> ModuleRuntime:
        """Parse ``source`` and add it to the project."""
        parsed = parse_module(source, name=name, kind=kind)
        if not parsed.name:
            parsed.name = name or f"Module{len(self.modules) + 1}"
        runtime = ModuleRuntime(parsed, self)
        self.modules[parsed.name.lower()] = runtime
        return runtime

    def module(self, name: str) -> ModuleRuntime:
        found = self.modules.get(name.lower())
        if found is None:
            raise VBACompileError(f"there is no module named {name}")
        return found

    def bind_document(self, name: str, host: VBAObject) -> UserClassInstance:
        """Make document module ``name`` the code of ``host``: Me is ``host``, its members are the module's by
        name, and ``host`` answers the module's Public members to code outside it."""
        runtime = self.module(name)
        instance = UserClassInstance(runtime, self)
        instance.host = host
        runtime.document = instance
        host.vba_document = instance
        return instance

    def initialise(self) -> None:
        for runtime in list(self.modules.values()):
            if not runtime.is_class:
                runtime.initialise()

    # --- console and the world outside the model ---------------------------------------

    def write_line(self, text: str) -> None:
        self.console.append(text)

    def write_print(self, text: str, *, holding: bool = False) -> None:
        """Print's output, which a trailing semicolon holds on one line."""
        if self._line_open and self.console:
            self.console[-1] += text
        else:
            self.console.append(text)
        self._line_open = holding

    def clock(self) -> Any:
        """Now, as the run sees it.  Fixed by ``freeze`` for a repeatable run."""
        import datetime as _dt

        return self.frozen_now if self.frozen_now is not None else _dt.datetime.now()

    def freeze(self, when: Any) -> None:
        """Pin Now, Date, Time and Timer, so a run can be compared with another."""
        self.frozen_now = when

    def show_message(self, prompt: str, style: int, title: str) -> object:
        """MsgBox without a screen: record it, answer from the queue."""
        self.dialogs.append(("MsgBox", prompt, title, style))
        if self.answers:
            return self.answers.pop(0)
        return VBAInt(_default_button(style), "Integer")

    def ask_for_input(self, prompt: str, title: str, default: str) -> object:
        self.dialogs.append(("InputBox", prompt, title, 0))
        if self.answers:
            return to_text(self.answers.pop(0))
        return default

    def create_com(self, class_name: str) -> object:
        if class_name.lower() == "scripting.dictionary":
            from pyopenvba.interpreter._scripting import Dictionary

            return Dictionary()
        made = self.host.create(class_name)
        if made is not UNRESOLVED:
            return made
        raise VBAUnsupportedError(
            f"CreateObject({class_name!r}) needs a real COM server, which pyOpenVBA does not start"
        )

    def current_where(self) -> str:
        if not self.frames:
            return ""
        frame = self.frames[-1]
        return f"{frame.module.name}.{frame.procedure.name}"

    # --- running ---------------------------------------------------------------------

    def run(self, procedure_name: str, args: Sequence[object] = (), *, module: str = "") -> object:
        """Run one procedure by name, as Application.Run would."""
        self.initialise()
        target, runtime = self.find_procedure(procedure_name, module)
        if target is None or runtime is None:
            raise VBARuntimeError(
                35, f"Sub or Function not defined: {procedure_name}", where=module or "the project"
            )
        try:
            # A sheet's or the workbook's own module runs as that object's code.
            return self.call(target, runtime, list(args), {}, me=runtime.document)
        except _EndSignal:
            return EMPTY

    def find_procedure(self, name: str, module: str = "") -> tuple[A.Procedure | None, ModuleRuntime | None]:
        if "." in name and not module:
            module, _, name = name.rpartition(".")
        if module:
            runtime = self.module(module)
            return runtime.procedure(name), runtime
        for runtime in self.modules.values():
            if runtime.is_class:
                continue
            found = runtime.procedure(name)
            if found is not None:
                return found, runtime
        return None, None

    def call(
        self,
        procedure: A.Procedure,
        module: ModuleRuntime,
        args: list[object],
        named: dict[str, object],
        *,
        me: object = None,
    ) -> object:
        """Call one procedure and give back what it returns."""
        if len(self.frames) >= MAX_DEPTH:
            raise error(28)
        module.initialise()
        frame = Frame(procedure=procedure, module=module, me=me)
        self._bind_arguments(frame, procedure, args, named)
        # A procedure starts with no error, whatever its caller's Err held; what it leaves in Err stays after it.
        self.err.reset()
        if procedure.kind in ("function", "get"):
            frame.locals[procedure.name.lower()] = Slot(
                default_for(procedure.returns), procedure.returns
            )
        self.frames.append(frame)
        try:
            self._run_body(frame)
        except _ExitSignal as signal:
            if signal.what not in ("sub", "function", "property"):
                raise
        except RecursionError:
            # Python ran out of stack before the VBA depth cap did.
            # Either way the macro recursed away, which is error 28.
            raise error(28) from None
        finally:
            self.frames.pop()
        if procedure.kind in ("function", "get"):
            return frame.locals[procedure.name.lower()].get()
        return EMPTY

    def _bind_arguments(
        self, frame: Frame, procedure: A.Procedure, args: list[object], named: dict[str, object]
    ) -> None:
        laid: list[object] = list(args)
        params = procedure.params
        if len(laid) > len(params) and not (params and params[-1].param_array):
            raise error(450, f"{procedure.name} takes {len(params)} arguments")
        for key, value in named.items():
            for index, parameter in enumerate(params):
                if parameter.name.lower() == key.lower():
                    while len(laid) <= index:
                        laid.append(MISSING)
                    laid[index] = value
                    break
            else:
                raise error(448, f"{procedure.name} has no argument named {key}")
        for index, parameter in enumerate(params):
            if parameter.param_array:
                rest = [self._by_value(one) for one in laid[index:]]
                array = VBAArray(
                    [(0, len(rest) - 1)] if rest else [(0, -1)],
                    items=rest if rest else [],
                )
                frame.locals[parameter.name.lower()] = Slot(array, "Variant")
                return
            given = laid[index] if index < len(laid) else MISSING
            if given is MISSING or given is None:
                if not parameter.optional and parameter.default is None:
                    if index < len(laid):
                        raise error(449, f"{procedure.name} needs an argument for {parameter.name}")
                    raise error(449, f"{procedure.name} needs an argument for {parameter.name}")
                if parameter.default is not None:
                    frame.locals[parameter.name.lower()] = Slot(
                        coerce(self.evaluate_constant(parameter.default, frame.module), parameter.declared),
                        parameter.declared,
                    )
                else:
                    frame.locals[parameter.name.lower()] = Slot(
                        MISSING if parameter.declared == "Variant" else default_for(parameter.declared),
                        "Variant" if parameter.declared == "Variant" else parameter.declared,
                    )
                continue
            if isinstance(given, Slot) and not parameter.by_val:
                frame.locals[parameter.name.lower()] = given
                continue
            value = self._by_value(given)
            frame.locals[parameter.name.lower()] = Slot(
                coerce(value, parameter.declared), parameter.declared
            )

    @staticmethod
    def _by_value(given: object) -> object:
        return given.get() if isinstance(given, Slot) else given

    def _run_body(self, frame: Frame) -> None:
        body = frame.procedure.body
        self._declare_locals(frame, body)
        at = 0
        while at < len(body):
            try:
                self._execute_block(body, frame, start=at)
                return
            except _GotoSignal as jump:
                index = _label_index(body, jump.label)
                if index is None:
                    raise self._unknown_label(jump.label, frame) from None
                at = index

    def _unknown_label(self, label: str, frame: Frame) -> Exception:
        return VBACompileError(
            f"label {label} is not in {frame.procedure.name}",
            where=f"{frame.module.name}.{frame.procedure.name}",
        )

    def _declare_locals(self, frame: Frame, body: Sequence[A.Stmt]) -> None:
        """Give every Dim in the procedure its slot before anything runs.

        VBA allocates a procedure's locals on entry, which is why a Dim
        inside an If still exists on the path that skipped it.
        """
        for statement in body:
            if isinstance(statement, A.Dim):
                for declaration in statement.decls:
                    key = declaration.name.lower()
                    if statement.static:
                        store = frame.module.statics.setdefault(frame.procedure.name.lower(), {})
                        if key not in store:
                            store[key] = self.make_slot(declaration, frame.module)
                        frame.locals[key] = store[key]
                    elif key not in frame.locals:
                        frame.locals[key] = self.make_slot(declaration, frame.module)
            elif isinstance(statement, A.ConstDecl):
                continue
            for nested in _nested_bodies(statement):
                self._declare_locals(frame, nested)

    # --- statements --------------------------------------------------------------------

    def _execute_block(self, body: Sequence[A.Stmt], frame: Frame, *, start: int = 0) -> None:
        at = start
        while at < len(body):
            statement = body[at]
            try:
                self._execute(statement, frame)
            except VBARuntimeError as failure:
                outcome = self._handle(failure, frame, statement)
                if outcome == "next":
                    at += 1
                    continue
                if outcome == "same":
                    continue
                raise
            except _GotoSignal as jump:
                index = _label_index(body, jump.label)
                if index is None:
                    raise
                at = index
                continue
            at += 1

    def _handle(self, failure: VBARuntimeError, frame: Frame, statement: A.Stmt) -> str:
        """What to do with a run-time error: resume, run a handler, or pass it on."""
        if not failure.where:
            failure.where = f"{frame.module.name}.{frame.procedure.name} line {statement.line}"
            failure.args = (f"{failure.where}: {failure.message}",)
        if frame.in_handler or frame.handler == "none":
            return "raise"
        self.err.capture(failure)
        if frame.handler == "resume_next":
            return "next"
        index = _label_index(frame.procedure.body, frame.handler_label)
        if index is None:
            raise self._unknown_label(frame.handler_label, frame) from failure
        frame.in_handler = True
        try:
            self._execute_block(frame.procedure.body, frame, start=index)
        except _ResumeSignal as resume:
            frame.in_handler = False
            self.err.reset()
            if resume.mode == "next":
                return "next"
            if resume.mode == "same":
                return "same"
            raise _GotoSignal(resume.label) from None
        except _GotoSignal:
            frame.in_handler = False
            raise
        except _ExitSignal:
            # Leaving the procedure from its handler leaves no error behind.
            self.err.reset()
            raise
        finally:
            if frame.in_handler:
                frame.in_handler = False
        self.err.reset()
        raise _ExitSignal(_exit_word(frame.procedure.kind))

    @classmethod
    def statement_handlers(cls) -> dict[type, Callable[..., None]]:
        """Which method runs each kind of statement."""
        table: dict[type, Callable[..., None]] = {
            A.Assign: cls._do_assign,
            A.CallStmt: cls._do_call,
            A.Dim: cls._do_dim,
            A.ConstDecl: cls._do_const,
            A.ReDim: cls._do_redim,
            A.Erase: cls._do_erase,
            A.If: cls._do_if,
            A.For: cls._do_for,
            A.ForEach: cls._do_for_each,
            A.DoLoop: cls._do_do,
            A.WhileLoop: cls._do_while,
            A.With: cls._do_with,
            A.SelectCase: cls._do_select,
            A.ExitStmt: cls._do_exit,
            A.GoTo: cls._do_goto,
            A.OnGoto: cls._do_on_goto,
            A.ReturnStmt: cls._do_return,
            A.Label: cls._do_label,
            A.OnError: cls._do_on_error,
            A.Resume: cls._do_resume,
            A.ErrorStmt: cls._do_error,
            A.Stop: cls._do_stop,
            A.EndStmt: cls._do_end,
            A.OptionStmt: cls._do_option,
            A.MidAssign: cls._do_mid_assign,
            A.RaiseEvent: cls._do_raise_event,
            A.Unsupported: cls._do_unsupported,
        }
        return table

    def _execute(self, statement: A.Stmt, frame: Frame) -> None:
        handler = _STATEMENTS.get(type(statement))
        if handler is None:  # pragma: no cover - every node has one
            raise VBAUnsupportedError(f"{type(statement).__name__} is not implemented")
        handler(self, statement, frame)

    def _do_assign(self, statement: A.Assign, frame: Frame) -> None:
        value = self.evaluate(statement.value, frame, want_object=statement.kind == "set")
        if statement.kind == "set" and not is_object(value) and value is not NOTHING:
            raise error(ERR_MEMBER_NOT_FOUND, "Set needs an object on the right")
        if statement.kind == "let" and isinstance(value, VBAObject):
            value = value.vba_value()
        target = self.resolve_target(statement.target, frame)
        target.put(value, statement.kind == "set")

    def _do_call(self, statement: A.CallStmt, frame: Frame) -> None:
        self.evaluate_call(statement.callee, statement.args, frame, statement_context=True)

    def _do_dim(self, statement: A.Dim, frame: Frame) -> None:
        for declaration in statement.decls:
            key = declaration.name.lower()
            if key not in frame.locals:
                frame.locals[key] = self.make_slot(declaration, frame.module)

    def _do_const(self, statement: A.ConstDecl, frame: Frame) -> None:
        for name, declared, expression in statement.names:
            frame.locals[name.lower()] = Slot(
                coerce(self.evaluate(expression, frame), declared), declared
            )

    def _do_redim(self, statement: A.ReDim, frame: Frame) -> None:
        for declaration in statement.decls:
            bounds = self._bounds(declaration, frame)
            target = self.resolve_target(A.Name(line=statement.line, name=declaration.name), frame)
            current = target.take()
            element = declaration.declared if declaration.declared != "Variant" else "Variant"
            if statement.preserve and isinstance(current, VBAArray):
                target.put(current.resized(bounds, preserve=True), False)
            else:
                if statement.preserve and not isinstance(current, VBAArray):
                    raise error(ERR_TYPE_MISMATCH, "ReDim Preserve needs an array")
                keep = current.element_type if isinstance(current, VBAArray) else element
                target.put(VBAArray(bounds, element_type=keep), False)

    def _do_erase(self, statement: A.Erase, frame: Frame) -> None:
        for name in statement.names:
            target = self.resolve_target(name, frame)
            current = target.take()
            if isinstance(current, VBAArray):
                if current.fixed:
                    fresh = VBAArray(current.bounds, element_type=current.element_type, fixed=True)
                    target.put(fresh, False)
                else:
                    target.put(VBAArray([(0, -1)], element_type=current.element_type), False)

    def _do_if(self, statement: A.If, frame: Frame) -> None:
        for branch in statement.branches:
            if to_bool(self.evaluate(branch.condition, frame)):
                self._execute_block(branch.body, frame)
                return
        if statement.otherwise:
            self._execute_block(statement.otherwise, frame)

    def _do_for(self, statement: A.For, frame: Frame) -> None:
        target = self.resolve_target(statement.target, frame)
        start = self.evaluate(statement.start, frame)
        limit = self.evaluate(statement.limit, frame)
        step = self.evaluate(statement.step, frame) if statement.step is not None else VBAInt(1, "Integer")
        counter = start
        target.put(counter, False)
        going_up = not to_bool(compare("<", step, VBAInt(0, "Integer")))
        while True:
            current = target.take()
            if going_up:
                if to_bool(compare(">", current, limit)):
                    return
            elif to_bool(compare("<", current, limit)):
                return
            try:
                self._execute_block(statement.body, frame)
            except _ExitSignal as signal:
                if signal.what == "for":
                    return
                raise
            target.put(add(target.take(), step), False)

    def _do_for_each(self, statement: A.ForEach, frame: Frame) -> None:
        target = self.resolve_target(statement.target, frame)
        subject = self.evaluate(statement.collection, frame, want_object=True)
        for item in _iterate(subject):
            target.put(item, isinstance(item, VBAObject))
            try:
                self._execute_block(statement.body, frame)
            except _ExitSignal as signal:
                if signal.what == "for":
                    return
                raise

    def _do_do(self, statement: A.DoLoop, frame: Frame) -> None:
        while True:
            if not statement.at_end and statement.test:
                truth = to_bool(self.evaluate(statement.condition, frame))
                if statement.test == "while" and not truth:
                    return
                if statement.test == "until" and truth:
                    return
            try:
                self._execute_block(statement.body, frame)
            except _ExitSignal as signal:
                if signal.what == "do":
                    return
                raise
            if statement.at_end and statement.test:
                truth = to_bool(self.evaluate(statement.condition, frame))
                if statement.test == "while" and not truth:
                    return
                if statement.test == "until" and truth:
                    return

    def _do_while(self, statement: A.WhileLoop, frame: Frame) -> None:
        while to_bool(self.evaluate(statement.condition, frame)):
            try:
                self._execute_block(statement.body, frame)
            except _ExitSignal as signal:
                if signal.what == "do":
                    return
                raise

    def _do_with(self, statement: A.With, frame: Frame) -> None:
        subject = self.evaluate(statement.subject, frame, want_object=True)
        frame.with_stack.append(subject)
        try:
            self._execute_block(statement.body, frame)
        finally:
            frame.with_stack.pop()

    def _do_select(self, statement: A.SelectCase, frame: Frame) -> None:
        subject = self.evaluate(statement.subject, frame)
        for block in statement.blocks:
            for clause in block.clauses:
                if self._case_matches(clause, subject, frame):
                    self._execute_block(block.body, frame)
                    return
        if statement.otherwise is not None:
            self._execute_block(statement.otherwise, frame)

    def _case_matches(self, clause: A.CaseClause, subject: object, frame: Frame) -> bool:
        text = self.option_compare_text
        if clause.kind == "value":
            return to_bool(compare("=", subject, self.evaluate(clause.value, frame), text_compare=text))
        if clause.kind == "range":
            low = self.evaluate(clause.value, frame)
            high = self.evaluate(clause.upper, frame)
            return to_bool(compare(">=", subject, low, text_compare=text)) and to_bool(
                compare("<=", subject, high, text_compare=text)
            )
        return to_bool(compare(clause.op, subject, self.evaluate(clause.value, frame), text_compare=text))

    def _do_exit(self, statement: A.ExitStmt, frame: Frame) -> None:
        raise _ExitSignal(statement.what)

    def _do_goto(self, statement: A.GoTo, frame: Frame) -> None:
        if statement.gosub:
            raise VBAUnsupportedError("GoSub is not implemented by pyOpenVBA")
        raise _GotoSignal(statement.label)

    def _do_on_goto(self, statement: A.OnGoto, frame: Frame) -> None:
        index = int(to_integer(self.evaluate(statement.value, frame), "Long"))
        if statement.gosub:
            raise VBAUnsupportedError("On ... GoSub is not implemented by pyOpenVBA")
        if 1 <= index <= len(statement.labels):
            raise _GotoSignal(statement.labels[index - 1])

    def _do_return(self, statement: A.ReturnStmt, frame: Frame) -> None:
        raise VBAUnsupportedError("Return, the GoSub kind, is not implemented by pyOpenVBA")

    def _do_label(self, statement: A.Label, frame: Frame) -> None:
        return None

    def _do_on_error(self, statement: A.OnError, frame: Frame) -> None:
        # Every On Error statement clears Err (tests/fixtures/vba_semantics/err_lifetime.json).
        self.err.reset()
        if statement.mode == "clear":
            frame.handler = "none"
            frame.handler_label = ""
        elif statement.mode == "resume_next":
            frame.handler = "resume_next"
        else:
            frame.handler = "goto"
            frame.handler_label = statement.label

    def _do_resume(self, statement: A.Resume, frame: Frame) -> None:
        if not frame.in_handler:
            raise error(20)
        raise _ResumeSignal(statement.mode, statement.label)

    def _do_error(self, statement: A.ErrorStmt, frame: Frame) -> None:
        raise error(int(to_integer(self.evaluate(statement.number, frame), "Long")))

    def _do_stop(self, statement: A.Stop, frame: Frame) -> None:
        self.write_line(f"[Stop at {self.current_where()}]")

    def _do_end(self, statement: A.EndStmt, frame: Frame) -> None:
        raise _EndSignal()

    def _do_option(self, statement: A.OptionStmt, frame: Frame) -> None:
        return None

    def _do_mid_assign(self, statement: A.MidAssign, frame: Frame) -> None:
        target = self.resolve_target(statement.target, frame)
        text = to_text(target.take())
        start = int(to_integer(self.evaluate(statement.start, frame), "Long"))
        if start < 1:
            raise error(5)
        replacement = to_text(self.evaluate(statement.value, frame))
        room = len(text) - start + 1
        width = room if statement.length is None else int(to_integer(self.evaluate(statement.length, frame), "Long"))
        width = max(0, min(width, room, len(replacement)))
        target.put(text[: start - 1] + replacement[:width] + text[start - 1 + width :], False)

    def _do_raise_event(self, statement: A.RaiseEvent, frame: Frame) -> None:
        args = [self.evaluate(argument.value, frame) for argument in statement.args if argument.value is not None]
        if isinstance(frame.me, UserClassInstance):
            self.raise_event(frame.me, statement.name, args)
            return
        raise VBAUnsupportedError("RaiseEvent outside a class module is not implemented by pyOpenVBA")

    def raise_event(self, source: object, name: str, args: list[object]) -> None:
        """Events are recognised and not delivered: nothing is listening."""
        raise VBAUnsupportedError(
            f"RaiseEvent {name}: pyOpenVBA does not connect WithEvents sinks"
        )

    def _do_unsupported(self, statement: A.Unsupported, frame: Frame) -> None:
        raise VBAUnsupportedError(
            f"{statement.text.split(' ')[0]} is {statement.reason}, which pyOpenVBA does not implement",
            where=f"{frame.module.name}.{frame.procedure.name} line {statement.line}",
        )

    # --- expressions --------------------------------------------------------------------

    def evaluate(self, expression: A.Expr | None, frame: Frame | None, *, want_object: bool = False) -> object:
        if expression is None:
            return MISSING
        value = self._evaluate(expression, frame)
        if isinstance(value, Slot):
            value = value.get()
        if not want_object and isinstance(value, VBAObject) and not isinstance(value, (Collection,)):
            return value
        return value

    def evaluate_constant(self, expression: A.Expr, module: ModuleRuntime) -> object:
        """An expression in a declaration, where no frame exists yet."""
        frame = Frame(procedure=A.Procedure(name="<declarations>"), module=module)
        return self.evaluate(expression, frame)

    def _evaluate(self, expression: A.Expr, frame: Frame | None) -> object:
        kind = type(expression)
        if kind is A.Literal:
            return expression.value  # type: ignore[attr-defined]
        if kind is A.Name:
            return self._read_name(expression.name, frame, expression)  # type: ignore[attr-defined]
        if kind is A.Member:
            return self._read_member(expression, frame)  # type: ignore[arg-type]
        if kind is A.Index:
            return self.evaluate_call(expression.target, expression.args, frame)  # type: ignore[attr-defined]
        if kind is A.Binary:
            return self._binary(expression, frame)  # type: ignore[arg-type]
        if kind is A.Unary:
            return self._unary(expression, frame)  # type: ignore[arg-type]
        if kind is A.Bang:
            return self._read_bang(expression, frame)  # type: ignore[arg-type]
        if kind is A.NewExpr:
            return self.create(expression.type_name)  # type: ignore[attr-defined]
        if kind is A.TypeOfIs:
            value = self.evaluate(expression.value, frame, want_object=True)  # type: ignore[attr-defined]
            return _is_type(value, expression.type_name)  # type: ignore[attr-defined]
        if kind is A.Bracket:
            return self.host.evaluate_bracket(expression.text)  # type: ignore[attr-defined]
        if kind is A.AddressOf:
            raise VBAUnsupportedError("AddressOf is not implemented by pyOpenVBA")
        raise VBAUnsupportedError(f"{kind.__name__} is not implemented by pyOpenVBA")

    def _binary(self, expression: A.Binary, frame: Frame | None) -> object:
        op = expression.op
        if op == "and":
            left = self.evaluate(expression.left, frame)
            return logical_and(left, self.evaluate(expression.right, frame))
        if op == "or":
            left = self.evaluate(expression.left, frame)
            return logical_or(left, self.evaluate(expression.right, frame))
        left = self.evaluate(expression.left, frame, want_object=op == "is")
        right = self.evaluate(expression.right, frame, want_object=op == "is")
        if op != "is":
            left = _scalar(left)
            right = _scalar(right)
        if op == "+":
            return add(left, right)
        if op == "-":
            return subtract(left, right)
        if op == "*":
            return multiply(left, right)
        if op == "/":
            return divide(left, right)
        if op == "\\":
            return int_divide(left, right)
        if op == "mod":
            return modulo(left, right)
        if op == "^":
            return power(left, right)
        if op == "&":
            return concat(left, right)
        if op == "like":
            if left is NULL or right is NULL:
                return NULL
            return like_match(to_text(left), to_text(right), text_compare=self.option_compare_text)
        if op == "xor":
            return logical_xor(left, right)
        if op == "eqv":
            return logical_eqv(left, right)
        if op == "imp":
            return logical_imp(left, right)
        return compare(op, left, right, text_compare=self.option_compare_text)

    def _unary(self, expression: A.Unary, frame: Frame | None) -> object:
        value = self.evaluate(expression.operand, frame)
        if expression.op == "-":
            return negate(_scalar(value))
        return logical_not(_scalar(value))

    def _read_name(self, name: str, frame: Frame | None, expression: A.Expr) -> object:
        found = self.lookup(name, frame)
        if found is UNRESOLVED:
            raise self._undefined(name, frame, expression)
        return found

    def _undefined(self, name: str, frame: Frame | None, expression: A.Expr) -> Exception:
        where = ""
        if frame is not None:
            where = f"{frame.module.name}.{frame.procedure.name} line {expression.line}"
        return VBACompileError(f"Variable not defined: {name}", where=where)

    def lookup(self, name: str, frame: Frame | None, *, want_slot: bool = False) -> object:
        """Resolve a bare name the way VBA resolves one."""
        key = _bare(name)
        if frame is not None:
            if key == "me" and frame.me is not None:
                # Me in a sheet's or the workbook's own module is the sheet or the workbook.
                return frame.me.host if isinstance(frame.me, UserClassInstance) and frame.me.host is not None \
                    else frame.me
            slot = frame.locals.get(key)
            if slot is not None:
                return slot if want_slot else slot.get()
            if frame.me is not None and isinstance(frame.me, UserClassInstance):
                instance = frame.me
                if key in instance.variables:
                    return instance.variables[key] if want_slot else instance.variables[key].get()
                if instance.module.procedure(key) is not None and key != frame.procedure.name.lower():
                    return self.call_named(key, [], {}, instance.module, me=instance)
            module = frame.module
            module.initialise()
            if key in module.constants:
                return module.constants[key]
            if key in module.variables:
                return module.variables[key] if want_slot else module.variables[key].get()
            if key in module.enum_members:
                return module.enum_members[key]
            procedure = module.procedure(key)
            if procedure is not None:
                return self.call(procedure, module, [], {}, me=frame.me if module.is_class else None)
            host = _host_of(frame)
            if host is not None and host.vba_member(key) is not None:
                # A sheet's own module reaches the sheet's members by name: Range is Me.Range.
                return _dispatch_get(host, key, [], {})
        for runtime in self.modules.values():
            if runtime.is_class:
                continue
            runtime.initialise()
            if key in runtime.constants:
                return runtime.constants[key]
            if key in runtime.variables:
                return runtime.variables[key] if want_slot else runtime.variables[key].get()
            if key in runtime.enum_members:
                return runtime.enum_members[key]
            procedure = runtime.procedure(key)
            if procedure is not None and procedure.scope != "private":
                return self.call(procedure, runtime, [], {})
        if key == "err":
            return self.err
        if key == "debug":
            return self.debug
        constant = self.host.constant(key)
        if constant is not UNRESOLVED:
            return constant
        language = vba_constant(key)
        if language is not UNRESOLVED:
            return language
        supplied = self.host.global_object(key)
        if supplied is not UNRESOLVED:
            return supplied
        if self.host.has_global_member(key):
            return self.host.call_global(key, [], {})
        if key in self._intrinsics:
            return self.call_intrinsic(key, [], {})
        return UNRESOLVED

    def _read_member(self, expression: A.Member, frame: Frame | None) -> object:
        target = self._member_target(expression, frame)
        return _dispatch_get(target, expression.name, [], {})

    def _member_target(self, expression: A.Member | A.Bang, frame: Frame | None) -> object:
        if expression.target is None:
            if frame is None or not frame.with_stack:
                raise error(91, "a leading dot needs an open With block")
            return frame.with_stack[-1]
        value = self.evaluate(expression.target, frame, want_object=True)
        if value is NOTHING:
            raise error(ERR_OBJECT_VARIABLE_NOT_SET)
        return value

    def _read_bang(self, expression: A.Bang, frame: Frame | None) -> object:
        target = self._member_target(expression, frame)
        if isinstance(target, VBAObject):
            spec = target.vba_default_member()
            if spec is not None and spec.getter is not None:
                return target.vba_get(spec.name, [expression.name])
        raise error(ERR_MEMBER_NOT_FOUND, f"{type_name(target)} has no default member for !")

    # --- calls ---------------------------------------------------------------------------

    def evaluate_call(
        self,
        callee: A.Expr | None,
        args: list[A.Argument],
        frame: Frame | None,
        *,
        statement_context: bool = False,
    ) -> object:
        """Everything that looks like ``thing(args)``, including no args at all."""
        if callee is None:
            raise VBACompileError("a call with nothing to call")
        if isinstance(callee, A.Name):
            return self._call_name(callee, args, frame, statement_context=statement_context)
        if isinstance(callee, A.Member):
            target = self._member_target(callee, frame)
            positional, named = self._arguments(args, frame, target=target, name=callee.name)
            return _dispatch_get(target, callee.name, positional, named)
        if isinstance(callee, A.Index):
            inner = self.evaluate_call(callee.target, callee.args, frame)
            positional, named = self._arguments(args, frame)
            return _call_value(inner, positional, named)
        value = self.evaluate(callee, frame, want_object=True)
        positional, named = self._arguments(args, frame)
        return _call_value(value, positional, named)

    def _call_name(
        self,
        callee: A.Name,
        args: list[A.Argument],
        frame: Frame | None,
        *,
        statement_context: bool = False,
    ) -> object:
        key = _bare(callee.name)
        if frame is not None:
            if args and key == frame.procedure.name.lower() and frame.procedure.kind in ("function", "get"):
                # Inside a Function, its own name with an argument list
                # is a recursive call; without one it is the return
                # value.  That is VBA's rule, and it is what makes
                # Fact = n * Fact(n - 1) work.
                positional, named = self._arguments(args, frame, procedure=frame.procedure)
                return self.call(frame.procedure, frame.module, positional, named, me=frame.me)
            slot = frame.locals.get(key)
            if slot is not None:
                return self._index_or_call(slot.get(), args, frame)
            if isinstance(frame.me, UserClassInstance) and key in frame.me.variables:
                return self._index_or_call(frame.me.variables[key].get(), args, frame)
            module = frame.module
            module.initialise()
            if key in module.variables:
                return self._index_or_call(module.variables[key].get(), args, frame)
            if key in module.constants:
                return self._index_or_call(module.constants[key], args, frame)
            procedure = module.procedure(key)
            if procedure is not None:
                positional, named = self._arguments(args, frame, procedure=procedure)
                return self.call(procedure, module, positional, named, me=frame.me if module.is_class else None)
            host = _host_of(frame)
            if host is not None and host.vba_member(key) is not None:
                positional, named = self._arguments(args, frame, target=host, name=key)
                return _dispatch_get(host, key, positional, named)
            if key in module.declares:
                raise VBAUnsupportedError(
                    f"{module.declares[key].name} is a Declare into "
                    f"{module.declares[key].library or 'a library'}, which pyOpenVBA cannot call"
                )
        for runtime in self.modules.values():
            if runtime.is_class:
                continue
            runtime.initialise()
            if key in runtime.variables:
                return self._index_or_call(runtime.variables[key].get(), args, frame)
            procedure = runtime.procedure(key)
            if procedure is not None and procedure.scope != "private":
                positional, named = self._arguments(args, frame, procedure=procedure)
                return self.call(procedure, runtime, positional, named)
            if key in runtime.declares and runtime.declares[key].name:
                raise VBAUnsupportedError(
                    f"{runtime.declares[key].name} is a Declare into "
                    f"{runtime.declares[key].library or 'a library'}, which pyOpenVBA cannot call"
                )
        if key in ("len", "lenb"):
            measured = self._len_of_declared(args, frame)
            if measured is not None:
                return measured
        if key in self._intrinsics:
            positional, named = self._arguments(args, frame, intrinsic=key)
            return self.call_intrinsic(key, positional, named)
        supplied = self.host.global_object(key)
        if supplied is not UNRESOLVED:
            positional, named = self._arguments(args, frame, target=supplied)
            return _call_value(supplied, positional, named)
        if self.host.has_global_member(key):
            positional, named = self._arguments(args, frame)
            return self.host.call_global(key, positional, named)
        constant = self.host.constant(key)
        if constant is not UNRESOLVED:
            positional, named = self._arguments(args, frame)
            return _call_value(constant, positional, named)
        raise VBARuntimeError(35, f"Sub or Function not defined: {callee.name}")

    def _len_of_declared(self, args: list[A.Argument], frame: Frame | None) -> object | None:
        """Len of a variable that is not a String is its size in bytes.

        VBA settles this at compile time from the declaration, which is
        why Len(n) is 4 for a Long holding 12345 while Len of the same
        number in a Variant is 5.  Measured against Excel.
        """
        if frame is None or len(args) != 1 or args[0].name or not isinstance(args[0].value, A.Name):
            return None
        slot = self.lookup(args[0].value.name, frame, want_slot=True)
        if not isinstance(slot, Slot):
            return None
        width = STORAGE_WIDTH.get(slot.declared)
        return VBAInt(width, "Integer") if width is not None else None

    def _index_or_call(self, value: object, args: list[A.Argument], frame: Frame | None) -> object:
        positional, named = self._arguments(args, frame, target=value)
        return _call_value(value, positional, named)

    def _arguments(
        self,
        args: Sequence[A.Argument],
        frame: Frame | None,
        *,
        procedure: A.Procedure | None = None,
        target: object = None,
        name: str = "",
        intrinsic: str = "",
    ) -> tuple[list[object], dict[str, object]]:
        """Evaluate a call's arguments, keeping ByRef ones as slots."""
        positional: list[object] = []
        named: dict[str, object] = {}
        by_ref = _by_ref_positions(procedure)
        trailing = len(args)
        while trailing and args[trailing - 1].omitted and not args[trailing - 1].name:
            trailing -= 1
        for index, argument in enumerate(args[:trailing]):
            if argument.omitted:
                if argument.name:
                    continue
                positional.append(MISSING)
                continue
            wants_slot = procedure is not None and by_ref(index, argument.name)
            value: object
            if wants_slot and isinstance(argument.value, A.Name):
                slot = self.lookup(argument.value.name, frame, want_slot=True)
                if isinstance(slot, Slot):
                    value = slot
                else:
                    value = self.evaluate(argument.value, frame, want_object=True)
            else:
                value = self.evaluate(argument.value, frame, want_object=True)
            if argument.name:
                named[argument.name] = value
            else:
                positional.append(value)
        return positional, named

    def call_named(
        self,
        name: str,
        args: list[object],
        named: dict[str, object],
        module: ModuleRuntime,
        *,
        me: object = None,
    ) -> object:
        procedure = module.procedure(name)
        if procedure is None:
            raise VBARuntimeError(35, f"Sub or Function not defined: {name}")
        return self.call(procedure, module, args, named, me=me)

    def call_intrinsic(self, name: str, args: list[object], named: dict[str, object]) -> object:
        function = self._intrinsics[name]
        return function(self, [self._by_value(one) for one in args], named)

    def create(self, type_name_: str) -> object:
        """``New`` on a class module, a Collection, or a host class."""
        key = type_name_.lower()
        if key == "collection":
            return Collection()
        runtime = self.modules.get(key)
        if runtime is not None and runtime.is_class:
            runtime.initialise()
            instance = UserClassInstance(runtime, self)
            initialise = runtime.procedure("Class_Initialize")
            if initialise is not None:
                self.call(initialise, runtime, [], {}, me=instance)
            return instance
        if key in ("scripting.dictionary", "dictionary"):
            # New Dictionary, with a reference to the Scripting Runtime; a class module of that name comes first.
            from pyopenvba.interpreter._scripting import Dictionary

            return Dictionary()
        made = self.host.create(type_name_)
        if made is not UNRESOLVED:
            return made
        raise VBAUnsupportedError(f"New {type_name_} is not a class pyOpenVBA can make")

    # --- targets -----------------------------------------------------------------------

    def resolve_target(self, expression: A.Expr | None, frame: Frame) -> Assignable:
        """Where a value goes when a statement assigns to ``expression``."""
        if isinstance(expression, A.Name):
            return self._name_target(expression, frame)
        if isinstance(expression, A.Member):
            target = self._member_target(expression, frame)
            name = expression.name
            return Assignable(
                put=lambda value, by_ref: _dispatch_set(target, name, value, [], {}, by_ref),
                take=lambda: _dispatch_get(target, name, [], {}),
            )
        if isinstance(expression, A.Index):
            return self._index_target(expression, frame)
        if isinstance(expression, A.Bang):
            target = self._member_target(expression, frame)
            name = expression.name
            return Assignable(
                put=lambda value, by_ref: _dispatch_set(target, "Item", value, [name], {}, by_ref),
                take=lambda: _dispatch_get(target, "Item", [name], {}),
            )
        raise VBACompileError("this cannot be assigned to")

    def _name_target(self, expression: A.Name, frame: Frame) -> Assignable:
        key = _bare(expression.name)
        slot = frame.locals.get(key)
        if slot is None and isinstance(frame.me, UserClassInstance) and key in frame.me.variables:
            slot = frame.me.variables[key]
        if slot is None:
            frame.module.initialise()
            slot = frame.module.variables.get(key)
        host = _host_of(frame)
        if slot is None and host is not None and host.vba_member(key) is not None:
            # A sheet's own module sets the sheet's members by name: Name = "x" renames the sheet.
            owner = host
            return Assignable(put=lambda value, by_ref: owner.vba_set(key, value, by_ref=by_ref),
                              take=lambda: owner.vba_get(key))
        if slot is None:
            for runtime in self.modules.values():
                if runtime.is_class:
                    continue
                runtime.initialise()
                if key in runtime.variables:
                    slot = runtime.variables[key]
                    break
        if slot is None:
            if isinstance(frame.me, UserClassInstance):
                mine = frame.me.module.procedure(key, "let") or frame.me.module.procedure(key, "set")
                if mine is not None:
                    instance = frame.me

                    def put_on_instance(value: object, by_ref: bool) -> None:
                        self.call(mine, instance.module, [value], {}, me=instance)

                    return Assignable(put=put_on_instance, take=lambda: instance.vba_get(key))
            procedure, runtime = self.find_procedure(key)
            if procedure is not None and procedure.kind in ("let", "set") and runtime is not None:
                found = procedure
                owner = runtime

                def put_on_module(value: object, by_ref: bool) -> None:
                    self.call(found, owner, [value], {})

                return Assignable(put=put_on_module, take=lambda: EMPTY)
            if self.host.has_global_member(key):
                return Assignable(
                    put=lambda value, by_ref: self.host.set_global(key, value, by_ref),
                    take=lambda: self.host.call_global(key, [], {}),
                )
            if self.option_explicit(frame):
                raise VBACompileError(
                    f"Variable not defined: {expression.name}",
                    where=f"{frame.module.name}.{frame.procedure.name} line {expression.line}",
                )
            slot = Slot(EMPTY, "Variant")
            frame.locals[key] = slot
        target = slot
        return Assignable(put=lambda value, by_ref: target.set(value), take=target.get)

    def option_explicit(self, frame: Frame) -> bool:
        return frame.module.parsed.option_explicit

    def _index_target(self, expression: A.Index, frame: Frame) -> Assignable:
        inner = expression.target
        args = [self.evaluate(argument.value, frame) for argument in expression.args if argument.value is not None]
        named = {
            argument.name: self.evaluate(argument.value, frame)
            for argument in expression.args
            if argument.name and argument.value is not None
        }
        positional = [value for argument, value in zip(expression.args, args) if not argument.name]
        if isinstance(inner, A.Name):
            key = _bare(inner.name)
            slot = frame.locals.get(key)
            if slot is None and isinstance(frame.me, UserClassInstance):
                slot = frame.me.variables.get(key)
            if slot is None:
                frame.module.initialise()
                slot = frame.module.variables.get(key)
            if slot is not None and isinstance(slot.value, VBAArray):
                array = slot.value
                subscripts = [int(to_integer(one, "Long")) for one in positional]
                return Assignable(
                    put=lambda value, by_ref: array.set(subscripts, value),
                    take=lambda: array.get(subscripts),
                )
            if slot is not None and isinstance(slot.value, VBAObject):
                held = slot.value
                return Assignable(
                    put=lambda value, by_ref: _dispatch_set(held, "Item", value, positional, named, by_ref),
                    take=lambda: _dispatch_get(held, "Item", positional, named),
                )
            procedure, runtime = self.find_procedure(key)
            if procedure is not None and runtime is not None:
                found = procedure.name
                owner = runtime
                return Assignable(
                    put=lambda value, by_ref: self._property_put(owner, found, positional, named, value, by_ref),
                    take=lambda: self.call_named(found, list(positional), named, owner),
                )
            host = self.host.global_object("application") if self.host.has_global_member(key) else UNRESOLVED
            if isinstance(host, VBAObject):
                # Range("A1") = 5: the host's own member, as Application.Range("A1") = 5 would be.
                application = host
                return Assignable(
                    put=lambda value, by_ref: _dispatch_set(application, key, value, positional, named, by_ref),
                    take=lambda: _dispatch_get(application, key, positional, named),
                )
        if isinstance(inner, A.Member):
            target = self._member_target(inner, frame)
            name = inner.name
            return Assignable(
                put=lambda value, by_ref: _dispatch_set(target, name, value, positional, named, by_ref),
                take=lambda: _dispatch_get(target, name, positional, named),
            )
        value = self.evaluate(inner, frame, want_object=True)
        if isinstance(value, VBAObject):
            held = value
            return Assignable(
                put=lambda put, by_ref: _dispatch_set(held, "Item", put, positional, named, by_ref),
                take=lambda: _dispatch_get(held, "Item", positional, named),
            )
        raise VBACompileError("this cannot be assigned to")

    def _property_put(
        self,
        module: ModuleRuntime,
        name: str,
        args: list[object],
        named: dict[str, object],
        value: object,
        by_ref: bool,
    ) -> None:
        procedure = module.procedure(name, "set" if by_ref else "let") or module.procedure(name, "let")
        if procedure is None:
            raise error(ERR_MEMBER_NOT_FOUND, f"{name} has no Property Let")
        self.call(procedure, module, [*args, value], named)

    def make_slot(self, declaration: A.VarDecl, module: ModuleRuntime) -> Slot:
        """A fresh variable, array bounds and all."""
        declared = declaration.declared
        if declaration.is_array:
            if declaration.bounds:
                bounds = self._bounds(declaration, None, module=module)
                return Slot(VBAArray(bounds, element_type=declared, fixed=True), "Variant")
            return Slot(VBAArray([(0, -1)], element_type=declared), "Variant")
        if declaration.as_new:
            return Slot(self.create(declared), declared)
        lower = declared.lower()
        if lower in module.types:
            return Slot(UserTypeValue(module.types[lower], self, module), declared)
        for runtime in self.modules.values():
            if lower in runtime.types:
                return Slot(UserTypeValue(runtime.types[lower], self, runtime), declared)
        return Slot(default_for(declared), declared)

    def _bounds(
        self, declaration: A.VarDecl, frame: Frame | None, *, module: ModuleRuntime | None = None
    ) -> list[tuple[int, int]]:
        base = (module or (frame.module if frame else None))
        option_base = base.parsed.option_base if base is not None else 0
        bounds: list[tuple[int, int]] = []
        for lower, upper in declaration.bounds:
            if lower is None:
                bounds.append((option_base, int(to_integer(self.evaluate(upper, frame), "Long"))))
            else:
                bounds.append(
                    (
                        int(to_integer(self.evaluate(lower, frame), "Long")),
                        int(to_integer(self.evaluate(upper, frame), "Long")),
                    )
                )
        return bounds


# --- helpers ---------------------------------------------------------------------------


class _Unresolved:
    def __repr__(self) -> str:
        return "<unresolved>"


UNRESOLVED: Final = _Unresolved()


class HostBridge:
    """What a host application offers a VBA project.

    The default bridge offers nothing, which is enough to run language
    tests; :mod:`pyopenvba.apps.excel` supplies a real one.
    """

    def global_object(self, name: str) -> object:
        return UNRESOLVED

    def has_global_member(self, name: str) -> bool:
        return False

    def call_global(self, name: str, args: list[object], named: dict[str, object]) -> object:
        raise VBAUnsupportedError(f"{name} is not something this host provides")

    def set_global(self, name: str, value: object, by_ref: bool) -> None:
        raise VBAUnsupportedError(f"{name} is not something this host provides")

    def constant(self, name: str) -> object:
        return UNRESOLVED

    def create(self, type_name_: str) -> object:
        return UNRESOLVED

    def evaluate_bracket(self, text: str) -> object:
        raise VBAUnsupportedError("[] needs a host that can evaluate it")


def _bare(name: str) -> str:
    from pyopenvba.interpreter._lex import strip_suffix

    return strip_suffix(name).lower()


def _host_of(frame: Frame) -> VBAObject | None:
    """The object a document module running in ``frame`` is the code of, if it is one."""
    me = frame.me
    return me.host if isinstance(me, UserClassInstance) else None


@lru_cache(maxsize=1)
def _vba_constants() -> dict[str, int | float | str]:
    from pyopenvba.interpreter._constants_data import VBA_CONSTANTS

    return {name.lower(): value for name, value in VBA_CONSTANTS.items()}


def vba_constant(name: str) -> object:
    """A constant of the VBA library itself, which every host has.

    vbCrLf and vbYes belong to the language rather than to Excel, so
    they resolve with no host at all. A whole number is a Long, as an
    enum's member is however small, unless it is declared an Integer, as
    the key codes are.
    """
    from pyopenvba.interpreter._constants_data import INTEGER_CONSTANTS

    found = _vba_constants().get(name)
    if found is None:
        return UNRESOLVED
    if isinstance(found, int) and not isinstance(found, bool):
        return VBAInt(found, "Integer" if name.lower() in INTEGER_CONSTANTS else "Long")
    return found


def _scalar(value: object) -> object:
    return value.vba_value() if isinstance(value, VBAObject) else value


def _dispatch_get(target: object, name: str, args: list[object], named: dict[str, object]) -> object:
    if isinstance(target, VBAObject):
        return target.vba_get(name, args, named)
    if target is NOTHING:
        raise error(ERR_OBJECT_VARIABLE_NOT_SET)
    raise error(424, f"{type_name(target)} has no member named {name}")


def _dispatch_set(
    target: object, name: str, value: object, args: list[object], named: dict[str, object], by_ref: bool
) -> None:
    if isinstance(target, VBAObject):
        target.vba_set(name, value, args, named, by_ref=by_ref)
        return
    if target is NOTHING:
        raise error(ERR_OBJECT_VARIABLE_NOT_SET)
    raise error(424, f"{type_name(target)} has no member named {name}")


def _call_value(value: object, args: list[object], named: dict[str, object]) -> object:
    """``value(args)`` where value is already in hand."""
    if isinstance(value, VBAArray):
        if not args:
            return value
        return value.get([int(to_integer(one, "Long")) for one in args])
    if isinstance(value, VBAObject):
        if not args and not named:
            return value
        spec = value.vba_default_member()
        if spec is None:
            raise error(ERR_MEMBER_NOT_FOUND, f"{value.vba_type_name} has no default member")
        return value.vba_get(spec.name, args, named)
    if isinstance(value, str) and args:
        raise error(ERR_TYPE_MISMATCH, "a String cannot be subscripted")
    if args:
        raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
    return value


def _iterate(subject: object) -> Iterator[object]:
    if isinstance(subject, VBAArray):
        return iter(subject.elements())
    if isinstance(subject, VBAObject):
        return subject.vba_iterate()
    if subject is NOTHING:
        raise error(ERR_OBJECT_VARIABLE_NOT_SET)
    raise error(451, f"{type_name(subject)} cannot be walked by For Each")


def _is_type(value: object, wanted: str) -> bool:
    if value is NOTHING:
        return False
    if isinstance(value, VBAObject):
        if value.vba_type_name.lower() == wanted.lower():
            return True
        return any(base.__name__.lower() == wanted.lower() for base in type(value).__mro__)
    return type_name(value).lower() == wanted.lower()


def _label_index(body: Sequence[A.Stmt], label: str) -> int | None:
    for index, statement in enumerate(body):
        if isinstance(statement, A.Label) and statement.name.lower() == label.lower():
            return index
    return None


def _nested_bodies(statement: A.Stmt) -> Iterator[list[A.Stmt]]:
    if isinstance(statement, A.If):
        for branch in statement.branches:
            yield branch.body
        yield statement.otherwise
    elif isinstance(statement, (A.For, A.ForEach, A.DoLoop, A.WhileLoop, A.With)):
        yield statement.body
    elif isinstance(statement, A.SelectCase):
        for block in statement.blocks:
            yield block.body
        if statement.otherwise is not None:
            yield statement.otherwise


def _default_button(style: int) -> int:
    """Which button MsgBox would come back with if nobody clicked."""
    return {0: 1, 1: 1, 2: 3, 3: 6, 4: 6, 5: 4}.get(style & 0x0F, 1)


def _exit_word(kind: str) -> str:
    if kind == "sub":
        return "sub"
    if kind == "function":
        return "function"
    return "property"


def _by_ref_positions(procedure: A.Procedure | None) -> Callable[[int, str], bool]:
    if procedure is None:
        return lambda index, name: False

    def wanted(index: int, name: str) -> bool:
        if name:
            for parameter in procedure.params:
                if parameter.name.lower() == name.lower():
                    return not parameter.by_val
            return False
        if index < len(procedure.params):
            return not procedure.params[index].by_val
        return False

    return wanted


def _load_intrinsics() -> dict[str, Callable[[Interpreter, list[object], dict[str, object]], object]]:
    from pyopenvba.interpreter._intrinsics import INTRINSICS

    return INTRINSICS


_STATEMENTS: Final[dict[type, Callable[..., None]]] = (
    Interpreter.statement_handlers()
)
