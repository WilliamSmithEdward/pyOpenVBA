"""COM classes a host supplies: what CreateObject and GetObject make, from a factory the caller registers (set_com_class).

pyOpenVBA starts no COM server and looks at no running application. A host that wants a macro to have an object registers its ProgID with a
factory, which is handed a ComRequest and returns the object, or None for a class that does not exist, which the macro sees as error 429.
Registering a class is the host's grant: the sandbox lets it through. What the object then does is the host's.

A ProgID comes from the macro, so a factory must allow-list. One that passes any name to the system's COM (win32com's Dispatch, say) hands the
macro the machine: shells, mail, databases on any server. docs/vba_runtime.md shows the pattern.

What the factory returns:

- an object of the model (a VBAObject) crosses as it is;
- any other Python object crosses wrapped in HostObject, late bound, and so does every object its members answer, at every depth.

Values cross converted, both ways, as the table in docs/vba_runtime.md says. A VBAError the factory or the object raises goes on as it is, so
either can give any error number. Any other exception is run-time error 440, VBA's Automation error, with the original as its __cause__.
"""
from __future__ import annotations

import datetime as _dt
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBAArray,
    VBACurrency,
    VBADate,
    VBADecimal,
    VBAInt,
    VBASingle,
    error,
)

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter

#: VBA's Automation error, for what a host's object raises that is no VBA error.
AUTOMATION_ERROR = 440
#: A class that does not exist, or cannot be made.
CANNOT_CREATE = 429


@dataclass(frozen=True)
class ComRequest:
    """One CreateObject or GetObject a macro asked for.

    ``prog_id`` is the name as the macro spelled it. ``kind`` is "create" for CreateObject and GetObject("", class), and "get" for
    GetObject(, class) and GetObject(path, class). ``server`` is CreateObject's ServerName, "" when none was given; no server is ever reached by
    the model. ``path`` is GetObject's file, past the sandbox's folder gate, or None. ``caller`` is the procedure making the call, as
    "Module1.Main", and ``entry`` what the host itself started the run with: the macro it ran, or the event that fired. ``document`` is the
    Workbook, Document or Presentation whose project runs the call, or None where the host has none.
    """

    prog_id: str
    kind: str
    server: str = ""
    path: Path | None = None
    caller: str = ""
    entry: str = ""
    document: object = None


Factory = Callable[[ComRequest], object]


def set_com_class(interpreter: Interpreter, prog_id: object, factory: object, folders: object = ()) -> None:
    """Register ``factory`` for ``prog_id``, or with None take it away again. A registered name wins over pyOpenVBA's own class of that name.
    ``folders`` are where a file GetObject names may be while the sandbox is on, as set_hook's are; with none, no file passes."""
    if not isinstance(prog_id, str) or not prog_id or prog_id != prog_id.strip():
        raise ValueError(f"a ProgID is a name with no spaces round it, not {prog_id!r}")
    if factory is None:
        interpreter.com_classes.pop(prog_id.lower(), None)
        return
    if not callable(factory):
        raise TypeError(f"the factory for {prog_id!r} must be callable, not {factory!r}")
    named: Sequence[object] = (folders,) if isinstance(folders, (str, Path)) else cast("Sequence[object]", folders)
    interpreter.com_classes[prog_id.lower()] = (cast("Factory", factory), tuple(str(one) for one in named))


def gated(interpreter: Interpreter, prog_id: str, path: str) -> Path:
    """The file GetObject names, refused while the sandbox is on unless it lies in a folder registered with the class, before the factory hears of
    it. The refusal says nothing of where the path leads."""
    from pyopenvba.exceptions import VBASandboxError
    from pyopenvba.interpreter._files import final_path, is_inside

    # One conversion, checked and then handed on: the factory opens exactly what the gate let through.
    final = final_path(path)
    if interpreter.sandbox and interpreter.frames:
        folders = interpreter.com_classes[prog_id.lower()][1]
        if not folders:
            interpreter.outside(f"GetObject of a file for {prog_id}")
        if not is_inside(final, folders):
            raise VBASandboxError(f"GetObject may not open that file: it is outside the folders named with set_com_class for {prog_id}",
                                  where=interpreter.current_where())
    return final


def from_factory(interpreter: Interpreter, prog_id: str, kind: str, *, server: str = "", path: Path | None = None) -> object | None:
    """What the factory registered for ``prog_id`` makes, ready for VBA; None where no factory is registered. A factory that answers None says
    there is no such class, 429."""
    registered = interpreter.com_classes.get(prog_id.lower())
    if registered is None:
        return None
    factory = registered[0]
    request = ComRequest(prog_id, kind, server, path, caller=interpreter.current_where(), entry=interpreter.steps.entry,
                         document=interpreter.host.document())
    made = _guarded(lambda: factory(request), f"the factory for {prog_id}", interpreter)
    if made is None:
        raise error(CANNOT_CREATE, "ActiveX component can't create object")
    return to_vba(made, interpreter)


def _guarded(call: Callable[[], object], what: str, interpreter: Interpreter) -> object:
    try:
        return call()
    except VBAError:
        raise
    except Exception as failed:  # noqa: BLE001 - whatever a host's code raises becomes the Automation error, its cause kept
        raise error(AUTOMATION_ERROR, f"Automation error: {what} raised {type(failed).__name__}: {failed}") from failed


# --- values ---------------------------------------------------------------------------------------------------------------------------

def to_python(value: object) -> object:
    """A VBA value as the host's object is given it."""
    if isinstance(value, HostObject):
        return value.wrapped
    if value is EMPTY or value is MISSING:
        return None
    if value is NULL or value is NOTHING or isinstance(value, (VBAObject, VBAArray)):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, VBACurrency):
        return Decimal(value)
    if isinstance(value, VBADecimal):
        return Decimal(value)
    if isinstance(value, VBADate):
        return value.to_datetime()
    if isinstance(value, int):
        return int(value)
    if isinstance(value, float):
        return float(value)
    if isinstance(value, str):
        return str(value)
    return value


def to_vba(value: object, interpreter: Interpreter) -> object:
    """What a host's object answers, as VBA has it: any object of the host's adapted, at every depth."""
    if value is None:
        return EMPTY
    if value is NULL or value is NOTHING or value is EMPTY or isinstance(value, (VBAObject, VBAArray, VBAInt, VBASingle, VBACurrency, VBADecimal,
                                                                                     VBADate)):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return VBAInt(value, "Long") if -2 ** 31 <= value < 2 ** 31 else float(value)
    if isinstance(value, float):
        return value
    if isinstance(value, Decimal):
        return VBACurrency(value)
    if isinstance(value, (_dt.datetime, _dt.date)):
        return VBADate.from_datetime(value)
    if isinstance(value, str):
        return value
    if isinstance(value, (list, tuple)):
        items = [to_vba(one, interpreter) for one in cast("Sequence[object]", value)]
        return VBAArray([(0, len(items) - 1)], items=items)
    return HostObject(value, interpreter)


# --- the adapter ----------------------------------------------------------------------------------------------------------------------

class HostObject(VBAObject):
    """A host's Python object as VBA sees it: its public attributes are its members, found case-insensitively, a method called and anything else
    read or set. Bound late, as an object with no type library is."""

    vba_type_name = "Object"

    def __init__(self, wrapped: object, interpreter: Interpreter) -> None:
        self.wrapped = wrapped
        self.interpreter = interpreter
        self.vba_type_name = str(getattr(wrapped, "vba_type_name", type(wrapped).__name__))

    def _attribute(self, name: str) -> str:
        """The wrapped object's own name for a member, or 438. Nothing that starts with an underscore is a member."""
        wanted = name.lower()
        for one in dir(self.wrapped):
            if not one.startswith("_") and one.lower() == wanted:
                return one
        raise error(438, f"Object doesn't support this property or method: {name}")

    def _arguments(self, args: Sequence[object], named: dict[str, object] | None) -> tuple[list[object], dict[str, object]]:
        positional = [to_python(one) for one in args]
        while positional and args[len(positional) - 1] is MISSING:
            positional.pop()
        return positional, {key: to_python(one) for key, one in (named or {}).items() if one is not MISSING}

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        attribute = self._attribute(name)
        found = _guarded(lambda: getattr(self.wrapped, attribute), f"{self.vba_type_name}.{name}", self.interpreter)
        positional, keywords = self._arguments(args, named)
        if callable(found):
            return to_vba(_guarded(lambda: found(*positional, **keywords), f"{self.vba_type_name}.{name}", self.interpreter), self.interpreter)
        if positional or keywords:
            # A property given arguments hands them to what it answers, as VBA does.
            return _call_value(to_vba(found, self.interpreter), list(args))
        return to_vba(found, self.interpreter)

    def vba_call(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        return self.vba_get(name, args, named)

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None, *, by_ref: bool = False) -> None:
        attribute = self._attribute(name)
        if args or named:
            raise VBAUnsupportedError(f"a property of {self.vba_type_name} set with arguments is not implemented")
        current = getattr(self.wrapped, attribute, None)
        if callable(current):
            raise error(438, f"Object doesn't support this property or method: {name}")
        _guarded(lambda: setattr(self.wrapped, attribute, to_python(value)), f"{self.vba_type_name}.{name}", self.interpreter)

    def vba_member(self, name: str) -> Any:
        return None

    def vba_default_member(self) -> Any:
        return None

    def vba_default_call(self, args: list[object], named: dict[str, object]) -> object:
        """o(1) and o("key"): what the object answers being indexed."""
        getter = getattr(self.wrapped, "__getitem__", None)
        if getter is None:
            raise error(438, f"{self.vba_type_name} has no default member")
        positional, _ = self._arguments(args, named)
        key = positional[0] if len(positional) == 1 else tuple(positional)
        return to_vba(_guarded(lambda: getter(key), f"{self.vba_type_name}(...)", self.interpreter), self.interpreter)

    def vba_value(self) -> object:
        raise error(438, f"{self.vba_type_name} has no default member")

    def vba_iterate(self) -> Iterator[object]:
        if not hasattr(self.wrapped, "__iter__"):
            raise error(438, f"{self.vba_type_name} cannot be walked by For Each")
        what = f"For Each over {self.vba_type_name}"
        iterator: Iterator[object] = _guarded(lambda: iter(self.wrapped), what, self.interpreter)  # type: ignore[call-overload, assignment]
        while True:
            # The end of the walk is StopIteration, which is no failure of the host's code.
            try:
                one = next(iterator)
            except StopIteration:
                return
            except VBAError:
                raise
            except Exception as failed:  # noqa: BLE001 - as _guarded does
                raise error(AUTOMATION_ERROR, f"Automation error: {what} raised {type(failed).__name__}: {failed}") from failed
            yield to_vba(one, self.interpreter)


def _call_value(value: object, args: list[object]) -> object:
    from pyopenvba.interpreter._runtime import _call_value as call  # pyright: ignore[reportPrivateUsage]

    return call(value, args, {})
