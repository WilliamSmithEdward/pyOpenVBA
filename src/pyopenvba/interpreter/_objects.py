"""How VBA reaches an object, and how an object answers.

Everything a macro can put a dot after implements :class:`VBAObject`.  A
member is an ordinary Python method carrying a marker, and its Python
signature is its VBA signature: the parameter names are what a named
argument matches, and a parameter defaulting to ``MISSING`` is what
``IsMissing`` reports as missing.

An unknown member is where the three errors separate.  If the real
application has the member and this does not implement it, that is
:class:`~pyopenvba.exceptions.VBAUnsupportedError`; if the application
has no such member either, that is VBA's run-time error 438.  The
difference is decided from the type library inventory in
:mod:`pyopenvba.interpreter._inventory`, so it is measured rather than
guessed.
"""

from __future__ import annotations

import inspect
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass
from typing import Any, ClassVar

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import (
    ERR_MEMBER_NOT_FOUND,
    ERR_OBJECT_VARIABLE_NOT_SET,
    MISSING,
    NOTHING,
    VBAArray,
    error,
    to_integer,
)


@dataclass(slots=True)
class MemberSpec:
    """One member of a host class, and how to reach it."""

    name: str
    kind: str
    getter: Callable[..., Any] | None = None
    setter: Callable[..., Any] | None = None
    ref_setter: Callable[..., Any] | None = None
    default: bool = False
    parameters: tuple[str, ...] = ()
    #: True where the Python method takes *args, as Debug.Print does.
    varargs: bool = False

    @property
    def settable(self) -> bool:
        return self.setter is not None or self.ref_setter is not None


def member(
    function: Callable[..., Any] | None = None,
    *,
    name: str = "",
    default: bool = False,
    kind: str = "property",
) -> Any:
    """Mark a method as a VBA member of the class it is defined in."""

    def mark(target: Callable[..., Any]) -> Callable[..., Any]:
        target._vba_member = (name or target.__name__, kind, default)  # type: ignore[attr-defined]
        return target

    return mark(function) if function is not None else mark


def method(function: Callable[..., Any] | None = None, *, name: str = "", default: bool = False) -> Any:
    """Mark a method as a VBA method rather than a property."""
    return member(function, name=name, default=default, kind="method")


def setter(name: str) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Mark a method as the ``Property Let`` of ``name``."""

    def mark(target: Callable[..., Any]) -> Callable[..., Any]:
        target._vba_setter = name  # type: ignore[attr-defined]
        return target

    return mark


def ref_setter(name: str) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Mark a method as the ``Property Set`` of ``name``."""

    def mark(target: Callable[..., Any]) -> Callable[..., Any]:
        target._vba_ref_setter = name  # type: ignore[attr-defined]
        return target

    return mark


class VBAObject:
    """The base every object a macro can touch derives from."""

    #: The name TypeName() gives an instance, and the key into the
    #: member inventory.  Not a ClassVar: a user class module names its
    #: own instances, so an instance may carry its own.
    vba_type_name: str = "Object"

    #: The library the type belongs to, for the inventory lookup.
    vba_library: str = ""

    _vba_members: ClassVar[dict[str, MemberSpec]] = {}

    def __init_subclass__(cls, **kwargs: Any) -> None:
        super().__init_subclass__(**kwargs)
        members: dict[str, MemberSpec] = {}
        for base in reversed(cls.__mro__[1:]):
            members.update(getattr(base, "_vba_members", {}))
        pending_setters: list[tuple[str, Callable[..., Any], bool]] = []
        for attribute in vars(cls).values():
            if not callable(attribute):
                continue
            marked = getattr(attribute, "_vba_member", None)
            if marked is not None:
                name, kind, default = marked
                signature = inspect.signature(attribute)
                declared = list(signature.parameters.values())[1:]
                varargs = any(one.kind is inspect.Parameter.VAR_POSITIONAL for one in declared)
                parameters = tuple(
                    one.name for one in declared if one.kind is not inspect.Parameter.VAR_POSITIONAL
                )
                members[name.lower()] = MemberSpec(
                    name=name,
                    kind=kind,
                    getter=attribute,
                    default=default,
                    parameters=parameters,
                    varargs=varargs,
                )
            named = getattr(attribute, "_vba_setter", None)
            if named is not None:
                pending_setters.append((named, attribute, False))
            named = getattr(attribute, "_vba_ref_setter", None)
            if named is not None:
                pending_setters.append((named, attribute, True))
        for name, function, by_ref in pending_setters:
            spec = members.get(name.lower())
            if spec is None:
                spec = MemberSpec(name=name, kind="property")
            spec = MemberSpec(
                name=spec.name,
                kind=spec.kind,
                getter=spec.getter,
                setter=function if not by_ref else spec.setter,
                ref_setter=function if by_ref else spec.ref_setter,
                default=spec.default,
                parameters=spec.parameters,
            )
            members[name.lower()] = spec
        cls._vba_members = members
        if "vba_type_name" not in vars(cls):
            cls.vba_type_name = cls.__name__

    @classmethod
    def vba_add_member(cls, spec: MemberSpec) -> None:
        """Register a member made at run time rather than written as a method, unless the class has one of the name."""
        cls._vba_members.setdefault(spec.name.lower(), spec)

    # --- dispatch ---------------------------------------------------------------

    def vba_member(self, name: str) -> MemberSpec | None:
        return self._vba_members.get(name.lower())

    def vba_default_member(self) -> MemberSpec | None:
        for spec in self._vba_members.values():
            if spec.default:
                return spec
        return None

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        spec = self.vba_member(name)
        if spec is None or spec.getter is None:
            raise self.vba_no_member(name)
        if args and not named and spec.kind == "property" and not spec.parameters and not spec.varargs:
            # VBA hands arguments a property does not take to what it answers: Range("C1:C3").Formula(2, 1)
            # is an item of the array Formula answers, and an object's default member takes them.
            found = spec.getter(self)
            if isinstance(found, VBAArray):
                return found.get([int(to_integer(one, "Long")) for one in args])
            default = found.vba_default_member() if isinstance(found, VBAObject) else None
            if isinstance(found, VBAObject) and default is not None:
                return found.vba_get(default.name, args)
        return spec.getter(self, *self._bind(spec, args, named))

    def vba_set(
        self,
        name: str,
        value: object,
        args: Sequence[object] = (),
        named: dict[str, object] | None = None,
        *,
        by_ref: bool = False,
    ) -> None:
        spec = self.vba_member(name)
        if spec is None:
            raise self.vba_no_member(name)
        function = spec.ref_setter if by_ref and spec.ref_setter is not None else spec.setter
        if function is None and by_ref and spec.setter is not None:
            function = spec.setter
        if function is None:
            if not by_ref and spec.getter is not None:
                # A read-only member that answers an object takes a value in
                # that object's default member: ws.Range("A1") = 5 fills A1.
                found = spec.getter(self, *self._bind(spec, args, named))
                default = found.vba_default_member() if isinstance(found, VBAObject) else None
                if isinstance(found, VBAObject) and default is not None and default.setter is not None:
                    found.vba_set(default.name, value)
                    return
            raise error(ERR_MEMBER_NOT_FOUND, f"{self.vba_type_name}.{spec.name} cannot be assigned to")
        if _takes_only_the_value(function):
            # The common shape: a setter that ignores the getter's
            # arguments, as Range.Value's does.
            function(self, value)
            return
        function(self, *self._bind(spec, args, named), value)

    def vba_call(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        return self.vba_get(name, args, named)

    def _bind(
        self, spec: MemberSpec, args: Sequence[object], named: dict[str, object] | None
    ) -> list[object]:
        """Positional and named arguments laid out in the member's order."""
        if spec.varargs:
            if named:
                raise error(446, f"{self.vba_type_name}.{spec.name} does not take named arguments")
            return list(args)
        if not spec.parameters:
            if args or named:
                raise error(450, f"{self.vba_type_name}.{spec.name} takes no arguments")
            return []
        laid: list[object] = [MISSING] * len(spec.parameters)
        if len(args) > len(spec.parameters):
            raise error(450, f"{self.vba_type_name}.{spec.name} takes {len(spec.parameters)} arguments")
        for index, value in enumerate(args):
            laid[index] = value
        for key, value in (named or {}).items():
            for index, parameter in enumerate(spec.parameters):
                if parameter.lower() == key.lower():
                    laid[index] = value
                    break
            else:
                raise error(448, f"{self.vba_type_name}.{spec.name} has no argument named {key}")
        return laid

    def vba_no_member(self, name: str) -> Exception:
        """The right error for a member this object does not answer."""
        from pyopenvba.interpreter._inventory import member_exists, type_known

        if type_known(self.vba_type_name, self.vba_library):
            if member_exists(self.vba_type_name, name, self.vba_library):
                return VBAUnsupportedError(
                    f"{self.vba_type_name}.{name} is a real member that pyOpenVBA does not implement"
                )
            return error(ERR_MEMBER_NOT_FOUND, f"{self.vba_type_name} has no member named {name}")
        return VBAUnsupportedError(f"{self.vba_type_name}.{name} is not implemented by pyOpenVBA")

    # --- what the language needs from every object ------------------------------

    def vba_value(self) -> object:
        """The scalar an object turns into where a value is wanted."""
        spec = self.vba_default_member()
        if spec is None or spec.getter is None:
            raise error(ERR_MEMBER_NOT_FOUND, f"{self.vba_type_name} has no default member")
        return spec.getter(self, *([MISSING] * len(spec.parameters)))

    def vba_iterate(self) -> Iterator[object]:
        """What ``For Each`` walks.  Not every object has one."""
        raise error(451, f"{self.vba_type_name} cannot be walked by For Each")

    def describe(self, indent: str = "") -> str:
        """One developer-readable line, or a block for a container."""
        return f"{indent}{self.vba_type_name}"

    def __repr__(self) -> str:
        return f"<{self.vba_type_name}>"


class VBACollection(VBAObject):
    """A collection: Count, Item as the default member, and For Each."""

    vba_type_name = "Collection"

    def vba_items(self) -> list[object]:  # pragma: no cover - overridden
        raise NotImplementedError

    @member
    def Count(self) -> object:
        from pyopenvba.interpreter._values import VBAInt

        return VBAInt(len(self.vba_items()), "Long")

    @member(default=True)
    def Item(self, Index: object = MISSING) -> object:
        items = self.vba_items()
        if Index is MISSING:
            raise error(449)
        return self.vba_lookup(Index, items)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        from pyopenvba.interpreter._values import ERR_SUBSCRIPT_OUT_OF_RANGE, to_integer

        position = int(to_integer(index, "Long"))
        if 1 <= position <= len(items):
            return items[position - 1]
        raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)

    def vba_iterate(self) -> Iterator[object]:
        return iter(self.vba_items())


def _takes_only_the_value(function: Callable[..., Any]) -> bool:
    """Whether a Property Let was written as ``(self, value)`` alone."""
    cached = getattr(function, "_vba_only_value", None)
    if cached is None:
        parameters = list(inspect.signature(function).parameters.values())
        cached = len(parameters) == 2 and all(
            one.kind is not inspect.Parameter.VAR_POSITIONAL for one in parameters
        )
        function._vba_only_value = cached  # type: ignore[attr-defined]
    return cached


def require_object(value: object, *, where: str = "") -> VBAObject:
    """``value`` as an object, or the error VBA raises when it is not one."""
    if isinstance(value, VBAObject):
        return value
    if value is NOTHING:
        raise error(ERR_OBJECT_VARIABLE_NOT_SET, where=where)
    raise error(424, "Object required", where=where)
