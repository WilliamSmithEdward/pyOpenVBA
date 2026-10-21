"""Optional host late-binding hook preserves ordinary VBA dispatch by default."""
from __future__ import annotations

from collections.abc import Sequence

import pytest

from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.interpreter import Interpreter
from pyopenvba.interpreter._intrinsics import INTRINSICS
from pyopenvba.interpreter._objects import VBAObject, member, method, setter
from pyopenvba.interpreter._values import EMPTY, NOTHING


class PlainHost(VBAObject):
    def __init__(self) -> None:
        self.value: object = 3

    @member
    def Value(self) -> object:
        return self.value

    @setter('Value')
    def _value(self, value: object) -> None:
        self.value = value

    @method
    def Echo(self, value: object) -> object:
        return value


class ObservedHost(PlainHost):
    def __init__(self) -> None:
        super().__init__()
        self.calls: list[tuple[str, tuple[object, ...], bool, bool]] = []

    def vba_call_by_name(self, name: str, args: Sequence[object], *, assignment: bool = False,
                         by_ref: bool = False) -> object:
        self.calls.append((name, tuple(args), assignment, by_ref))
        return super().vba_call_by_name(name, args, assignment=assignment, by_ref=by_ref)


def late(host: object, name: str, kind: int, *args: object) -> object:
    return INTRINSICS['callbyname'](Interpreter(), [host, name, kind, *args], {})


@pytest.mark.parametrize('host_type', [PlainHost, ObservedHost])
def test_default_get_let_method_and_direct_access_are_unchanged(host_type: type[PlainHost]) -> None:
    host = host_type()
    assert host.vba_get('Value') == 3
    assert late(host, 'value', 2) == 3
    assert late(host, 'Value', 4, 7) is EMPTY
    assert host.vba_get('Value') == 7
    assert late(host, 'Echo', 1, 'unchanged') == 'unchanged'
    if isinstance(host, ObservedHost):
        assert host.calls == [('value', (), False, False), ('Value', (7,), True, False),
                              ('Echo', ('unchanged',), False, False)]


def test_existing_argument_type_and_call_kind_validation_precedes_override() -> None:
    host = ObservedHost()
    for target, name, kind, args, expected in [
        (NOTHING, 'Value', 2, (), 5),
        (42, 'Value', 2, (), 424),
        (host, 'Value', 1, (), 438),
        (host, 'Value', 4, (), 438),
        (host, 'Value', 8, (3,), 438),
        (host, 'Value', 0, (), 438),
    ]:
        with pytest.raises(VBARuntimeError) as failure:
            late(target, name, kind, *args)
        assert failure.value.number == expected
    assert not host.calls
