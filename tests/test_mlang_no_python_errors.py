"""Every M library function, given awkward numbers, answers, raises an M error or refuses openly: no Python exception reaches the macro that
refreshed the query, where On Error could not treat it. The numbers are the ones that let Python errors out before: infinity, just past a 32-bit
integer, -1, a fraction, 1e300 and 9e15, each alone and after a text, a list, a table, a date, a duration and a datetime.

The calls come from Python, where the sandbox does not apply, so the walk runs in an empty temporary folder and checks it stays empty. No value
it gives names a file, a folder or a server.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.mlang import MError, MSyntaxError, evaluate_query
from pyopenvba.mlang._library import LIBRARY
from pyopenvba.mlang._values import Builtin

NUMBERS = ('1.8e308', '-1.8e308', '2147483648', '-2147483649', '-1', '1.5', '1e300', '0', '9e15')
FIRSTS = (None, '"abc"', '{1, 2, 3}', '#table({"a"}, {{1}})', '#date(2020, 1, 1)', '#duration(1, 0, 0, 0)', '#datetime(2020, 1, 1, 0, 0, 0)')


def _calls(name: str) -> list[str]:
    builtin = LIBRARY[name]
    assert isinstance(builtin, Builtin)
    top = builtin.maximum if builtin.maximum is not None else builtin.minimum + 1
    calls: list[str] = []
    for arity in range(max(builtin.minimum, 1), min(top, 3) + 1):
        for first in FIRSTS:
            for number in NUMBERS:
                args = [number] * arity
                if first is not None:
                    args[0] = first
                calls.append(f'let x = {name}({", ".join(args)}) in x')
    return calls


@pytest.mark.parametrize('name', sorted(name for name, entry in LIBRARY.items() if isinstance(entry, Builtin)))
def test_awkward_numbers_raise_nothing_of_pythons(name: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    escaped: list[str] = []
    for source in _calls(name):
        try:
            evaluate_query(source)
        except (MError, MSyntaxError, VBAUnsupportedError):
            pass
        except Exception as failed:  # noqa: BLE001 - the test is that nothing else is raised
            escaped.append(f'{source}: {type(failed).__name__}: {failed}')
    assert not escaped
    assert not any(tmp_path.iterdir())
