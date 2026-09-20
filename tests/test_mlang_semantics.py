"""The M evaluator against what Power Query answered.

Every probe in ``fixtures/mlang/probes.txt`` was evaluated by live
Excel's own Power Query engine through ``scripts/measure_m.py``, and
what it answered is in ``measured.json``.  This replays them here and
requires the same answer: the type M gave the value and a short
rendering of it, or the error it raised.

The description is written twice, once in M for the measuring and once
here, and the two have to agree character for character.
"""

from __future__ import annotations

import datetime as _dt
import json
import math
from pathlib import Path

import pytest

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.mlang import Duration, MError, MSyntaxError, Record, Table, evaluate_query
from pyopenvba.mlang._values import Builtin, Function, MType, Scaled, duration_text, is_list

FIXTURES = Path(__file__).parent / "fixtures" / "mlang"


def read_probes() -> list[str]:
    out: list[str] = []
    for line in (FIXTURES / "probes.txt").read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith("//"):
            out.append(stripped)
    return out


def read_measured() -> dict[str, str]:
    return json.loads((FIXTURES / "measured.json").read_text(encoding="utf-8"))["probes"]


def text_of(value: object) -> str:
    """``Text.From`` as the measuring side wrote it."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, Scaled) and value.places:
        return f"{float(value):.{value.places}f}"
    if isinstance(value, float):
        if math.isinf(value):
            return "∞" if value > 0 else "-∞"
        if math.isnan(value):
            return "NaN"
        return str(int(value)) if value.is_integer() else repr(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, _dt.datetime):
        return value.strftime("%m/%d/%Y %H:%M:%S")
    if isinstance(value, _dt.date):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, _dt.time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, Duration):
        return duration_text(value)
    return str(value)


def describe(value: object) -> str:
    if value is None:
        return "null|null"
    if isinstance(value, Table):
        rows = ";".join(",".join(text_of(one) for one in row) for row in value.rows)
        return f"table|{value.height}x{value.width}|{','.join(value.columns)}|{rows}"
    if isinstance(value, Record):
        values = ",".join(text_of(one) for one in value.fields.values())
        return f"record|{','.join(value.names)}|{values}"
    if is_list(value):
        return f"list|{len(value)}|{','.join(text_of(one) for one in value)}"
    if isinstance(value, (Function, Builtin)):
        return "function|"
    if isinstance(value, bool):
        return f"logical|{text_of(value)}"
    if isinstance(value, (int, float)):
        return f"number|{text_of(value)}"
    if isinstance(value, str):
        return f"text|{value}"
    if isinstance(value, _dt.datetime):
        return f"datetime|{value.strftime('%Y-%m-%d %H:%M:%S')}"
    if isinstance(value, _dt.date):
        return f"date|{value.strftime('%Y-%m-%d')}"
    if isinstance(value, _dt.time):
        # Time.ToText writes a clock, not a stopwatch: 12:30 PM.
        hour = value.hour % 12 or 12
        suffix = "AM" if value.hour < 12 else "PM"
        return f"time|{hour}:{value.minute:02d} {suffix}"
    if isinstance(value, Duration):
        return f"duration|{text_of(value)}"
    if isinstance(value, MType):
        return "other|"
    return "other|"


def evaluate(probe: str) -> str:
    try:
        return describe(evaluate_query(probe))
    except MError as failure:
        return f"!{failure.reason}"
    except MSyntaxError:
        return "!syntax"
    except VBAUnsupportedError as gap:
        return f"!unsupported: {gap}"


PROBES = read_probes()
MEASURED = read_measured()

#: Probes whose measured answer is not a comparison of behaviour: the
#: describing itself failed on the Power Query side, so there is nothing
#: to hold the evaluator to.
UNCOMPARABLE = frozenset(name for name, answer in MEASURED.items() if answer in ("", "!unmeasured"))


def test_every_probe_was_measured() -> None:
    missing = [probe for probe in PROBES if probe not in MEASURED]
    assert not missing, f"{len(missing)} probes have no measured answer, first {missing[0]!r}"


@pytest.mark.parametrize(
    "probe", [one for one in PROBES if one not in UNCOMPARABLE], ids=lambda one: str(PROBES.index(one))
)
def test_matches_power_query(probe: str) -> None:
    want = MEASURED[probe]
    got = evaluate(probe)
    assert got == want, f"{probe}\n  Power Query: {want}\n  ours       : {got}"
