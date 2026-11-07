"""Every Workbook property a macro reads, on books Excel saved, typed and late, as live Excel answered (scripts/measure_workbook_reads.py).

Each round is a book Excel built and saved: a plain one opened read-only, and a macro-enabled one with its settings away from their defaults,
opened read-write. The model opens the same file Excel read, copied into an empty temporary folder under the same name, and runs the same
macro. The calls come from Python with the sandbox off, so the folder is checked to hold nothing else afterwards.
"""
from __future__ import annotations

import importlib.util
import json
import shutil
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_workbook_reads', ROOT / 'scripts/measure_workbook_reads.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD: dict[str, Any] = json.loads((ROOT / 'tests/fixtures/workbook_reads.json').read_text(encoding='utf-8'))
#: Answered here, but not as Excel did, by round and declared type. Each leaves when it is fixed.
GAPS: dict[tuple[str, str], frozenset[str]] = {}
#: Refused openly here, so left out of the model's macro for that round: the custom views a loaded book holds are not read.
REFUSED: dict[str, frozenset[str]] = {'settings': frozenset({'CustomViews'})}
#: Books the model does not open at all, so it refuses them openly: Excel's answers are kept for when it reads them.
UNOPENED = frozenset({'binary', 'legacy'})


def _model(round_name: str, declared: str, folder: Path) -> dict[str, str]:
    source = MEASURE.book(round_name)
    path = folder / f'Probe{source.suffix}'
    shutil.copyfile(source, path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    names = [name for name in RECORD['names'] if name not in REFUSED.get(round_name, frozenset())]
    app.add_module(MEASURE.macro(declared, names), name='H')
    found = str(app.run('H.Probe', str(path), MEASURE.ROUNDS[round_name][2])).split(MEASURE.SEPARATOR)
    assert sorted(item.name for item in folder.iterdir()) == [path.name]
    return dict(zip(names, found))


def test_the_names_and_rounds_are_the_ones_measured() -> None:
    assert RECORD['names'] == MEASURE.NAMES
    assert sorted(RECORD['answers']) == sorted(MEASURE.ROUNDS)


@pytest.mark.parametrize('declared', ['Workbook', 'Object'])
@pytest.mark.parametrize('round_name', sorted(set(MEASURE.ROUNDS) - UNOPENED))
def test_every_read_answers_as_excel_did(round_name: str, declared: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    excel = dict(zip(RECORD['names'], RECORD['answers'][round_name][declared]))
    model = _model(round_name, declared, tmp_path)
    differ = {name: (excel[name], model[name]) for name in model if excel[name] != model[name]}
    assert set(differ) == GAPS.get((round_name, declared), frozenset()), differ


@pytest.mark.parametrize('round_name', sorted(REFUSED))
def test_what_is_left_out_is_refused_openly(round_name: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    monkeypatch.chdir(tmp_path)
    for name in sorted(REFUSED[round_name]):
        with pytest.raises(VBAUnsupportedError):
            _model_reads(round_name, name, tmp_path)


def _model_reads(round_name: str, name: str, folder: Path) -> None:
    source = MEASURE.book(round_name)
    path = folder / f'Probe{source.suffix}'
    shutil.copyfile(source, path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURE.macro('Object', [name]), name='H')
    app.run('H.Probe', str(path), MEASURE.ROUNDS[round_name][2])


@pytest.mark.parametrize('round_name', sorted(UNOPENED))
def test_a_book_the_model_does_not_open_is_refused_openly(round_name: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel._io import WorkbookFileError
    from pyopenvba.exceptions import VBAUnsupportedError

    monkeypatch.chdir(tmp_path)
    with pytest.raises((VBAUnsupportedError, WorkbookFileError)):
        _model(round_name, 'Object', tmp_path)
