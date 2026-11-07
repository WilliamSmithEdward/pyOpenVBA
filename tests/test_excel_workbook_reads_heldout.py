"""Held-out rounds of the Workbook reads (scripts/measure_workbook_reads_heldout.py): books chosen from a seed, built and saved by a macro,
opened again and read typed and late, with the model's answers written before live Excel's.

Each round's predictions stay as they were written, digest and all, so the miss rate they had is kept. The model then replays every book of
the round from the specs Excel's answers record. DoNotPromptForConvert is left out of the comparison: it is a user-level setting, kept in the
registry across workbooks, and round1's first book set it False for the rest of the round. Round1 missed 20 of 864 reads: 8 of
DoNotPromptForConvert, for that reason, and 12 where the model saved a macro-enabled book with no code with a VBA project, which Excel
does not write.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
FOLDER = ROOT / 'tests/fixtures/workbook_reads_heldout'
_SPEC = importlib.util.spec_from_file_location('measure_workbook_reads_heldout', ROOT / 'scripts/measure_workbook_reads_heldout.py')
assert _SPEC is not None and _SPEC.loader is not None
HELDOUT: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(HELDOUT)
#: Each round's predictions, as their digest was recorded before Excel answered (.git/excel-coordination/BOARD.md).
DIGESTS = {'round1': '3b9544fda88f7ddb7e73c74b718bffcb59b136344365bc574eebf7642e411603'}
MISSES = {'round1': 20}
#: User-level: what Excel reads depends on what an earlier book of the round set, not on the book.
USER_LEVEL = frozenset({'DoNotPromptForConvert'})
ROUNDS = sorted(path.name.removesuffix('_excel.json') for path in FOLDER.glob('*_excel.json'))


def _load(round_name: str, kind: str) -> dict[str, Any]:
    return json.loads((FOLDER / f'{round_name}_{kind}.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('round_name', ROUNDS)
def test_the_predictions_are_as_they_were_written(round_name: str) -> None:
    text = (FOLDER / f'{round_name}_predictions.json').read_bytes().replace(b'\r\n', b'\n')
    assert hashlib.sha256(text).hexdigest() == DIGESTS[round_name]


@pytest.mark.parametrize('round_name', ROUNDS)
def test_the_miss_rate_is_the_one_recorded(round_name: str) -> None:
    predicted, excel = _load(round_name, 'predictions'), _load(round_name, 'excel')
    misses = sum(a != b for mine, theirs in zip(predicted['books'], excel['books']) for declared in ('Workbook', 'Object')
                 for a, b in zip(mine['answers'][declared], theirs['answers'][declared]))
    assert misses == MISSES[round_name]


@pytest.mark.parametrize('round_name', ROUNDS)
def test_every_book_reads_as_excel_did(round_name: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    excel = _load(round_name, 'excel')
    differ: list[tuple[int, str, str, str, str]] = []
    for index, entry in enumerate(excel['books']):
        book = entry['book']
        path = tmp_path / f'Book{index}.{book["format"]}'
        builder = ExcelApplication(sandbox=False)
        builder.add_workbook()
        builder.add_module(HELDOUT.build(book), name='B')
        assert str(builder.run('B.Build', str(path))) == entry['failed_sets']
        for declared in ('Workbook', 'Object'):
            reader = ExcelApplication(sandbox=False)
            reader.add_workbook()
            reader.add_module(HELDOUT.READS.macro(declared), name='H')
            found = str(reader.run('H.Probe', str(path), book['read_only'])).split(HELDOUT.READS.SEPARATOR)
            differ += [(index, declared, name, mine, theirs) for name, mine, theirs in zip(excel['names'], found, entry['answers'][declared])
                       if mine != theirs and name not in USER_LEVEL]
    assert differ == []
