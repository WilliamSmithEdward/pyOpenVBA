"""FillFormat as live Excel answered four held-out rounds (scripts/measure_fill_format_round.py): the members the model has, the texture
members on fills with no texture, PictureEffects, and the picture methods up to the picture itself. Each case runs alone in the model,
typed and late, read by read, then saved, opened again and read once more, and its saved spPr is compared. A read or case the model
refuses is allowed; one it answers must be Excel's."""
from __future__ import annotations

import importlib.util
import json
import tempfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBASandboxError

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_fill_format_round', ROOT / 'scripts/measure_fill_format_round.py')
assert _SPEC is not None and _SPEC.loader is not None
PROBE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(PROBE)


@pytest.mark.parametrize('number', [1, 2, 3, 4])
def test_each_case_answers_as_excel_did(number: int, monkeypatch: pytest.MonkeyPatch) -> None:
    record: dict[str, Any] = json.loads((ROOT / f'tests/fixtures/fill_format_round_{number}.json').read_text(encoding='utf-8'))
    monkeypatch.setattr(PROBE, 'ROUND', number)
    monkeypatch.setattr(PROBE, 'READS', record['reads'])
    names = ['set', *record['reads'], *(f'reopened {one}' for one in record['reads'])]
    differ: list[tuple[Any, ...]] = []
    with tempfile.TemporaryDirectory() as folder:
        PROBE.write_files(Path(folder))
        for index, (case, excel) in enumerate(zip(record['cases'], record['excel'])):
            got = PROBE.alone(case, index, Path(folder))
            if got[0] == 'refused':
                continue
            differ += [(case, name, mine, theirs) for name, mine, theirs in zip(names, got[0].split('|'), excel[0].split('|'))
                       if mine not in (theirs, 'refused')]
            if got[1] != excel[1]:
                differ.append((case, 'spPr', got[1], excel[1]))
    assert differ == []


def test_a_picture_from_a_file_is_refused_under_the_sandbox(tmp_path: Path) -> None:
    (tmp_path / 'p.png').write_bytes(PROBE.files()['p.png'])
    for member in ('UserPicture', 'UserTextured'):
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(f'Sub P(ByVal path As String)\n    ActiveSheet.Shapes.AddShape(msoShapeOval, 20, 20, 120, 60).Fill.{member} path\nEnd Sub\n',
                       name='H')
        with pytest.raises(VBASandboxError):
            app.run('H.P', str(tmp_path / 'p.png'))
