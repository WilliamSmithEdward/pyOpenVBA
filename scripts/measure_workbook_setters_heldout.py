"""Held-out setter rounds for twelve Workbook settings: every kind of value, typed and late, the model's answers written before live Excel's.

Each case makes a new workbook, gives it a state first (left as made, or the setting set the other way), sets the setting to one kind of
value, and records the error, then all twelve settings as TypeName:value; then it saves the book, as .xlsx or .xlsm, opens it again and reads
all twelve again. The value kinds are those docs/host_completeness.md asks of a settable property: valid codes and past their bounds, a
fraction, text of a number, text that is no number, empty text, Null, Empty, True, a Currency, a Date and a range of separate areas. The seed
picks each case's first state and file format, so each round's cases differ. No setting here is user-level or reaches outside Excel
(scripts/probe_limits.py).

    python scripts/measure_workbook_setters_heldout.py predict <round>
    python scripts/measure_workbook_setters_heldout.py measure <round>

write tests/fixtures/workbook_setters_heldout/<round>_predictions.json and <round>_excel.json.
"""
from __future__ import annotations

import hashlib
import json
import random
import shutil
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FOLDER = ROOT / 'tests/fixtures/workbook_setters_heldout'
SEPARATOR = '%%'
#: Each setting, and the value that sets it away from a new workbook's.
SETTINGS = {'CheckCompatibility': 'True', 'Date1904': 'True', 'DisplayDrawingObjects': 'xlHide', 'DisplayInkComments': 'False',
            'EnableAutoRecover': 'False', 'ForceFullCalculation': 'True', 'InactiveListBorderVisible': 'False', 'PrecisionAsDisplayed': 'True',
            'SaveLinkValues': 'False', 'ShowPivotChartActiveFields': 'True', 'ShowPivotTableFieldList': 'False', 'UpdateLinks': '2'}
#: Every kind of value a settable property is asked to take.
KINDS = ('True', 'False', '0', '1', '-1', '2', '3', '4', '-4104', '1.5', '"1"', '"x"', '""', 'Null', 'Empty', 'CCur(2)', '#1/2/2020#',
         'ws.Range("A1,C1")', '2147483647', '1E+300')
#: Cases per harness call, each making, saving and opening a workbook.
CHUNK = 20


def cases(round_name: str) -> list[dict[str, Any]]:
    chooser = random.Random(f'workbook-setters-{round_name}')
    return [{'setting': setting, 'value': value, 'first': chooser.choice(('new', 'set')), 'format': chooser.choice(('xlsx', 'xlsm'))}
            for setting in SETTINGS for value in KINDS]


def _reads() -> str:
    return ''.join(f'''    Err.Clear: v = Empty
    v = wb.{name}
    s = s & "/" & IIf(Err.Number, "e" & Err.Number, TypeName(v) & ":" & CStr(v))
''' for name in SETTINGS)


def macro(declared: str, chunk: list[dict[str, Any]]) -> str:
    bodies = ''.join(f'''Private Function Case{index}(ByVal folder As String) As String
    Dim wb As {declared}, ws As Object, v As Variant, s As String, e As Long
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    On Error Resume Next
{f"    wb.{case['setting']} = {SETTINGS[case['setting']]}" + chr(10) if case['first'] == 'set' else ''}    Err.Clear
    wb.{case['setting']} = {case['value']}
    e = Err.Number
    s = "e" & e
{_reads()}    Err.Clear
    wb.SaveAs folder & "\\c{index}.{case['format']}", {'xlOpenXMLWorkbook' if case['format'] == 'xlsx' else 'xlOpenXMLWorkbookMacroEnabled'}
    s = s & "|save e" & Err.Number
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(folder & "\\c{index}.{case['format']}")
{_reads()}    wb.Close SaveChanges:=False
    Case{index} = s
End Function
''' for index, case in enumerate(chunk))
    calls = ''.join(f'    out = out & "{SEPARATOR}" & Case{index}(folder)\n' for index in range(len(chunk)))
    return f'''{bodies}Function Probe(ByVal folder As String) As String
    Dim out As String
    Application.DisplayAlerts = False
{calls}    Application.DisplayAlerts = True
    Probe = Mid(out, {len(SEPARATOR) + 1})
End Function
'''


def _answers(run: Any, round_name: str, answers: dict[str, list[str]] | None = None, keep: Any = None) -> dict[str, list[str]]:
    """Every case's answer by declared type, going on from those already in ``answers``, and handing each chunk's to ``keep``."""
    every = cases(round_name)
    answers = {} if answers is None else answers
    for declared in ('Workbook', 'Object'):
        found = answers.setdefault(declared, [])
        for at in range(len(found), len(every), CHUNK):
            folder = Path(tempfile.mkdtemp(prefix='wbsetters_'))
            try:
                found += run(declared, every[at:at + CHUNK], folder).split(SEPARATOR)
            finally:
                shutil.rmtree(folder, ignore_errors=True)
            if keep is not None:
                keep(answers)
            print(declared, len(found), 'of', len(every), flush=True)
    return answers


def predict(round_name: str) -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication

    from pyopenvba.exceptions import VBAUnsupportedError

    def once(source: str, folder: Path) -> str:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source, name='H')
        return str(app.run('H.Probe', str(folder)))

    def run(declared: str, chunk: list[dict[str, Any]], folder: Path) -> str:
        """A chunk, or, where the model refuses a case openly, each case alone, the refused ones answered as refusals."""
        try:
            return once(macro(declared, chunk), folder)
        except VBAUnsupportedError:
            answers: list[str] = []
            for case in chunk:
                try:
                    answers.append(once(macro(declared, [case]), folder))
                except VBAUnsupportedError as refused:
                    answers.append(f'refused: {refused}')
            return SEPARATOR.join(answers)

    every = cases(round_name)
    text = json.dumps({'round': round_name, 'cases': every, 'answers': _answers(run, round_name)}, indent=1) + '\n'
    FOLDER.mkdir(parents=True, exist_ok=True)
    (FOLDER / f'{round_name}_predictions.json').write_text(text, encoding='utf-8', newline='\n')
    print('predictions', round_name, 'sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure(round_name: str) -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    assert (FOLDER / f'{round_name}_predictions.json').exists(), 'write the predictions first'
    out = FOLDER / f'{round_name}_excel.json'
    record: dict[str, Any] = json.loads(out.read_text(encoding='utf-8')) if out.exists() else {'round': round_name, 'cases': cases(round_name),
                                                                                                'answers': {}}

    def keep(answers: dict[str, list[str]]) -> None:
        record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers})
        out.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')

    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.run_vba('Function Hello() As String\n    Hello = "x"\nEnd Function\n', 'Hello', (), timeout=60)

        def run(declared: str, chunk: list[dict[str, Any]], folder: Path) -> str:
            result = excel.run_vba(macro(declared, chunk), 'Probe', (str(folder),), timeout=280, line_numbers=False)
            assert result.ok, (result.error, [one.message for one in result.dialogs])
            return str(result.value)

        _answers(run, round_name, record['answers'], keep)
    print('measured', round_name)


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]](sys.argv[2])
