"""Held-out rounds for the Workbook reads: books chosen from a seed, the model's answers written down before live Excel's.

Each round makes a few books, each from the seed: how many sheets, which workbook settings it sets and to what, a title and custom
properties, whether it is saved as .xlsx or .xlsm, and whether it is opened again read-only. A book is built by a macro, saved to a temporary
folder, closed, opened again, and read by scripts/measure_workbook_reads.py's probe, typed (As Workbook) and late (As Object). The settings
leave out mail, the cloud, rights and encryption (scripts/probe_limits.py, tests/test_probes_leave_crypto_alone.py), and AccuracyVersion,
whose set the model refuses openly.

    python scripts/measure_workbook_reads_heldout.py predict <round>    the model's answers, in an empty temporary folder; no Excel
    python scripts/measure_workbook_reads_heldout.py measure <round>    live Excel's, once the predictions and their digest are written

write tests/fixtures/workbook_reads_heldout/<round>_predictions.json and <round>_excel.json.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import random
import shutil
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FOLDER = ROOT / 'tests/fixtures/workbook_reads_heldout'
_SPEC = importlib.util.spec_from_file_location('measure_workbook_reads', ROOT / 'scripts/measure_workbook_reads.py')
assert _SPEC is not None and _SPEC.loader is not None
READS: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(READS)
BOOKS = 4
#: Workbook settings a round may set, each with the values it chooses among.
SETTINGS: dict[str, tuple[str, ...]] = {
    'Date1904': ('True', 'False'), 'CheckCompatibility': ('True', 'False'), 'DisplayDrawingObjects': ('xlDisplayShapes', 'xlPlaceholders', 'xlHide'),
    'InactiveListBorderVisible': ('True', 'False'), 'UpdateRemoteReferences': ('True', 'False'), 'SaveLinkValues': ('True', 'False'),
    'ForceFullCalculation': ('True', 'False'), 'ShowPivotTableFieldList': ('True', 'False'), 'ShowPivotChartActiveFields': ('True', 'False'),
    'EnableAutoRecover': ('True', 'False'), 'RemovePersonalInformation': ('True', 'False'), 'PrecisionAsDisplayed': ('True', 'False'),
    'DisplayInkComments': ('True', 'False'), 'UpdateLinks': ('1', '2', '3'), 'ConflictResolution': ('1', '2', '3'),
    'DoNotPromptForConvert': ('True', 'False'), 'PersonalViewListSettings': ('True', 'False'), 'PersonalViewPrintSettings': ('True', 'False'),
    'TemplateRemoveExtData': ('True', 'False'), 'KeepChangeHistory': ('True', 'False'), 'HighlightChangesOnScreen': ('True', 'False'),
    'ListChangesOnNewSheet': ('True', 'False'), 'DefaultTableStyle': ('"TableStyleLight9"', '"TableStyleMedium2"', '"TableStyleDark1"'),
    'DefaultPivotTableStyle': ('"PivotStyleLight16"', '"PivotStyleMedium9"'),
}
WORDS = ('North', 'Ledger', 'Quarter', 'Draft', 'Summary', 'Budget', 'Plan', 'Review')


def books(round_name: str) -> list[dict[str, Any]]:
    """The round's books, chosen from its name as the seed."""
    chooser = random.Random(f'workbook-reads-{round_name}')
    made: list[dict[str, Any]] = []
    for _ in range(BOOKS):
        chosen = chooser.sample(sorted(SETTINGS), chooser.randint(4, 10))
        made.append({'sheets': chooser.randint(1, 4), 'settings': {name: chooser.choice(SETTINGS[name]) for name in chosen},
                     'title': ' '.join(chooser.sample(WORDS, 2)), 'custom': chooser.randint(0, 3),
                     'format': chooser.choice(('xlsx', 'xlsm')), 'read_only': chooser.random() < 0.5})
    return made


def build(book: dict[str, Any]) -> str:
    """The macro that makes one book and saves it to the path it is given; it answers the settings it could not set, as name=error."""
    sets = ''.join(f'''    Err.Clear
    wb.{name} = {value}
    If Err.Number <> 0 Then failed = failed & "{name}=" & Err.Number & ";"
''' for name, value in book['settings'].items())
    custom = ''.join(f'''    wb.CustomDocumentProperties.Add Name:="Item{index}", LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:={index * 7}
''' for index in range(1, book['custom'] + 1))
    file_format = 'xlOpenXMLWorkbook' if book['format'] == 'xlsx' else 'xlOpenXMLWorkbookMacroEnabled'
    return f'''Function Build(ByVal path As String) As String
    Dim wb As Workbook, i As Long, failed As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    For i = 2 To {book['sheets']}
        wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    Next
    wb.Worksheets(1).Range("A1").Value = 1.25
    wb.Names.Add Name:="Base", RefersTo:="=" & wb.Worksheets(1).Name & "!$A$1"
    wb.BuiltinDocumentProperties("Title").Value = "{book['title']}"
{custom}    On Error Resume Next
{sets}    On Error GoTo 0
    wb.SaveAs path, {file_format}
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Build = failed
End Function
'''


def _answers(run_build: Any, run_probe: Any, folder: Path, round_name: str) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for index, book in enumerate(books(round_name)):
        path = folder / f'Book{index}.{book["format"]}'
        failed = run_build(build(book), path)
        out.append({'book': book, 'failed_sets': failed,
                    'answers': {declared: run_probe(READS.macro(declared), path, book['read_only']).split(READS.SEPARATOR)
                                for declared in ('Workbook', 'Object')}})
    return out


def predict(round_name: str) -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication

    def run_build(source: str, path: Path) -> str:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source, name='B')
        return str(app.run('B.Build', str(path)))

    def run_probe(source: str, path: Path, read_only: bool) -> str:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source, name='H')
        return str(app.run('H.Probe', str(path), read_only))

    folder = Path(tempfile.mkdtemp(prefix='wbheld_'))
    try:
        record = {'round': round_name, 'names': READS.NAMES, 'books': _answers(run_build, run_probe, folder, round_name)}
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    text = json.dumps(record, indent=1) + '\n'
    FOLDER.mkdir(parents=True, exist_ok=True)
    (FOLDER / f'{round_name}_predictions.json').write_text(text, encoding='utf-8', newline='\n')
    print('predictions', round_name, 'sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure(round_name: str) -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    assert (FOLDER / f'{round_name}_predictions.json').exists(), 'write the predictions first'
    folder = Path(tempfile.mkdtemp(prefix='wbheld_'))
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            excel.run_vba('Function Hello() As String\n    Hello = "x"\nEnd Function\n', 'Hello', (), timeout=60)

            def run_build(source: str, path: Path) -> str:
                result = excel.run_vba(source, 'Build', (str(path),), timeout=120, line_numbers=False)
                assert result.ok, (result.error, [one.message for one in result.dialogs])
                return str(result.value)

            def run_probe(source: str, path: Path, read_only: bool) -> str:
                result = excel.run_vba(source, 'Probe', (str(path), read_only), timeout=300, line_numbers=False)
                assert result.ok, (result.error, [one.message for one in result.dialogs])
                return str(result.value)

            record = {'round': round_name, 'measured_at': datetime.now(timezone.utc).isoformat(), 'names': READS.NAMES,
                      'books': _answers(run_build, run_probe, folder, round_name)}
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    (FOLDER / f'{round_name}_excel.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    print('measured', round_name)


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]](sys.argv[2])
