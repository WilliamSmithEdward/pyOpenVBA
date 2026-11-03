"""Count and Item on every collection scripts/measure_object_trio.py reaches: Count's value and type, and Item with no
argument and with each kind of index a macro can pass, late bound: first, last, past either end, a fraction, text of a
number, a name in its own case and another, a name it lacks, empty text, Empty, Null, True and an object. The sheet the
probe builds is the populated state, and the second sheet's collections the empty one.

    python scripts/measure_collection_items.py

writes tests/fixtures/collection_items.json, which tests/test_excel_collection_items.py replays. A collection that
ends the run is recorded as KILLS, and left out of the test.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

import measure_object_trio as trio
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/collection_items.json'
#: Axes.Item takes a type and a group, and Range's is a cell, measured elsewhere.
NOT_BY_INDEX = {'Axes', 'Range', 'Application', 'WorksheetFunction'}
#: The same collections with nothing in them.
EMPTY = [('Shapes', 'ws2.Shapes'), ('ChartObjects', 'ws2.ChartObjects'), ('Comments', 'ws2.Comments'),
         ('CommentsThreaded', 'ws2.CommentsThreaded'), ('ListObjects', 'ws2.ListObjects'),
         ('Hyperlinks', 'ws2.Hyperlinks'), ('Scenarios', 'ws2.Scenarios'), ('HPageBreaks', 'ws2.HPageBreaks'),
         ('VPageBreaks', 'ws2.VPageBreaks'), ('CustomProperties', 'ws2.CustomProperties'),
         ('PivotTables', 'ws.PivotTables'), ('Names', 'ws2.Names'), ('FormatConditions', 'ws2.Range("A1").FormatConditions'),
         ('SparklineGroups', 'ws2.Range("A1").SparklineGroups'), ('SortFields', 'ws2.Sort.SortFields'),
         ('Errors', 'ws2.Range("A1").Errors')]
#: What each case passes, by name. "named" and "cased" use the first item's own Name; count is o.Count.
INDEXES = {'first': '1', 'last': 'o.Count', 'past': 'o.Count + 1', 'zero': '0', 'negative': '-1', 'fraction': '1.5',
           'text': '"1"', 'named': 'nm', 'cased': 'UCase(nm)', 'lacking': '"no such name"', 'blank': '""',
           'empty': 'Empty', 'null': 'Null', 'true': 'True', 'object': 'Application'}

_HELPERS = '''
Private Function ItemNamed(o As Object) As String
    Dim q As Object
    On Error Resume Next
    Set q = o.Item(1)
    ItemNamed = q.Name
End Function

Private Function Shown(o As Object, ByVal index As Variant) As String
    Dim q As Object, v As Variant, n As String
    On Error Resume Next
    Set q = o.Item(index)
    If Err.Number = 0 Then
        n = q.Name
        Err.Clear
        Shown = TypeName(q) & ":" & n
        Exit Function
    End If
    If Err.Number <> 424 Then
        Shown = "e" & Err.Number
        Exit Function
    End If
    Err.Clear
    v = o.Item(index)
    If Err.Number <> 0 Then
        Shown = "e" & Err.Number
    Else
        Shown = TypeName(v) & " " & CStr(v)
    End If
End Function
'''


def collections() -> list[tuple[str, str]]:
    """Each (class, way in) whose class has both Count and Item on the populated sheet, then the empty ones."""
    named = sorted({name for name, _ in trio.REACHES})
    return [(name, reach) for name, reach in trio.REACHES if name in COUNTED and name in named] + EMPTY


COUNTED = {'Areas', 'CalculatedFields', 'CalculatedItems', 'ChartGroups', 'ChartObjects', 'ColorScaleCriteria',
           'ColorStops', 'Comments', 'CommentsThreaded', 'Connections', 'CustomProperties', 'DataLabels',
           'DocumentProperties', 'Errors', 'Filters', 'FormatConditions', 'FullSeriesCollection', 'GroupShapes',
           'HPageBreaks', 'Hyperlinks', 'IconCriteria', 'IconSet', 'IconSets', 'LegendEntries', 'ListColumns',
           'ListObjects', 'ListRows', 'Names', 'Panes', 'PivotCaches', 'PivotFields', 'PivotFilters', 'PivotItemList',
           'PivotItems', 'PivotTables', 'Points', 'Queries', 'Scenarios', 'SeriesCollection', 'ShapeRange', 'Shapes',
           'SheetViews', 'Sheets', 'Slicers', 'SortFields', 'SparklineGroup', 'SparklineGroups', 'Styles',
           'TableStyles', 'Trendlines', 'VPageBreaks', 'Windows', 'Workbooks'}


def case_function(number: int, reach: str, case: str) -> str:
    """One case as a function, so the model can run each alone and report a gap for just that one."""
    head = (f'\nPublic Function C{number}() As String\n    Dim o As Object, nm As String, v As Variant\n'
            f'    On Error Resume Next\n    Set o = {reach}\n    If Err.Number <> 0 Then\n'
            f'        C{number} = "reach e" & Err.Number\n        Exit Function\n    End If\n')
    if case == 'count':
        body = (f'    v = o.Count\n    If Err.Number <> 0 Then\n        C{number} = "e" & Err.Number\n    Else\n'
                f'        C{number} = TypeName(v) & " " & v\n    End If\n')
    elif case == 'missing':
        body = f'    v = o.Item\n    C{number} = "e" & Err.Number\n'
    else:
        index = INDEXES[case]
        name = '    nm = ItemNamed(o)\n' if case in ('named', 'cased') else ''
        guard = (f'    If nm = "" Then\n        C{number} = "no name"\n        Exit Function\n    End If\n'
                 if name else '')
        body = f'{name}{guard}    C{number} = Shown(o, {index})\n'
    return f'{head}{body}End Function\n'


def all_cases() -> list[tuple[str, str, str]]:
    return [(name, reach, case) for name, reach in collections() for case in ('count', 'missing', *INDEXES)]


def module(chosen: list[int], cases: list[tuple[str, str, str]], log: str = '') -> str:
    """The setup, the helpers, one function per case and a Probe calling those chosen."""
    source = trio.source()
    body = source[:source.index('Private Function Trio')] + _HELPERS + source[source.index('Public Sub Setup'):]
    body = body[:body.index('Public Function Reached')]
    for number in chosen:
        _, reach, case = cases[number]
        body += case_function(number, reach, case)
    calls = ''.join((f'    Open "{log}" For Append As #1: Print #1, {number}: Close #1\n' if log else '')
                    + f'    out = out & "@@" & C{number}()\n' for number in chosen)
    return (body + 'Public Function BookName() As String\n    BookName = wb.Name\nEnd Function\n'
            + 'Public Function Probe() As String\n    Dim out As String\n    Setup\n' + calls
            + '    Probe = Replace(Mid(out, 3), wb.Name, "<book>")\n    wb.Close False\nEnd Function\n')


def main() -> None:
    cases = all_cases()
    pending = list(range(len(cases)))
    answers: dict[int, str] = {}
    excel = None
    with tempfile.TemporaryDirectory() as folder:
        log = str(Path(folder) / 'progress.txt')
        while pending:
            chosen = pending[:80]
            Path(log).unlink(missing_ok=True)
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False))
                excel.__enter__()
                excel.new_document()
            result = excel.run_vba(module(chosen, cases, log), 'Probe', (), timeout=120.0)
            if result.ok and not result.dialogs:
                answers.update(zip(chosen, str(result.value).split('@@'), strict=True))
                pending = pending[len(chosen):]
                print(len(pending), 'left', flush=True)
                continue
            seen = Path(log).read_text().split() if Path(log).exists() else [str(chosen[0])]
            stopped = int(seen[-1])
            print('stopped at', cases[stopped], result.outcome, flush=True)
            answers[stopped] = 'KILLS'
            pending.remove(stopped)
            excel.__exit__(None, None, None)
            excel = None
    if excel is not None:
        excel.__exit__(None, None, None)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(),
                               'cases': [[*cases[number], answers[number]] for number in range(len(cases))]},
                              indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
