"""Measure the macro metadata of a defined name: MacroType, ShortcutKey, Category and CategoryLocal, and
WorkbookParameter and ValidWorkbookParameter, read and set on Name and given to Names.Add.

Each case runs on a new workbook and answers each property as text; some save it, and workbook.xml's definedNames
are recorded.

    python scripts/measure_name_macro.py

writes tests/fixtures/name_macro.json, which tests/test_excel_name_macro.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/name_macro.json'
PROPS = ['MacroType', 'ShortcutKey', 'Category', 'CategoryLocal', 'WorkbookParameter', 'ValidWorkbookParameter']
PLAIN = 'Set n = wb.Names.Add("Mac", "=Sheet1!$A$1")'
#: Each case: statements on a new workbook wb with sheet ws and a name n (late bound); the file saved, if any.
CASES: list[tuple[list[str], str]] = [
    ([PLAIN], 'plain.xlsx'),
    ([PLAIN, 'n.MacroType = xlFunction'], 'function.xlsx'),
    ([PLAIN, 'n.MacroType = xlCommand'], 'command.xlsx'),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.MacroType = xlNotXLM'], 'back.xlsx'),
    ([PLAIN, 'n.MacroType = 0'], ''),
    ([PLAIN, 'n.MacroType = 4'], ''),
    ([PLAIN, 'n.MacroType = "2"'], ''),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.ShortcutKey = "k"'], 'key.xlsx'),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.ShortcutKey = "K"'], 'shift_key.xlsx'),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.ShortcutKey = "ab"'], ''),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.ShortcutKey = "1"'], ''),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.ShortcutKey = "k"', 'n.ShortcutKey = ""'], 'key_cleared.xlsx'),
    ([PLAIN, 'n.ShortcutKey = "k"'], ''),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.ShortcutKey = "k"'], ''),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "Financial"'], 'financial.xlsx'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "text"'], 'text.xlsx'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "My Things"'], 'custom.xlsx'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.CategoryLocal = "Statistical"'], 'local.xlsx'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = ""'], ''),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.Category = "Database"'], 'command_category.xlsx'),
    ([PLAIN, 'n.Category = "Financial"'], 'plain_category.xlsx'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "Financial"', 'n.MacroType = xlNotXLM'], 'dropped.xlsx'),
    (['Set n = wb.Names.Add(Name:="Mac", RefersTo:="=Sheet1!$A$1", MacroType:=xlCommand, ShortcutKey:="q")'],
     'add_key.xlsx'),
    (['Set n = wb.Names.Add(Name:="Mac", RefersTo:="=Sheet1!$A$1", MacroType:=xlFunction, Category:="Text")'],
     'add_category.xlsx'),
    (['Set n = wb.Names.Add(Name:="Mac", RefersTo:="=Sheet1!$A$1", ShortcutKey:="q")'], ''),
    (['Set n = wb.Names.Add(Name:="Mac", RefersTo:="=Sheet1!$A$1", Category:="Text")'], ''),
    (['Set n = wb.Names.Add(Name:="Mac", RefersTo:="=Sheet1!$A$1", MacroType:=xlFunction, CategoryLocal:="Logical")'],
     ''),
    (['Set n = wb.Names.Add(Name:="Mac", RefersTo:="=Sheet1!$A$1", MacroType:=7)'], ''),
    ([PLAIN, 'n.WorkbookParameter = True'], 'parameter.xlsx'),
    (['Set n = wb.Names.Add("Mac", "=Sheet1!$A$1:$A$2")', 'n.WorkbookParameter = True'], ''),
    (['Set n = wb.Names.Add("Mac", "=5")', 'n.WorkbookParameter = True'], ''),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.WorkbookParameter = True'], ''),
    ([PLAIN, 'n.WorkbookParameter = True', 'n.WorkbookParameter = False'], ''),
    (['Set n = wb.Names.Add("Mac", "=Sheet1!$A$1:$A$2")'], ''),
    (['Set n = wb.Names.Add("Mac", "=5")'], ''),
    (['Set n = ws.Names.Add("Mac", "=Sheet1!$A$1")', 'n.WorkbookParameter = True'], ''),
    (['Set n = wb.Names.Add("Mac", "=Sheet1!$A$1", False)', 'n.MacroType = xlCommand', 'n.ShortcutKey = "k"'],
     'hidden_key.xlsx'),
    ([PLAIN, 'n.MacroType = xlFunction'], 'function.xlsm'),
    ([PLAIN, 'n.MacroType = xlCommand'], 'command.xlsm'),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.ShortcutKey = "K"'], 'shift_key.xlsm'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "Financial"'], 'financial.xlsm'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "Text"'], 'text.xlsm'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "User Defined"'], 'user.xlsm'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "My Things"'], 'custom.xlsm'),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.Category = "Database"'], 'command_category.xlsm'),
    (['Set n = wb.Names.Add("Mac", "=Sheet1!$A$1", False)', 'n.MacroType = xlCommand', 'n.ShortcutKey = "k"'],
     'hidden_key.xlsm'),
    ([PLAIN, 'wb.Names.Add "Other", "=Sheet1!$B$1"', 'n.MacroType = xlCommand'], 'two.xlsx'),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "My Things"', 'Set n = wb.Names.Add("Fn", "=Sheet1!$A$2")',
      'n.MacroType = xlFunction', 'n.Category = "Zed"'], 'two_custom.xlsm'),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.ShortcutKey = "k"', 'n.RefersTo = "=Sheet1!$B$2"'], ''),
    ([PLAIN, 'n.MacroType = xlCommand', 'n.ShortcutKey = "k"', 'Set n = wb.Names.Add("Mac", "=Sheet1!$B$2")'], ''),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "Financial"', 'n.ShortcutKey = "k"',
      'wb.SaveAs folder & "\\reopened.xlsm", xlOpenXMLWorkbookMacroEnabled', 'wb.Close False',
      'Set wb = Workbooks.Open(folder & "\\reopened.xlsm")', 'Set n = wb.Names("Mac")'], ''),
    ([PLAIN, 'n.MacroType = xlFunction', 'n.Category = "My Things"',
      'wb.SaveAs folder & "\\reopened_custom.xlsm", xlOpenXMLWorkbookMacroEnabled', 'wb.Close False',
      'Set wb = Workbooks.Open(folder & "\\reopened_custom.xlsm")', 'Set n = wb.Names("Mac")'], ''),
]
for refers, owner in (('=Sheet1!$A$1:$A$2', 'wb'), ('=Sheet1!A1', 'wb'), ('=5', 'wb'), ('=Sheet1!$A$1', 'ws')):
    CASES.append(([f'Set n = {owner}.Names.Add("Mac", "{refers}")'], f'saved{len(CASES)}.xlsx'))
#: Excel's function categories, each given to a function name and saved, for the functionGroupId it writes.
CATEGORIES = ['Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database', 'Text',
              'Logical', 'Information', 'Commands', 'Customizing', 'Macro Control', 'DDE/External', 'User Defined',
              'Engineering', 'Cube', 'Compatibility', 'Web', 'Date and Time', 'lookup & reference']
for at, category in enumerate(CATEGORIES):
    CASES.append(([PLAIN, 'n.MacroType = xlFunction', f'n.Category = "{category}"'], f'category{at}.xlsm'))


def _probe() -> str:
    lines = ['Private Function Prop(ByVal n As Object, ByVal what As String) As String',
             '    Dim v As Variant', '    On Error Resume Next', '    v = CallByName(n, what, VbGet)',
             '    If Err.Number <> 0 Then Prop = "E" & Err.Number Else Prop = TypeName(v) & ":" & CStr(v)',
             'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, n As Object, p As Variant, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "Sheet1"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        if saved:
            kind = 'xlOpenXMLWorkbookMacroEnabled' if saved.endswith('.xlsm') else 'xlOpenXMLWorkbook'
            lines.append(f'        wb.SaveAs folder & "\\{saved}", {kind}')
    props = ', '.join(f'"{name}"' for name in PROPS)
    lines += ['    End Select', f'    For Each p In Array({props})', '        s = s & "|" & Prop(n, CStr(p))',
              '    Next', '    Tried = Mid(s, 2)', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _names(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    return (re.search(r'<definedNames>.*?</definedNames>', workbook, re.DOTALL) or [''])[0]


def _run(source: str, folder: str, first: int, last: int) -> list[str] | None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=300.0)
    if result.outcome == 'runner-error':
        return None
    assert result.ok and not result.dialogs, result
    return str(result.value).split('@@@')


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        found = _run(source, folder, 0, len(CASES) - 1)
        if found is None:
            found = [(_run(source, folder, index, index) or ['Excel crashed'])[0] for index in range(len(CASES))]
        names = {saved: _names(Path(folder) / saved) for _, saved in CASES
                 if saved and (Path(folder) / saved).exists()}
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, saved, answer] for (statements, saved), answer in zip(CASES, found, strict=True)],
              'names': names}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, saved, answer in record['cases']:
        print(statements[-1][:60], '->', answer, '|', names.get(saved, ''))


if __name__ == '__main__':
    main()
