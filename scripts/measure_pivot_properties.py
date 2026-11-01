"""Small properties of PivotTable, PivotCache, PivotField and PivotItem on a pivot table over a worksheet range:
what each reads, which an OLAP or external source alone answers, what a set does, and the attributes a save then
writes into the pivot table's and cache's parts, against a save with nothing set.

    python scripts/measure_pivot_properties.py

writes tests/fixtures/pivot_properties.json, which tests/test_excel_pivot_properties.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_properties.json'
#: Each object's properties read on a new table: k a row field, p a page field, v a data field, q in no area.
READS = {
    'pc': ['MissingItemsLimit', 'OptimizeCache', 'RefreshPeriod', 'BackgroundQuery', 'CommandText', 'CommandType',
           'Connection', 'IsConnected', 'LocalConnection', 'MaintainConnection', 'QueryType', 'RobustConnect',
           'SavePassword', 'SourceConnectionFile', 'SourceDataFile', 'UpgradeOnRefresh', 'UseLocalConnection',
           'WorkbookConnection', 'ADOConnection', 'Recordset'],
    'pt': ['SmallGrid', 'PageFieldStyle', 'VacatedStyle', 'PrintDrillIndicators', 'ShowPageMultipleItemLabel',
           'DisplayMemberPropertyTooltips', 'EnableWriteback', 'AllocationMethod', 'AllocationValue', 'Allocation',
           'AllocationWeightExpression', 'TotalsAnnotation', 'ViewCalculatedMembers', 'VisualTotals',
           'VisualTotalsForSets', 'CalculatedMembersInFilters', 'Hidden', 'MDX', 'ChangeList', 'CubeFields',
           'CalculatedMembers', 'PivotChart', 'SelectionMode', 'PivotSelection', 'PivotSelectionStandard',
           'InnerDetail', 'AutoRefresh'],
    'pf': ['AutoShowCount', 'AutoShowField', 'AutoShowRange', 'AutoShowType', 'CurrentPageList', 'CurrentPageName',
           'DatabaseSort', 'DisplayAsCaption', 'DisplayAsTooltip', 'DisplayInReport', 'DragToColumn', 'DragToData',
           'DragToHide', 'DragToPage', 'DragToRow', 'DrilledDown', 'EnableItemSelection', 'Hidden', 'HiddenItemsList',
           'IncludeNewItemsInFilter', 'IsMemberProperty', 'PropertyOrder', 'ServerBased', 'ShowAllItems',
           'SourceCaption', 'UseMemberPropertyAsCaption', 'VisibleItemsList', 'CubeField', 'PropertyParentField',
           'MemberPropertyCaption'],
    'pp': ['CurrentPageName', 'ShowAllItems', 'DragToRow', 'EnableItemSelection', 'DatabaseSort', 'ServerBased'],
    'pq': ['DragToRow', 'ShowAllItems', 'EnableItemSelection', 'DatabaseSort', 'ServerBased'],
    'pd': ['DragToRow', 'ShowAllItems', 'EnableItemSelection', 'DatabaseSort', 'ServerBased', 'AutoShowCount'],
    'pi': ['DrilledDown', 'ParentShowDetail', 'RecordCount', 'SourceNameStandard'],
    'pi2': ['RecordCount', 'SourceNameStandard'],
}
#: Statements, each followed by a read of what it set and a save.
SETS = [
    ['pt.SmallGrid = True'], ['pt.PrintDrillIndicators = True'], ['pt.ShowPageMultipleItemLabel = False'],
    ['pt.DisplayMemberPropertyTooltips = False'], ['pt.TotalsAnnotation = True'], ['pt.VisualTotals = False'],
    ['pt.VisualTotalsForSets = True'], ['pt.CalculatedMembersInFilters = True'],
    ['pt.PageFieldStyle = "PivotStyleDark1"'], ['pt.VacatedStyle = "Good"'], ['pt.EnableWriteback = True'],
    ['pt.AutoRefresh = True'], ['pt.Hidden = True'], ['pc.MissingItemsLimit = xlMissingItemsNone'],
    ['pc.OptimizeCache = True'], ['pc.BackgroundQuery = True'], ['pc.SavePassword = True'],
    ['pc.UpgradeOnRefresh = True'], ['pc.RefreshPeriod = 5'], ['pf.DragToColumn = False'], ['pf.DragToRow = False'],
    ['pf.DragToPage = False'], ['pf.DragToData = False'], ['pf.DragToHide = False'], ['pf.ShowAllItems = False'],
    ['pf.IncludeNewItemsInFilter = True'], ['pf.EnableItemSelection = False'], ['pf.DatabaseSort = True'],
    ['pf.AutoShowCount = 3'], ['pf.ServerBased = True'], ['pf.SourceCaption = "zz"'],
    ['pt.SmallGrid = True', 'pt.SmallGrid = False'], ['pt.PrintDrillIndicators = True', 'pt.PrintDrillIndicators = False'],
    ['pt.ShowPageMultipleItemLabel = False', 'pt.ShowPageMultipleItemLabel = True'],
    ['pt.PageFieldStyle = "PivotStyleDark1"', 'pt.PageFieldStyle = ""'], ['pt.VacatedStyle = "Good"', 'pt.VacatedStyle = ""'],
    ['pc.MissingItemsLimit = xlMissingItemsNone', 'pc.MissingItemsLimit = xlMissingItemsDefault'],
    ['pc.MissingItemsLimit = xlMissingItemsMax2'], ['pc.MissingItemsLimit = 5'], ['pc.MissingItemsLimit = -2'],
    ['pf.DragToColumn = False', 'pf.DragToColumn = True'], ['pf.DragToHide = False', 'pf.DragToHide = True'],
    ['pf.IncludeNewItemsInFilter = True', 'pf.IncludeNewItemsInFilter = False'],
    ['pf.EnableItemSelection = False', 'pf.EnableItemSelection = True'],
    ['pt.PageFieldStyle = "nonsense"'], ['pt.SmallGrid = "x"'], ['pf.DragToRow = Null'], ['pq.DragToColumn = False'],
    ['pd.DragToColumn = False'], ['pp.ShowAllItems = True'], ['pp.DragToRow = False'], ['pq.ShowAllItems = True'],
]


def _probe() -> str:
    lines = ['Private Function GV(ByVal o As Object, ByVal n As String) As String', '    Dim v As Variant',
             '    On Error GoTo Failed', '    v = CallByName(o, n, VbGet)', '    If IsArray(v) Then',
             '        GV = "arr"', '    ElseIf IsNull(v) Then', '        GV = "Null"', '    Else',
             '        GV = TypeName(v) & ":" & CStr(v)', '    End If', '    Exit Function', 'Failed:',
             '    GV = "E" & Err.Number', 'End Function', '',
             'Private Function GO(ByVal o As Object, ByVal n As String) As String', '    Dim x As Object',
             '    On Error GoTo Failed', '    Set x = CallByName(o, n, VbGet)',
             '    If x Is Nothing Then GO = "Nothing" Else GO = "obj:" & TypeName(x)', '    Exit Function', 'Failed:',
             '    GO = "E" & Err.Number', 'End Function', '',
             'Private Function G(ByVal o As Object, ByVal n As String) As String', '    G = GV(o, n)',
             '    If G = "E450" Or G = "E438" Or G = "E13" Or G = "E91" Then G = G & "/" & GO(o, n)', 'End Function',
             '',
             'Private Function Made(ByVal wb As Workbook) As Object', '    Dim ws As Worksheet, pt As Object',
             '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:D5").Value = Application.Transpose(Application.Transpose(Array(Array("k", "p", "v", "q"), '
             'Array("a", "x", 1, "m"), Array("b", "y", 2, "n"), Array("a", "y", 3, "m"), Array("c", "x", 4, "n"))))',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:D5")).CreatePivotTable(ws.Range("I3"), "PT")',
             '    pt.PivotFields("k").Orientation = xlRowField', '    pt.PivotFields("p").Orientation = xlPageField',
             '    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum', '    Set Made = pt', 'End Function', '',
             'Public Function Reads() As String',
             '    Dim wb As Workbook, pt As Object, s As String, n As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set pt = Made(wb)']
    targets = {'pc': 'pt.PivotCache', 'pt': 'pt', 'pf': 'pt.PivotFields("k")', 'pp': 'pt.PivotFields("p")',
               'pq': 'pt.PivotFields("q")', 'pd': 'pt.DataFields(1)', 'pi': 'pt.PivotFields("k").PivotItems(1)',
               'pi2': 'pt.PivotFields("k").PivotItems("b")'}
    for key, names in READS.items():
        listed = ', '.join(f'"{name}"' for name in names)
        lines += [f'    For Each n In Array({listed})', f'        s = s & "~" & G({targets[key]}, CStr(n))', '    Next']
    lines += ['    wb.Close False', '    Reads = Mid(s, 2)', 'End Function', '',
              'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, pt As Object, pc As Object, pf As Object, pp As Object, pq As Object, pd As Object',
              '    Dim v As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set pt = Made(wb)',
              '    Set pc = pt.PivotCache', '    Set pf = pt.PivotFields("k")', '    Set pp = pt.PivotFields("p")',
              '    Set pq = pt.PivotFields("q")', '    Set pd = pt.DataFields(1)', '    On Error GoTo Failed',
              '    Select Case which']
    for index, statements in enumerate(SETS):
        target, name = statements[-1].split(' = ')[0].split('.')
        lines += [f'    Case {index}', *(f'        {statement}' for statement in statements),
                  f'        v = {target}.{name}']
    lines += [f'    Case {len(SETS)}', '        v = "plain"', '    End Select',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    Tried = "ok " & CStr(v)',
              'Done:', '    wb.Close False', '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number',
              '    Resume Done', 'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
              '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', '    s = Reads()',
              f'    For i = 0 To {len(SETS)}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = s', 'End Function']
    return '\n'.join(lines) + '\n'


def attributes(path: Path) -> dict[str, dict[str, str]]:
    """The pivot table's and cache's root attributes and each pivotField's, by part, revision ids left out."""
    found: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith(('xl/pivotTables/pivotTable', 'xl/pivotCache/pivotCacheDefinition')):
                text = package.read(name).decode('utf-8')
                root = re.search(r'<(?:pivotCacheDefinition|pivotTableDefinition)\b[^>]*>', text)
                kept = dict(re.findall(r'([\w:]+)="([^"]*)"', root.group(0) if root else ''))
                for index, field in enumerate(re.findall(r'<pivotField\b[^>]*>', text)):
                    kept.update({f'field{index}.{key}': value
                                 for key, value in re.findall(r'([\w:]+)="([^"]*)"', field)})
                found[name.rsplit('/', 1)[-1]] = {key: value for key, value in kept.items()
                                                  if key not in ('xr:uid', 'refreshedDate', 'refreshedBy', 'cacheId',
                                                                 'r:id') and not key.startswith('xmlns')}
    return found


def changes(path: Path, plain: dict[str, dict[str, str]]) -> dict[str, dict[str, list[str | None]]]:
    found = {}
    for part, now in attributes(path).items():
        before = plain.get(part, {})
        differ = {key: [before.get(key), now.get(key)] for key in sorted(set(before) | set(now))
                  if before.get(key) != now.get(key)}
        if differ:
            found[part] = differ
    return found


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=900.0)
        assert result.ok and not result.dialogs, result
        reads, *answers = str(result.value).split('@@@')
        plain = attributes(Path(folder) / f'c{len(SETS)}.xlsx')
        saved = [changes(Path(folder) / f'c{index}.xlsx', plain) if (Path(folder) / f'c{index}.xlsx').exists() else None
                 for index in range(len(SETS))]
    names = [[key, name] for key, listed in READS.items() for name in listed]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'reads': [[*pair, answer] for pair, answer in zip(names, reads.split('~'), strict=True)],
              'sets': [[statements, answer, change]
                       for statements, answer, change in zip(SETS, answers[:len(SETS)], saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for key, name, answer in record['reads']:
        print(key, name, answer)
    for statements, answer, change in record['sets']:
        print(statements, '->', answer, '|', change)


if __name__ == '__main__':
    main()
