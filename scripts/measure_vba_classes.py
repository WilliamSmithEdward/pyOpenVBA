"""Class modules in everyday VBA: creation and teardown order, properties, Implements, events through WithEvents, Collection, CallByName.

    python scripts/measure_vba_classes.py

writes tests/fixtures/vba_classes.json, which tests/test_vba_classes.py replays. The classes are in CLASSES and each probe's log is the
module-level string Log, so that Terminate can write to it.
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_classes.json'
CLASSES = {
    'CLog': '''Public Name As String
Private Sub Class_Initialize()
    Note "init"
End Sub
Private Sub Class_Terminate()
    Note "term" & Name
End Sub
Public Function Hello() As String
    Hello = "hi" & Name
End Function
''',
    'CProp': '''Private m As Long
Private arr(2) As Long
Private Const Limit As Long = 5
Public Property Get V() As Long
    V = m
End Property
Public Property Let V(ByVal x As Long)
    If x > Limit Then Err.Raise 1000 + vbObjectError, "CProp", "too big"
    m = x
End Property
Public Property Get Item(ByVal i As Long) As Long
    Item = arr(i)
End Property
Public Property Let Item(ByVal i As Long, ByVal x As Long)
    arr(i) = x
End Property
Public Property Get Self() As CProp
    Set Self = Me
End Property
Public Function Add(ByVal n As Long) As CProp
    m = m + n
    Set Add = Me
End Function
Friend Function Secret() As Long
    Secret = 7
End Function
Public Sub Bump(ByRef n As Long)
    n = n + 1
End Sub
Public Function Fail() As Long
    Fail = 1 / 0
End Function
Private Function Hidden() As Long
    Hidden = 1
End Function
Public Function Counter() As Long
    Static n As Long
    n = n + 1
    Counter = n
End Function
''',
    'IShape': '''Public Function Area() As Double
End Function
Public Property Get Name() As String
End Property
''',
    'CSquare': '''Implements IShape
Public Side As Double
Private Function IShape_Area() As Double
    IShape_Area = Side * Side
End Function
Private Property Get IShape_Name() As String
    IShape_Name = "square"
End Property
Public Function Own() As String
    Own = "own"
End Function
''',
    'CSource': '''Public Event Done(ByVal n As Long, ByRef answer As Long)
Public Event Plain()
Public Sub Fire(ByVal n As Long)
    Dim a As Long
    a = n
    RaiseEvent Done(n, a)
    Note "after" & a
    RaiseEvent Plain
End Sub
''',
    'CSink': '''Public WithEvents S As CSource
Private Sub S_Done(ByVal n As Long, ByRef answer As Long)
    Note "done" & n
    answer = answer * 2
End Sub
Private Sub S_Plain()
    Note "plain"
End Sub
''',
}
#: Stays in the project between probes, so that the classes, which call Note, always compile.
LOGGER = '''Public Log As String
Public Sub Note(ByVal text As String)
    Log = Log & text & ","
End Sub
'''
HELPERS = '''
Private Function Make(ByVal name As String) As CLog
    Set Make = New CLog
    Make.Name = name
End Function
Private Sub Takes(ByVal a As CLog)
    a.Name = "changed"
End Sub
Private Sub TakesRef(ByRef a As CLog)
    Set a = New CLog
    a.Name = "swapped"
End Sub
Private Sub LocalOne()
    Dim a As New CLog
    a.Name = "local"
End Sub
'''
CASES = {
    # Creation and teardown.
    'new_order': 'Dim a As CLog\nNote "before"\nSet a = New CLog\nNote "made"\na.Name = "a"\nSet a = Nothing\nNote "gone"',
    'new_auto_lazy': 'Dim a As New CLog\nNote "declared"\na.Name = "a"\nNote "used"\nSet a = Nothing\nNote "cleared"\ns = (a Is Nothing)\nNote "checked"\na.Name = "b"',
    'new_auto_is_nothing': 'Dim a As New CLog\ns = (a Is Nothing)',
    'replace_terminates': 'Dim a As CLog\nSet a = New CLog\na.Name = "1"\nSet a = New CLog\na.Name = "2"\nNote "end"\nSet a = Nothing',
    'two_refs': 'Dim a As CLog, b As CLog\nSet a = New CLog\na.Name = "x"\nSet b = a\nSet a = Nothing\nNote "one"\nSet b = Nothing\nNote "two"',
    'local_terminates': 'LocalOne\nNote "back"',
    'temporary': 'Note Make("t").Hello\nNote "after"',
    'with_temporary': 'With New CLog\n.Name = "w"\nNote "in"\nEnd With\nNote "out"',
    'collection_holds': 'Dim c As New Collection\nc.Add Make("c")\nNote "added"\nc.Remove 1\nNote "removed"',
    'array_of_objects': 'Dim a(1) As CLog\nSet a(0) = Make("z")\nNote "set"\nErase a\nNote "erased"',
    'variant_holds': 'Dim v As Variant\nSet v = Make("v")\nNote TypeName(v)\nv = 1\nNote "let"',
    'byval_object': 'Dim a As CLog\nSet a = Make("a")\nTakes a\nNote a.Name\nSet a = Nothing',
    'byref_object': 'Dim a As CLog\nSet a = Make("a")\nTakesRef a\nNote a.Name\nSet a = Nothing',
    'nothing_call': 'Dim a As CLog\nNote a.Name',
    'is_same': 'Dim a As CLog, b As CLog\nSet a = New CLog\nSet b = a\ns = (a Is b) & (a Is Nothing)\nSet b = New CLog\ns = s & (a Is b)\nSet a = Nothing\nSet b = Nothing',
    'typename_class': 'Dim a As Object\nSet a = New CProp\ns = TypeName(a) & (TypeOf a Is CProp) & (TypeOf a Is CLog) & (TypeOf a Is Object) & VarType(a) & IsObject(a)',
    # Properties.
    'property_let_get': 'Dim p As New CProp\np.V = 3\ns = p.V',
    'property_args': 'Dim p As New CProp\np.Item(1) = 9\ns = p.Item(1) & p.Item(0)',
    'property_bounds': 'Dim p As New CProp\ns = p.Item(9)',
    'property_raises': 'Dim p As New CProp\nOn Error Resume Next\np.V = 9\ns = (Err.Number - vbObjectError) & Err.Source & Err.Description & p.V',
    'class_const': 'Dim p As New CProp\ns = p.Limit',
    'chain_me': 'Dim p As New CProp\ns = p.Add(1).Add(2).Self.V',
    'friend_call': 'Dim p As New CProp\ns = p.Secret',
    'friend_late': 'Dim p As Object\nSet p = New CProp\ns = p.Secret',
    'private_late': 'Dim p As Object\nSet p = New CProp\ns = p.Hidden',
    'missing_late': 'Dim p As Object\nSet p = New CProp\ns = p.Nope',
    'byref_through_method': 'Dim p As New CProp, n As Long\np.Bump n\np.Bump (n)\ns = n',
    'error_in_method': 'Dim p As New CProp\nOn Error Resume Next\ns = p.Fail\ns = s & Err.Number & Err.Source',
    'static_per_instance': 'Dim p As New CProp, q As New CProp\nn1 = p.Counter\nn1 = p.Counter\ns = n1 & q.Counter',
    'callbyname': 'Dim p As New CProp\nCallByName p, "V", VbLet, 2\ns = CallByName(p, "V", VbGet) & CallByName(p, "Add", VbMethod, 1).V',
    'callbyname_missing': 'Dim p As New CProp\ns = CallByName(p, "Nope", VbGet)',
    'set_property_value': 'Dim p As New CProp\nSet p.Self = p',
    # Implements.
    'implements_call': 'Dim i As IShape, q As New CSquare\nq.Side = 2\nSet i = q\ns = i.Area & i.Name & TypeName(i) & (TypeOf i Is IShape) & (TypeOf i Is CSquare) & (TypeOf q Is IShape) & (i Is q)',
    'implements_own_through_interface_late': 'Dim i As IShape, o As Object\nSet i = New CSquare\nSet o = i\ns = o.Own',
    'implements_interface_late': 'Dim o As Object\nSet o = New CSquare\ns = o.Area',
    'implements_back': 'Dim i As IShape, q As CSquare\nSet i = New CSquare\nSet q = i\nq.Side = 3\ns = i.Area',
    'implements_mismatch': 'Dim i As IShape\nSet i = New CProp',
    'interface_new_itself': 'Dim i As IShape\nSet i = New IShape\ns = i.Area & "|" & i.Name',
    # Events.
    'events_fire': 'Dim k As New CSink\nSet k.S = New CSource\nk.S.Fire 3\ns = Log',
    'events_no_sink': 'Dim c As New CSource\nc.Fire 1',
    'events_unhooked': 'Dim k As New CSink, c As New CSource\nSet k.S = c\nSet k.S = Nothing\nc.Fire 1',
    'events_two_sinks': 'Dim j As New CSink, k As New CSink, c As New CSource\nSet j.S = c\nSet k.S = c\nc.Fire 1',
    # Collection.
    'collection_keys': 'Dim c As New Collection\nc.Add 1, "a"\nc.Add 2, "B"\ns = c("A") & c.Item("b") & c.Count & c(1)',
    'collection_dup_key': 'Dim c As New Collection\nc.Add 1, "a"\nc.Add 2, "A"',
    'collection_missing': 'Dim c As New Collection\ns = c(1)',
    'collection_missing_key': 'Dim c As New Collection\ns = c("x")',
    'collection_before_after': 'Dim c As New Collection, v As Variant\nc.Add "a"\nc.Add "b", , 1\nc.Add "c", , , 1\nFor Each v In c\ns = s & v\nNext',
    'collection_numeric_key': 'Dim c As New Collection\nc.Add "a", "1"\nc.Add "b", CStr(2)\ns = c("1") & c(2)\nc.Add "x", 5',
    'collection_remove_in_loop': 'Dim c As New Collection, v As Variant\nc.Add 1\nc.Add 2\nc.Add 3\nFor Each v In c\ns = s & v\nc.Remove 1\nNext\ns = s & c.Count',
    'collection_assign': 'Dim c As New Collection\nc.Add 1\nc(1) = 2',
    'collection_object_member': 'Dim c As New Collection\nc.Add Make("m")\ns = c(1).Name\nc(1).Name = "n"\ns = s & c(1).Name\nSet c = Nothing',
    'collection_type_key_number': 'Dim c As New Collection\nc.Add 1, Key:="k", Before:=Empty',
    'dictionary_basics': 'Dim d As Object, k As Variant\nSet d = CreateObject("Scripting.Dictionary")\nd("a") = 1\nd.Add "b", 2\nd.Item("c") = 3\nFor Each k In d.Keys\ns = s & k & d(k)\nNext\ns = s & d.Count & d.Exists("a") & d.Exists("z") & UBound(d.Items)\nd.Remove "a"\ns = s & d.Count & TypeName(d("new")) & d.Count\nd.RemoveAll\ns = s & d.Count',
    'dictionary_dup': 'Dim d As Object\nSet d = CreateObject("Scripting.Dictionary")\nd.Add "a", 1\nd.Add "a", 2',
    'dictionary_compare': 'Dim d As Object\nSet d = CreateObject("Scripting.Dictionary")\nd.CompareMode = vbTextCompare\nd("a") = 1\ns = d.Exists("A") & d.Count\nd.CompareMode = vbBinaryCompare',
    'dictionary_key_types': 'Dim d As Object\nSet d = CreateObject("Scripting.Dictionary")\nd(1) = "n"\nd("1") = "s"\nd(1.5) = "f"\ns = d.Count & d(1) & d("1")',
}


def macro(name: str) -> str:
    # n1 is implicit in one probe; each probe ends with the log.
    return HELPERS + "\nFunction Probe() As String\nDim s As String\nLog = \"\"\n" + CASES[name] + "\nProbe = s & \"#\" & Log\nEnd Function\n"


def main() -> None:
    probes: list[dict[str, object]] = []
    excel: ExcelSession | None = None
    try:
        for name in CASES:
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0))
                excel.__enter__()
                excel.new_document()
                first = excel.run_vba('Function Starts() As Long\nEnd Function', 'Starts', line_numbers=False)
                assert first.ok, first
                excel.add_module('Logger', LOGGER)
                for class_name, source in CLASSES.items():
                    excel.add_module(class_name, source, kind='class')
            result = excel.run_vba(macro(name), 'Probe', timeout=30.0, line_numbers=False)
            if any(dialog.classification == 'compile-error' for dialog in result.dialogs):
                probes.append({'name': name, 'macro': macro(name), 'compile': True})
                print(name, 'does not compile', flush=True)
                excel.__exit__(None, None, None)
                excel = None
                continue
            assert not result.dialogs, (name, result)
            if result.ok:
                probe: dict[str, object] = {'name': name, 'macro': macro(name), 'result': result.value}
            else:
                found = re.search(r'VBA error (-?\d+)', str(result.error))
                assert found, (name, result)
                probe = {'name': name, 'macro': macro(name), 'error': int(found.group(1))}
                excel.__exit__(None, None, None)
                excel = None
            probes.append(probe)
            print(name, probe.get('result', probe.get('error')), flush=True)
    finally:
        if excel is not None:
            excel.__exit__(None, None, None)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'logger': LOGGER, 'classes': CLASSES, 'probes': probes}, indent=1) + '\n',
                   encoding='utf-8')


if __name__ == '__main__':
    main()
