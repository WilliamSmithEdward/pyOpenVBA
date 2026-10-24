"""Measure native VBA object lifetime: when Class_Initialize and Class_Terminate run,
As New, and the For Each variable a finished loop leaves.

Each probe is a module whose ``Function Probe() As String`` the fixture records as a
value or a run-time error number; the class modules are added to every session
(see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_object_lifetime.json'
CLASSES = {
    "CItem": '''Public Name As String
Public Child As CItem
Private Sub Class_Initialize()
LogText = LogText & "+"
End Sub
Private Sub Class_Terminate()
LogText = LogText & "-" & Name & ";"
End Sub
''',
}
HEAD = r'''
Public LogText As String
Private Function Make(ByVal n As String) As CItem
Dim it As New CItem
it.Name = n
Set Make = it
End Function
Private Sub UseLocal()
Dim it As New CItem
it.Name = "local"
LogText = LogText & "use;"
End Sub
Private Sub HoldNothing()
Dim it As CItem
LogText = LogText & "hold;"
End Sub
'''
CASES = {
    "set_nothing": 'LogText = ""\nDim p As CItem\nSet p = New CItem\np.Name = "p"\nSet p = Nothing\nProbe = LogText',
    "as_new_lazy": 'LogText = ""\nDim p As New CItem\nLogText = LogText & "before;"\np.Name = "q"\nProbe = LogText',
    "scope_exit": 'LogText = ""\nUseLocal\nLogText = LogText & "after;"\nProbe = LogText',
    "unused_as_new": 'LogText = ""\nHoldNothing\nDim x As New CItem\nLogText = LogText & "end;"\nProbe = LogText',
    "returned_object": 'LogText = ""\nDim a As CItem\nSet a = Make("ret")\nLogText = LogText & "got;"\nSet a = Nothing\nProbe = LogText',
    "returned_discarded": 'LogText = ""\nMake "tmp"\nLogText = LogText & "after;"\nProbe = LogText',
    "returned_member": 'LogText = ""\nDim n As String\nn = Make("m").Name\nLogText = LogText & "after;"\nProbe = LogText & n',
    "reassign": 'LogText = ""\nDim a As CItem\nSet a = New CItem\na.Name = "one"\nSet a = New CItem\na.Name = "two"\nLogText = LogText & "mid;"\nSet a = Nothing\nProbe = LogText',
    "collection_holds": 'LogText = ""\nDim c As New Collection, a As CItem\nSet a = New CItem\na.Name = "c"\nc.Add a\nSet a = Nothing\nLogText = LogText & "held;"\nc.Remove 1\nLogText = LogText & "removed;"\nProbe = LogText',
    "collection_released": 'LogText = ""\nDim c As Collection\nSet c = New Collection\nc.Add Make("x")\nc.Add Make("y")\nLogText = LogText & "full;"\nSet c = Nothing\nProbe = LogText',
    "array_erase": 'LogText = ""\nDim arr(1) As CItem\nSet arr(0) = Make("a0")\nSet arr(1) = Make("a1")\nLogText = LogText & "set;"\nErase arr\nProbe = LogText',
    "array_element_replace": 'LogText = ""\nDim arr(0) As CItem\nSet arr(0) = Make("e1")\nSet arr(0) = Make("e2")\nLogText = LogText & "mid;"\nSet arr(0) = Nothing\nProbe = LogText',
    "variant_holds": 'LogText = ""\nDim v As Variant\nSet v = Make("v")\nLogText = LogText & "v;"\nv = 5\nProbe = LogText',
    "child_chain": 'LogText = ""\nDim p As CItem\nSet p = Make("parent")\nSet p.Child = Make("child")\nLogText = LogText & "built;"\nSet p = Nothing\nProbe = LogText',
    "cycle_never_terminates": 'LogText = ""\nDim a As CItem, b As CItem\nSet a = Make("A")\nSet b = Make("B")\nSet a.Child = b\nSet b.Child = a\nSet a = Nothing\nSet b = Nothing\nLogText = LogText & "dropped;"\nProbe = LogText',
    "as_new_recreate": 'LogText = ""\nDim a As New CItem\na.Name = "first"\nSet a = Nothing\nLogText = LogText & "gap;"\na.Name = "second"\nLogText = LogText & (a Is Nothing) & ";"\nProbe = LogText',
    "as_new_typename": 'LogText = ""\nDim a As New CItem\nDim n As String\nn = TypeName(a)\nLogText = LogText & n & ";"\nProbe = LogText',
    "two_variables_same": 'LogText = ""\nDim a As CItem, b As CItem\nSet a = Make("s")\nSet b = a\nSet a = Nothing\nLogText = LogText & "one;"\nSet b = Nothing\nProbe = LogText',
    "with_block_temp": 'LogText = ""\nWith Make("w")\n.Name = "w2"\nLogText = LogText & "in;"\nEnd With\nLogText = LogText & "out;"\nProbe = LogText',
    "for_each_releases_last": 'LogText = ""\nDim c As New Collection, it As CItem\nc.Add Make("f1")\nc.Add Make("f2")\nFor Each it In c\nNext\nLogText = LogText & "looped;"\nSet c = Nothing\nLogText = LogText & "cnull;"\nSet it = Nothing\nProbe = LogText',
    "for_each_variant_after_loop": 'Dim v As Variant\nFor Each v In Array(1, 2)\nNext\nProbe = TypeName(v) & "|" & IsEmpty(v)',
    "for_each_variant_after_exit": 'Dim v As Variant\nFor Each v In Array(1, 2)\nExit For\nNext\nProbe = TypeName(v) & "|" & v',
    "for_each_object_after_loop": 'Dim c As New Collection, o As Object\nc.Add New Collection\nFor Each o In c\nNext\nProbe = (o Is Nothing)',
    "for_each_object_after_exit": 'Dim c As New Collection, o As Object\nc.Add New Collection\nFor Each o In c\nExit For\nNext\nProbe = (o Is Nothing)',
    "for_each_empty_collection": 'Dim c As New Collection, v As Variant\nv = 5\nFor Each v In c\nNext\nProbe = TypeName(v) & "|" & v',
    "for_each_goto_out": 'Dim v As Variant\nFor Each v In Array(1, 2)\nIf v = 2 Then GoTo done\nNext\ndone:\nProbe = TypeName(v) & "|" & v',
    "set_value_into_object": 'Dim o As Object, v As Variant\nv = 7\nSet o = v\nProbe = "set"',
    "set_value_into_class": 'Dim p As CItem, v As Variant\nv = 7\nSet p = v\nProbe = "set"',
    "set_value_into_variant": 'Dim w As Variant, v As Variant\nv = 7\nSet w = v\nProbe = "set"',
    "set_property_value_into_object": 'Dim o As Object\nSet o = ActiveSheet.Name\nProbe = "set"',
}


def macro(name: str) -> str:
    return HEAD + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT, classes=CLASSES)


if __name__ == '__main__':
    main()
