"""Measure custom VBA event subscription and shared ByRef argument behavior."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

DESTINATION = Path(__file__).resolve().parents[1] / "tests/fixtures/vba_semantics/withevents.json"

SOURCE = '''Public Event Updated(ByRef Value As Long)
Public Event Observed(ByVal Value As Long)
Public Function Fire(ByVal n As Long) As Long
    RaiseEvent Updated(n)
    Fire = n
End Function
Public Function FireValue(ByVal n As Long) As Long
    RaiseEvent Observed(n)
    FireValue = n
End Function
'''
SINK = '''Public WithEvents Source As EventSource, OtherSource As EventSource
Public Label As String
Public Increment As Long
Public Disconnect As Boolean
Public Peer As EventSink
Public DisconnectPeer As Boolean
Public RebindPeer As Boolean
Public Reenter As Boolean
Private Sub Source_Updated(ByRef Value As Long)
    Log = Log & Label & ":" & CStr(Value) & "|"
    Value = Value + Increment
    If Disconnect Then Set Source = Nothing
    If DisconnectPeer Then Set Peer.Source = Nothing
    If RebindPeer Then Set Peer.Source = Source
    If Reenter Then
        Reenter = False
        Value = Source.Fire(Value)
    End If
End Sub
Private Sub Source_Observed(ByVal Value As Long)
    Log = Log & Label & ":" & CStr(Value) & "|"
    Value = Value + Increment
End Sub
Private Sub OtherSource_Updated(ByRef Value As Long)
    Log = Log & "other:" & CStr(Value) & "|"
    Value = Value + 1000
End Sub
'''
SETUP = '''Dim source As New EventSource, other As New EventSource
Dim a As New EventSink, b As New EventSink
a.Label = "A": a.Increment = 10
b.Label = "B": b.Increment = 100
Log = ""
'''
ACTIONS = {
    "no_listener": 'Probe = CStr(source.Fire(1))',
    "one_listener": 'Set a.Source = source\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "two_listeners": 'Set a.Source = source\nSet b.Source = source\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "reverse_binding": 'Set b.Source = source\nSet a.Source = source\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "detach": 'Set a.Source = source\nSet a.Source = Nothing\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "rebind": 'Set a.Source = source\nSet a.Source = other\nProbe = CStr(source.Fire(1)) & ":" & CStr(other.Fire(2)) & ":" & Log',
    "repeat_binding": 'Set a.Source = source\nSet b.Source = source\nSet a.Source = source\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "disconnect_in_handler": 'Set a.Source = source\nSet b.Source = source\na.Disconnect = True\nProbe = CStr(source.Fire(1)) & ":" & CStr(source.Fire(2)) & ":" & Log',
    "disconnect_later_listener": 'Set a.Source = source\nSet b.Source = source\nSet a.Peer = b\na.DisconnectPeer = True\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "disconnect_earlier_listener": 'Set a.Source = source\nSet b.Source = source\nSet b.Peer = a\nb.DisconnectPeer = True\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "byval_event": 'Set a.Source = source\nSet b.Source = source\nProbe = CStr(source.FireValue(1)) & ":" & Log',
    "listener_scope": 'Set a.Source = source\nSet a = Nothing\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "rebind_later_listener": 'Set a.Source = source\nSet b.Source = source\nSet a.Peer = b\na.RebindPeer = True\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "nested_event": 'Set a.Source = source\nSet b.Source = source\na.Reenter = True\nProbe = CStr(source.Fire(1)) & ":" & Log',
    "multiple_declarations": 'Set a.OtherSource = source\nProbe = CStr(source.Fire(1)) & ":" & Log',
}


def main():
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        excel.add_module("EventLog", "Public Log As String\n")
        lines = ['Function Install() As String', 'Dim part As Object, text As String']
        for name, code in (("EventSource", SOURCE), ("EventSink", SINK)):
            lines += ['Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', f'part.Name = "{name}"', 'text = ""']
            for line in code.splitlines():
                lines.append('text = text & "' + line.replace('"', '""') + '" & vbLf')
            lines.append('part.CodeModule.AddFromString text')
        lines += ['Install = Application.Version & "|" & Application.Build', 'End Function']
        installed = excel.run_vba('\n'.join(lines), "Install", timeout=30.0, module_name="Installer")
        assert installed.ok, installed
        rows = []
        for name, action in ACTIONS.items():
            code = 'Function Probe() As String\n' + SETUP + action + '\nEnd Function\n'
            result = excel.run_vba(code, "Probe", timeout=30.0)
            assert result.ok, (name, result)
            print(name, result.value, flush=True)
            rows.append({"name": name, "source": code, "result": result.value})
    DESTINATION.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "excel_version_build": installed.value, "modules": {"EventSource": SOURCE, "EventSink": SINK, "EventLog": "Public Log As String\n"}, "probes": rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
