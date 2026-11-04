"""Pane, and what is left of Application and Window that needs no screen, measured in live Excel. Nothing here opens a
dialog, sends a key, reaches another program, moves a window or sets a user's option. A value that names the machine
is reported by its shape: how many add-ins and recent files there are is the measuring user's.

    python scripts/measure_application_rest.py

writes tests/fixtures/application_rest.json, which tests/test_excel_application_rest.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/application_rest.json"

P = "ActiveWindow.ActivePane"
#: One statement each, in its own error trap, with what it reads.
ROWS = {
    "pane_type": f"Note TypeName({P})",
    "pane_parent": f"Note TypeName({P}.Parent)",
    "pane_application": f"Note TypeName({P}.Application)",
    "pane_creator": f"Note {P}.Creator",
    "pane_activate": f"Note {P}.Activate",
    "pane_scroll_row": f"Note {P}.ScrollRow",
    "pane_scroll_column": f"Note {P}.ScrollColumn",
    "pane_scroll_row_set": f'{P}.ScrollRow = 20\nNote {P}.ScrollRow & "," & ActiveWindow.ScrollRow',
    "pane_scroll_column_set": f'{P}.ScrollColumn = 7\nNote {P}.ScrollColumn & "," & ActiveWindow.ScrollColumn',
    "window_sets_pane": f'ActiveWindow.ScrollRow = 9\nActiveWindow.ScrollColumn = 3\nNote {P}.ScrollRow & "," & {P}.ScrollColumn',
    "pane_scroll_row_zero": f"{P}.ScrollRow = 0",
    "pane_scroll_row_far": f"{P}.ScrollRow = 2000000",
    "pane_scroll_column_zero": f"{P}.ScrollColumn = 0",
    "pane_scroll_column_far": f"{P}.ScrollColumn = 20000",
    "pane_scroll_row_text": f'{P}.ScrollRow = "x"',
    "pane_scroll_row_null": f"{P}.ScrollRow = Null",
    "pane_small_down": f'{P}.ScrollRow = 10\nNote {P}.SmallScroll(Down:=2)\nNote {P}.ScrollRow',
    "pane_small_up": f'{P}.ScrollRow = 10\n{P}.SmallScroll Up:=3\nNote {P}.ScrollRow',
    "pane_small_right": f'{P}.ScrollColumn = 10\n{P}.SmallScroll ToRight:=2\nNote {P}.ScrollColumn',
    "pane_small_left": f'{P}.ScrollColumn = 10\n{P}.SmallScroll ToLeft:=3\nNote {P}.ScrollColumn',
    "pane_small_past_top": f'{P}.ScrollRow = 2\n{P}.SmallScroll Up:=9\nNote {P}.ScrollRow',
    "pane_small_nothing": f'{P}.ScrollRow = 10\n{P}.SmallScroll\nNote {P}.ScrollRow',
    "panes_default": 'Note TypeName(ActiveWindow.Panes.[_Default](1))',
    "panes_item_out": 'Note TypeName(ActiveWindow.Panes(2))',
    "panes_item_zero": 'Note TypeName(ActiveWindow.Panes(0))',
    "panes_walk": 'Dim one As Object\nFor Each one In ActiveWindow.Panes\nNote TypeName(one)\nNext',
    "windows_walk": 'Dim one As Object\nFor Each one In Application.Windows\nNote TypeName(one)\nNext',
    "app_on_window": "Note \"[\" & Application.OnWindow & \"]\"",
    "app_on_window_set": 'Application.OnWindow = "Nothing"\nNote Application.OnWindow\nApplication.OnWindow = ""\nNote "[" & Application.OnWindow & "]"',
    "window_on_window": 'Note "[" & ActiveWindow.OnWindow & "]"',
    "window_on_window_set": 'ActiveWindow.OnWindow = "Nothing"\nNote ActiveWindow.OnWindow\nActiveWindow.OnWindow = ""\nNote "[" & ActiveWindow.OnWindow & "]"',
    "used_objects_type": "Note TypeName(Application.UsedObjects)",
    "used_objects_item": "Note TypeName(Application.UsedObjects(1))",
    "used_objects_parent": "Note TypeName(Application.UsedObjects.Parent)",
    "add_ins_type": "Note TypeName(Application.AddIns)",
    "add_ins2_type": "Note TypeName(Application.AddIns2)",
    "add_ins_nowhere": 'Note TypeName(Application.AddIns("No such add-in"))',
    "add_ins_out": "Note TypeName(Application.AddIns(9999))",
    "com_add_ins_type": "Note TypeName(Application.COMAddIns)",
    "recent_type": "Note TypeName(Application.RecentFiles)",
    "recent_maximum": "Note Application.RecentFiles.Maximum",
    "recent_out": "Note TypeName(Application.RecentFiles(9999))",
    "export_converters_type": "Note TypeName(Application.FileExportConverters)",
    "record_macro": 'Application.RecordMacro "Dim x"',
    "record_macro_nothing": "Application.RecordMacro",
    "register_nowhere": 'Note Application.RegisterXLL("C:\\no_such_folder_pyopenvba\\nothing.xll")',
    "dde_terminate_nothing": "Application.DDETerminate 99999",
    "dde_execute_nothing": 'Application.DDEExecute 99999, "x"',
    "dde_poke_nothing": 'Application.DDEPoke 99999, "x", ActiveSheet.Range("A1")',
    "dde_request_nothing": 'Note TypeName(Application.DDERequest(99999, "x"))',
    "dde_request_value": 'Note CStr(Application.DDERequest(99999, "x"))',
    "dde_request_low": 'Note CStr(Application.DDERequest(1, "x"))',
    "dde_low_channels": 'Application.DDETerminate 1\nApplication.DDEExecute 1, "x"\nApplication.DDEPoke 1, "x", ActiveSheet.Range("A1")\nNote "ok"',
    "phonetic": 'Note "[" & Application.GetPhonetic("abc") & "]"',
    "phonetic_again": 'Note "[" & Application.GetPhonetic() & "]"',
    "phonetic_empty": 'Note "[" & Application.GetPhonetic("") & "]"',
}

PROBE = '''Dim Log As String
Sub Note(ByVal s As String)
    Log = Log & s & "|"
End Sub
Function Probe() As String
    Application.DisplayAlerts = False
    On Error GoTo Failed
{action}
    Probe = Log
    Exit Function
Failed:
    Probe = Log & "!E" & Err.Number
End Function
'''


def main() -> None:
    rows: list[dict[str, str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=60.0)) as excel:
        for name, action in ROWS.items():
            excel.new_document()
            result = excel.run_vba(PROBE.format(action=action), "Probe", timeout=30.0)
            value = str(result.value) if result.ok else f"failed {result.outcome}"
            print(name, value[:90], flush=True)
            rows.append({"name": name, "action": action, "result": value})
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "rows": rows},
                              indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
