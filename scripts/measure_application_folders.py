"""What Application.DefaultFilePath and AltStartupPath keep of the folder a macro sets, typed and late bound.

scripts/measure_application_identity.py found that DefaultFilePath ignores a folder that is not there and keeps one that is as
given, and that AltStartupPath keeps any text, with a folder that is not there partly in capitals and a bare name under Office's
own folder. Its masks hid the case of the folder, so it could not say which part goes to capitals. Each case here sets "" first,
which both take, so an ignored set reads back "".

Predictions, written before the run: DefaultFilePath keeps a folder that is there exactly as typed, case, slashes and a trailing
backslash included, and ignores a file, a folder that is not there and a bare name. AltStartupPath turns a folder into its full
path with no trailing backslash, keeping the case of the part that is there as typed and putting the rest in capitals, and puts a
bare name under Excel's current folder.

It holds live.lock while it runs, since both settings outlast the session, and gives both back. It touches no member that acts on
the machine or shows something, and no member that could send mail.

    python scripts/measure_application_folders.py

writes tests/fixtures/application_folders.json, which tests/test_excel_application_identity.py replays.
"""
from __future__ import annotations

import getpass
import json
import os
import re
import subprocess
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/application_folders.json"
NAMES = ["DefaultFilePath", "AltStartupPath"]
#: Each value as VBA, where folder is a folder that holds a folder Sub and a file f.txt.
VALUES = ['folder & "\\Sub"', 'UCase(folder & "\\Sub")', 'LCase(folder & "\\Sub")', 'folder & "\\sub"', 'folder & "\\Sub\\"',
          'folder & "\\Sub\\\\"', 'folder & "/Sub"', 'folder & "\\Sub\\..\\Sub"', 'folder & "\\.\\Sub"', 'folder & "\\f.txt"',
          'folder & "\\gone"', 'folder & "\\Gone\\deeper"', 'folder & "\\gone\\"', 'folder & "\\Sub\\gone"', 'folder & "\\Sub "',
          'folder', '"."', '".."', '"Sub"', '"abc"', '"aBc\\Def"', '"\\abc"', '"True"', 'True', '5', '1.5', 'Empty', '"C:"', '"C:\\"',
          '" "', '"x:y"', '"a*b"', 'folder & "\\a?b"']
SOURCE = """Function Go(ByVal folder As String) As String
    {dim}
    Dim v As Variant
    Set a = Application
    On Error Resume Next
    a.{name} = ""
    Err.Clear
    a.{name} = {value}
    If Err.Number <> 0 Then
        Go = "E" & Err.Number
    Else
        v = a.{name}
        Go = TypeName(v) & ":" & CStr(v)
    End If
End Function
"""
BINDINGS = {"typed": "Dim a As Application", "late": "Dim a As Object"}
READ = """Function Go(ByVal folder As String) As String
    Go = CStr(Application.{name})
End Function
"""
WRITE = """Function Go(ByVal folder As String) As String
    Application.{name} = folder
    Go = CStr(Application.{name})
End Function
"""
CURDIR = """Function Go(ByVal folder As String) As String
    Go = CurDir
End Function
"""


def masker(folder: str, home: str, user: str):  # type: ignore[no-untyped-def]
    """Masks the machine's folders, saying of the run's folder whether it came back as typed, in capitals or in small letters."""

    def mark(found: re.Match[str]) -> str:
        text = found.group(0)
        return "<folder>" if text == folder else "<FOLDER>" if text == folder.upper() else "<folder:lower>" if text == folder.lower() else "<folder:other>"

    def masked(text: str) -> str:
        text = re.sub(re.escape(folder), mark, text, flags=re.I)
        return re.sub(re.escape(user), "<user>", re.sub(re.escape(home), "<home>", text, flags=re.I), flags=re.I)

    return masked


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    common = Path(subprocess.run(["git", "rev-parse", "--git-common-dir"], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip())
    lock = (common if common.is_absolute() else ROOT / common) / "excel-coordination" / "live.lock"
    for _ in range(900):
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            time.sleep(1)
    else:
        raise TimeoutError("live.lock stayed held")
    rows: dict[str, dict[str, str]] = {}
    try:
        with tempfile.TemporaryDirectory(prefix="pyopenvba_folders_", ignore_cleanup_errors=True) as made, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            folder = str(Path(made) / "Mixed")
            os.mkdir(folder)
            os.mkdir(Path(folder) / "Sub")
            (Path(folder) / "f.txt").write_text("x", encoding="utf-8")
            excel.new_document()
            masked = masker(folder, os.path.expanduser("~"), getpass.getuser())

            def run(source: str) -> str:
                answer = excel.run_vba(source, "Go", (folder,), timeout=60.0, line_numbers=False)
                return masked(str(answer.value)) if answer.ok else f"[{answer.outcome}]"

            start = {name: excel.run_vba(READ.format(name=name), "Go", (folder,), timeout=60.0, line_numbers=False).value for name in NAMES}
            assert all(isinstance(value, str) for value in start.values()), start
            try:
                rows["CurDir"] = {"late": run(CURDIR)}
                for name in NAMES:
                    for value in VALUES:
                        rows[f"set {name} = {value}"] = {binding: run(SOURCE.format(dim=dim, name=name, value=value)) for binding, dim in BINDINGS.items()}
                        print(name, value, rows[f"set {name} = {value}"], flush=True)
            finally:
                for name, value in start.items():
                    restore = '"' + str(value).replace('"', '""') + '"'
                    excel.run_vba(WRITE.format(name=name).replace("= folder", f"= {restore}"), "Go", (folder,), timeout=60.0, line_numbers=False)
                after = {name: excel.run_vba(READ.format(name=name), "Go", (folder,), timeout=60.0, line_numbers=False).value for name in NAMES}
                assert after == start, (after, start)
    finally:
        os.rmdir(lock)
    OUT.write_text(json.dumps({"rows": rows, "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
