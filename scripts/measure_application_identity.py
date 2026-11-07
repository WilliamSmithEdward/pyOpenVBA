"""Application's paths and names, which a macro reads to find a folder, typed and late bound: their types, and what a set of each does.

Predictions, written before the run: each path and name is a String; StartupPath, UserLibraryPath, TemplatesPath, NetworkTemplatesPath,
OrganizationName and ProductCode refuse a set; DefaultFilePath and AltStartupPath take a folder and read it back as given, and 1004 for one
that is not there; FeatureInstall is a Long of 0 to 2, 1004 beyond; SharePointVersion answers 0 for a site it cannot reach.

The settings are given back, and it holds live.lock while it runs: AltStartupPath outlasts the session. It leaves out the rows that act on the
machine or show something: SendKeys, ActivateMicrosoftApp, Help, FindFile, NextLetter, DisplayXMLSourcePane, and every member that could send mail.

    python scripts/measure_application_identity.py

writes tests/fixtures/application_identity.json, which tests/test_excel_application_identity.py replays.
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
OUT = ROOT / "tests/fixtures/application_identity.json"
READS = ["StartupPath", "AltStartupPath", "UserLibraryPath", "TemplatesPath", "NetworkTemplatesPath", "OrganizationName", "ProductCode", "DefaultFilePath", "ActivePrinter",
         "FeatureInstall", 'SharePointVersion("http://example.invalid/")', 'SharePointVersion("")', "SharePointVersion(5)", "SharePointVersion(True)"]
READ_ONLY = ["StartupPath", "UserLibraryPath", "TemplatesPath", "NetworkTemplatesPath", "OrganizationName", "ProductCode"]
#: {folder} is a folder that is there and {gone} one that is not.
SETS = {"DefaultFilePath": ['"{folder}"', '"{gone}"', '""', "5", "True", '"{folder}\\"', "Null"],
        "AltStartupPath": ['"{folder}"', '"{gone}"', '""', "5", "True", '"{folder}\\"', "Null"],
        "FeatureInstall": ["0", "1", "2", "3", "-1", '"1"', "True", "1.6", "Null"],
        "ActivePrinter": ['"Nobody at all"', '""', "5", "Null"]}
#: What a macro does through the object each set of bindings gives it.
SOURCE = """Function Go(ByVal folder As String) As String
    {dim}
    Dim v As Variant
    Set a = Application
    On Error Resume Next
    {action}
    If Err.Number <> 0 Then
        Go = "E" & Err.Number
    Else
        Go = TypeName(v) & ":" & CStr(v)
    End If
End Function
"""
BINDINGS = {"typed": "Dim a As Application", "late": "Dim a As Object"}


def source(binding: str, action: str) -> str:
    return SOURCE.format(dim=BINDINGS[binding], action=action)


def read_action(what: str) -> str:
    return f"v = a.{what}"


def set_action(name: str, value: str) -> str:
    return f"a.{name} = {value}: If Err.Number = 0 Then v = a.{name}"


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    common = Path(subprocess.run(["git", "rev-parse", "--git-common-dir"], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip())
    lock = (common if common.is_absolute() else ROOT / common) / "excel-coordination" / "live.lock"
    for _ in range(600):
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            time.sleep(1)
    else:
        raise TimeoutError("live.lock stayed held")
    made_for_restore: list[Path] = []
    try:
        with tempfile.TemporaryDirectory(prefix="pyopenvba_id_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            excel.new_document()
            home, user = os.path.expanduser("~"), getpass.getuser()

            def masked(text: str) -> str:
                return re.sub(re.escape(user), "<user>", re.sub(re.escape(home), "<home>", re.sub(re.escape(folder), "<folder>", text, flags=re.I), flags=re.I), flags=re.I)

            def run(binding: str, action: str) -> str:
                answer = excel.run_vba(source(binding, action), "Go", (folder,), timeout=60.0, line_numbers=False)
                return masked(str(answer.value)) if answer.ok else f"[{answer.outcome}]"

            rows: dict[str, dict[str, str]] = {}
            start = {name: excel.run_vba(source("late", f"v = a.{name}"), "Go", (folder,), timeout=60.0, line_numbers=False).value for name in SETS}
            try:
                for what in READS:
                    rows[f"read {what}"] = {binding: run(binding, read_action(what)) for binding in BINDINGS}
                for name in READ_ONLY:
                    # Typed it does not compile, which stops the run with a dialog and, a second time, the harness.
                    rows[f"set {name}"] = {"late": run("late", f'a.{name} = "x"')}
                for name, values in SETS.items():
                    for value in values:
                        action = set_action(name, value.replace("{folder}", '" & folder & "').replace("{gone}", '" & folder & "\\gone'))
                        rows[f"set {name} = {value}"] = {binding: run(binding, action) for binding in BINDINGS}
                        print(name, value, rows[f"set {name} = {value}"], flush=True)
            finally:
                # Excel ignores a set to a folder that is not there, and DefaultFilePath may be one, so it is there for the set.
                wanted = Path(str(start["DefaultFilePath"]).partition(":")[2])
                while wanted != wanted.parent and not wanted.exists():
                    made_for_restore.append(wanted)
                    wanted = wanted.parent
                for one in reversed(made_for_restore):
                    one.mkdir()
                for name, value in start.items():
                    given = str(value).partition(":")[2]
                    number = name == "FeatureInstall"
                    restore = given if number else '"' + given.replace('"', '""') + '"'
                    excel.run_vba(source("late", f"a.{name} = {restore}"), "Go", (folder,), timeout=60.0, line_numbers=False)
                after = {name: excel.run_vba(source("late", f"v = a.{name}"), "Go", (folder,), timeout=60.0, line_numbers=False).value for name in SETS}
                assert after == start, (after, start)
    finally:
        for one in made_for_restore:
            one.rmdir()
        os.rmdir(lock)
    for key, row in rows.items():
        print(key, row)
    OUT.write_text(json.dumps({"rows": rows, "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
