"""Reference-only edits must be observed before injecting any VBA module."""
import os
import sys
from pathlib import Path

import pytest

from pyopenvba import ExcelFile, WordFile, PowerPointFile

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_REFERENCES") != "1" or sys.platform != "win32",
    reason="set RUN_LIVE_REFERENCES=1 with desktop Office installed",
)


@pytest.mark.parametrize("host", ["excel", "word", "powerpoint"])
def test_reference_only_edits_in_office(host: str, tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    kinds = {"excel": ExcelFile, "word": WordFile, "powerpoint": PowerPointFile}
    suffixes = {"excel": ".xlsm", "word": ".docm", "powerpoint": ".pptm"}
    sessions = {"excel": harness.ExcelSession, "word": harness.WordSession, "powerpoint": harness.PowerPointSession}
    projects = {"excel": "Application.Workbooks(1).VBProject", "word": "Application.Documents(1).VBProject",
                "powerpoint": "Application.Presentations(1).VBProject"}
    kind = kinds[host]
    path = tmp_path / ("references" + suffixes[host])
    file = kind.create_new(path)
    before = file.vba_modules()
    library, type_name = ("Word", "Document") if host == "excel" else ("Excel", "Workbook")
    added = file.add_reference(library)
    file.save()
    file.close()
    project = projects[host]
    with sessions[host](harness.HarnessConfig(lock_wait_s=45.0)) as office:
        office.open_document(path)
        # eval is COM/VBScript; it does not edit the VBA project first.
        assert str(office.eval(f'{project}.References("{library}").GUID')).upper() == added.guid
        count = int(office.eval(f"{project}.References.Count"))
        code = f'''Public Function Probe() As String
Dim item As {library}.{type_name}
Probe = TypeName(item)
End Function'''
        result = office.run_vba(code, "Probe", timeout=120.0)
        assert result.ok, result.message
        assert result.value == "Nothing"
    with kind(path) as file:
        assert file.remove_reference(library)
        file.save()
    with kind(path) as reopened:
        assert reopened.vba_modules() == before
    with sessions[host](harness.HarnessConfig(lock_wait_s=45.0)) as office:
        office.open_document(path)
        assert int(office.eval(f"{project}.References.Count")) == count - 1
        names = [office.eval(f"{project}.References({index}).Name") for index in range(1, count)]
        assert library not in names
