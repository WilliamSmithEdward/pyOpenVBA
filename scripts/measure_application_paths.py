"""How Application's path members relate in live Excel, with no path written down: whether Path ends in root\\Office16, and
each other member's relation to Path. William chose the install path the model answers for Application.Path, as a real
one names the measuring machine; only the relations here are measured.

    python scripts/measure_application_paths.py

writes tests/fixtures/application_paths.json, which tests/test_excel_application_paths.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/application_paths.json"

PROBE = r'''Function Relations() As String
    Dim s As String, p As String, v As String
    p = Application.Path
    s = "office16=" & (LCase$(Right$(p, 14)) = "\root\office16")
    v = Application.LibraryPath
    s = s & "|library=" & (LCase$(Left$(v, Len(p))) = LCase$(p)) & ":" & Mid$(v, Len(p) + 1)
    s = s & "|startup=" & (LCase$(Left$(Application.StartupPath, Len(p))) = LCase$(p))
    s = s & "|templates=" & (LCase$(Left$(Application.TemplatesPath, Len(p))) = LCase$(p))
    s = s & "|user_library=" & (LCase$(Left$(Application.UserLibraryPath, Len(p))) = LCase$(p))
    s = s & "|default_file=" & (LCase$(Left$(Application.DefaultFilePath, Len(p))) = LCase$(p))
    Relations = s
End Function
'''


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=60.0)) as excel:
        excel.new_document()
        result = excel.run_vba(PROBE, "Relations", timeout=30.0)
        assert result.ok, result
    print(result.value)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "relations": str(result.value)},
                              indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
