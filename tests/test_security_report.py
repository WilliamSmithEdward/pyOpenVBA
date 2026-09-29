"""scripts/security_report.py, the check a release's security scans pass.

Every finding and every tool warning has to be one the accepted list
names, an accepted entry that matches nothing fails as well, and so does a
required scan that wrote no SARIF.  The script reads the list with
tomllib, so these run on Python 3.11 and later; CI runs the script on 3.12.
"""

from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

pytest.importorskip("tomllib")

ROOT = Path(__file__).parents[1]
SCRIPT = ROOT / "scripts" / "security_report.py"
ACCEPTED = ROOT / ".github" / "security" / "accepted.toml"


def load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("security_report", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def sarif(tool: str, results: list[dict[str, Any]], notes: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    return {
        "version": "2.1.0",
        "runs": [
            {
                "tool": {"driver": {"name": tool, "semanticVersion": "1.2.3", "rules": []}},
                "results": results,
                "invocations": [{"executionSuccessful": True, "toolExecutionNotifications": notes or []}],
            }
        ],
    }


def result(rule: str, path: str, line: int) -> dict[str, Any]:
    return {
        "ruleId": rule,
        "level": "warning",
        "message": {"text": "a message"},
        "locations": [{"physicalLocation": {"artifactLocation": {"uri": path}, "region": {"startLine": line}}}],
    }


ACCEPT_SHA1 = '''
[[finding]]
tool = "Semgrep"
rule = "sha1"
where = [{ path = "src/digest.py", line = "h = hashlib.sha1()" }]
reason = """
A fingerprint, not a signature.
"""
'''


DIGEST = "import hashlib\n\n\ndef digest():\n    h = hashlib.sha1()\n    return h\n"


def run(
    tmp_path: Path, scans: dict[str, dict[str, Any]], accepted: str, *require: str, source: str = DIGEST
) -> tuple[int, str]:
    """The script's exit status and report over `scans`, SARIF by scan name,
    against a source tree holding one module."""
    (tmp_path / "src").mkdir(exist_ok=True)
    (tmp_path / "src" / "digest.py").write_text(source, encoding="utf-8")
    sarif_dir = tmp_path / "sarif"
    sarif_dir.mkdir(exist_ok=True)
    for name, content in scans.items():
        (sarif_dir / f"{name}.sarif").write_text(json.dumps(content), encoding="utf-8")
    listed = tmp_path / "accepted.toml"
    listed.write_text(accepted, encoding="utf-8")
    out = tmp_path / "report.md"
    args = [str(sarif_dir), "--accepted", str(listed), "--out", str(out), "--source-root", str(tmp_path)]
    for name in require:
        args += ["--require", name]
    status = getattr(load(), "main")(args)
    return status, out.read_text(encoding="utf-8")


def test_an_accepted_finding_passes_and_the_report_gives_its_reason(tmp_path: Path) -> None:
    status, report = run(tmp_path, {"semgrep": sarif("Semgrep OSS", [result("sha1", "src/digest.py", 5)])}, ACCEPT_SHA1)
    assert status == 0
    assert "**Passed.**" in report
    assert "- src/digest.py:5: `h = hashlib.sha1()`" in report
    assert "A fingerprint, not a signature." in report
    assert "| semgrep | Semgrep 1.2.3 | 1 | 0 |" in report


def test_a_finding_not_on_the_list_fails(tmp_path: Path) -> None:
    findings = [result("sha1", "src/digest.py", 5), result("eval", "src/digest.py", 1)]
    status, report = run(tmp_path, {"semgrep": sarif("Semgrep OSS", findings)}, ACCEPT_SHA1)
    assert status == 1
    assert "**Failed.**" in report
    assert "- semgrep: `eval` (warning) at src/digest.py:1: `import hashlib`. a message" in report


def test_an_accepted_place_follows_its_line_but_not_a_change_to_it(tmp_path: Path) -> None:
    """The list quotes the flagged line, so code moving above it changes
    nothing, and a change to the line itself brings it back for review."""
    moved = "# a comment\n# and another\n" + DIGEST
    scans = {"semgrep": sarif("Semgrep OSS", [result("sha1", "src/digest.py", 7)])}
    status, report = run(tmp_path, scans, ACCEPT_SHA1, source=moved)
    assert status == 0
    assert "- src/digest.py:7: `h = hashlib.sha1()`" in report
    changed = DIGEST.replace("h = hashlib.sha1()", "h = hashlib.sha1(b'')")
    scans = {"semgrep": sarif("Semgrep OSS", [result("sha1", "src/digest.py", 5)])}
    status, report = run(tmp_path, scans, ACCEPT_SHA1, source=changed)
    assert status == 1
    assert "- semgrep: `sha1` (warning) at src/digest.py:5: `h = hashlib.sha1(b'')`" in report
    assert "Accepted, but no longer found: Semgrep sha1 at src/digest.py: `h = hashlib.sha1()`." in report


def test_an_accepted_entry_that_matches_nothing_fails(tmp_path: Path) -> None:
    status, report = run(tmp_path, {"semgrep": sarif("Semgrep OSS", [])}, ACCEPT_SHA1)
    assert status == 1
    assert "Accepted, but no longer found: Semgrep sha1 at src/digest.py" in report


def test_the_same_rule_from_another_tool_is_not_accepted(tmp_path: Path) -> None:
    scans = {
        "semgrep": sarif("Semgrep OSS", [result("sha1", "src/digest.py", 5)]),
        "codeql-python": sarif("CodeQL", [result("sha1", "src/digest.py", 5)]),
    }
    status, report = run(tmp_path, scans, ACCEPT_SHA1)
    assert status == 1
    assert "- codeql-python: `sha1`" in report


def test_a_required_scan_with_no_sarif_fails(tmp_path: Path) -> None:
    scans = {"semgrep": sarif("Semgrep OSS", [result("sha1", "src/digest.py", 5)])}
    status, report = run(tmp_path, scans, ACCEPT_SHA1, "semgrep", "codeql-python")
    assert status == 1
    assert "No SARIF from the `codeql-python` scan, which is required." in report


def test_a_tool_warning_fails_unless_accepted(tmp_path: Path) -> None:
    """A warning is a scan that may not have covered everything; one with
    no level is a warning, as SARIF defines it, and a note is not."""
    notes = [
        {"descriptor": {"id": "Syntax error"}, "message": {"text": "Syntax error at .github/workflows/x.yml:62"}},
        {"descriptor": {"id": "extracted"}, "level": "note", "message": {"text": "fine"}},
        {"descriptor": {"id": "extracted"}, "level": "none", "message": {"text": "fine"}},
    ]
    scans = {"semgrep": sarif("Semgrep OSS", [result("sha1", "src/digest.py", 5)], notes)}
    status, report = run(tmp_path, scans, ACCEPT_SHA1)
    assert status == 1
    assert "- semgrep: warning `Syntax error`: Syntax error at .github/workflows/x.yml:62" in report
    accepted = ACCEPT_SHA1 + (
        '\n[[notice]]\ntool = "Semgrep"\nid = "Syntax error"\ncontains = "x.yml"\nreason = "PowerShell, read as Bash."\n'
    )
    status, report = run(tmp_path, scans, accepted)
    assert status == 0
    assert "PowerShell, read as Bash." in report


def test_a_run_that_did_not_complete_fails(tmp_path: Path) -> None:
    broken = sarif("CodeQL", [])
    broken["runs"][0]["invocations"][0]["executionSuccessful"] = False
    status, report = run(tmp_path, {"codeql-python": broken}, "")
    assert status == 1
    assert "error `execution failed`" in report


def test_a_finding_takes_its_rule_s_level_when_it_gives_none(tmp_path: Path) -> None:
    """CodeQL leaves a result's level to the rule, which sits in a query
    pack's entry under tool.extensions."""
    scan = sarif("CodeQL", [result("py/weak-hash", "src/digest.py", 5)])
    del scan["runs"][0]["results"][0]["level"]
    scan["runs"][0]["tool"]["extensions"] = [
        {"name": "codeql/python-queries", "rules": [{"id": "py/weak-hash", "defaultConfiguration": {"level": "error"}}]}
    ]
    status, report = run(tmp_path, {"codeql-python": scan}, "")
    assert status == 1
    assert "- codeql-python: `py/weak-hash` (error) at src/digest.py:5" in report


@pytest.mark.parametrize(
    "entry",
    [
        '[[finding]]\ntool = "Semgrep"\nrule = "sha1"\nwhere = [{ path = "a.py", line = "x" }]\n',
        '[[finding]]\ntool = "Semgrep"\nrule = "sha1"\nwhere = []\nreason = "r"\n',
        '[[finding]]\ntool = "Semgrep"\nrule = "sha1"\nwhere = [{ path = "a.py", line = "x" }]\nreason = "r"\nrules = 1\n',
        '[[finding]]\ntool = "Semgrep"\nrule = "sha1"\nwhere = [{ path = "a.py", text = "x" }]\nreason = "r"\n',
        '[[finding]]\ntool = "Semgrep"\nrule = "sha1"\nwhere = [{ path = "a.py", line = "x" }]\nreason = "  "\n',
        '[[findings]]\ntool = "Semgrep"\n',
    ],
    ids=["no reason", "no place", "unknown key", "unknown place key", "blank reason", "unknown table"],
)
def test_a_malformed_list_is_refused_whole(entry: str) -> None:
    with pytest.raises(ValueError):
        getattr(load(), "load_accepted")(entry)


def test_the_repository_s_list_names_lines_the_source_holds() -> None:
    """Each accepted place quotes a line its file has, so the list can be
    read against the code without running a scan."""
    findings, _notices = getattr(load(), "load_accepted")(ACCEPTED.read_text(encoding="utf-8"))
    assert findings
    for entry in findings:
        assert entry.reason
        for place in entry.where:
            lines = [line.strip() for line in (ROOT / place.path).read_text(encoding="utf-8").splitlines()]
            assert place.line in lines, f"{place.path} has no line {place.line!r}"
