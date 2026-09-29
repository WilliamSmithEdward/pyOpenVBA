"""Hold the security scans to the accepted list and write the security report.

CodeQL and Semgrep each write SARIF.  Every result in it has to be one
that ``.github/security/accepted.toml`` lists, and so does every warning
or error a tool raised while it scanned.  Anything else is unexpected and
fails the check.  So does an accepted entry that matches nothing, since a
list that outlives its findings no longer says what the code does, and so
does a scan that was required but wrote no SARIF.

The report says what ran, what each scan found, what was accepted and
why, and what was not.  The Security workflow runs this on every push,
pull request and release, and a release carries the report and the SARIF
it was made from.

Usage::

    python scripts/security_report.py SARIF_DIR --accepted FILE --out FILE
        [--require NAME]... [--source-root DIR]

A scan is named after its SARIF file, ``codeql-python.sarif`` making the
scan ``codeql-python``.  The exit status is 1 when anything was unexpected.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import unquote

#: The keys an accepted entry may have, by kind; anything else is a typo.
FINDING_KEYS = {"tool", "rule", "reason", "where"}
PLACE_KEYS = {"path", "line"}
NOTICE_KEYS = {"tool", "id", "contains", "reason"}


@dataclass(frozen=True)
class Finding:
    """One result a scan reported, where it points and the line it flags."""

    scan: str
    tool: str
    rule: str
    level: str
    path: str
    line: int
    text: str
    message: str


@dataclass(frozen=True)
class Notice:
    """A warning or error a tool raised about the scan itself."""

    scan: str
    tool: str
    ident: str
    level: str
    message: str


@dataclass
class Scan:
    name: str
    tool: str
    version: str
    findings: list[Finding] = field(default_factory=list[Finding])
    notices: list[Notice] = field(default_factory=list[Notice])


@dataclass(frozen=True)
class Place:
    path: str
    line: str


@dataclass(frozen=True)
class AcceptedFinding:
    tool: str
    rule: str
    reason: str
    where: tuple[Place, ...]

    def matches(self, finding: Finding) -> Place | None:
        if finding.tool != self.tool or finding.rule != self.rule:
            return None
        return next((p for p in self.where if p.path == finding.path and p.line == finding.text), None)


@dataclass(frozen=True)
class AcceptedNotice:
    tool: str
    ident: str
    contains: str
    reason: str

    def matches(self, notice: Notice) -> bool:
        return notice.tool == self.tool and notice.ident == self.ident and self.contains in notice.message


@dataclass
class Outcome:
    scans: list[Scan]
    accepted: dict[AcceptedFinding, list[Finding]]
    accepted_notices: dict[AcceptedNotice, list[Notice]]
    unexpected: list[Finding]
    unexpected_notices: list[Notice]
    stale: list[str]
    missing: list[str]

    @property
    def passed(self) -> bool:
        return not (self.unexpected or self.unexpected_notices or self.stale or self.missing)


def tool_of(name: str) -> str:
    """The tool a SARIF run comes from, as the accepted list names it."""
    for tool in ("CodeQL", "Semgrep"):
        if name.startswith(tool):
            return tool
    return name


def load_accepted(text: str) -> tuple[list[AcceptedFinding], list[AcceptedNotice]]:
    """The accepted list, refused whole if an entry is malformed, so that
    a typo cannot accept something by accident."""
    import tomllib

    data: dict[str, Any] = tomllib.loads(text)
    unknown = set(data) - {"finding", "notice"}
    if unknown:
        raise ValueError(f"unknown tables in the accepted list: {sorted(unknown)}")
    findings: list[AcceptedFinding] = []
    for entry in data.get("finding", []):
        _check_keys(entry, FINDING_KEYS, "finding")
        places: list[Place] = []
        for place in entry["where"]:
            _check_keys(place, PLACE_KEYS, "where")
            places.append(Place(place["path"], place["line"].strip()))
        if not places:
            raise ValueError(f"accepted {entry['rule']} names no place")
        findings.append(AcceptedFinding(entry["tool"], entry["rule"], _reason(entry), tuple(places)))
    notices: list[AcceptedNotice] = []
    for entry in data.get("notice", []):
        _check_keys(entry, NOTICE_KEYS, "notice")
        notices.append(AcceptedNotice(entry["tool"], entry["id"], entry["contains"], _reason(entry)))
    return findings, notices


def _check_keys(entry: dict[str, Any], keys: set[str], kind: str) -> None:
    if set(entry) != keys:
        raise ValueError(f"an accepted {kind} needs exactly {sorted(keys)}, not {sorted(entry)}")


def _reason(entry: dict[str, Any]) -> str:
    reason = " ".join(str(entry["reason"]).split())
    if not reason:
        raise ValueError("an accepted entry needs a reason")
    return reason


def read_scan(name: str, sarif: dict[str, Any], source_root: Path) -> Scan:
    """One SARIF file's runs as a scan: every result, and every notification
    at warning level or above.  SARIF makes a notification a warning when it
    gives no level, and a failed run is an error of its own."""
    runs: list[dict[str, Any]] = sarif.get("runs", [])
    if not runs:
        raise ValueError(f"{name} holds no SARIF run")
    driver: dict[str, Any] = runs[0]["tool"]["driver"]
    scan = Scan(name, tool_of(driver.get("name", "")), driver.get("semanticVersion") or driver.get("version", ""))
    for run in runs:
        rules = _rule_levels(run)
        for result in run.get("results", []):
            scan.findings.append(_finding(scan, result, rules, source_root))
        for invocation in run.get("invocations", []):
            if invocation.get("executionSuccessful") is False:
                scan.notices.append(Notice(name, scan.tool, "execution failed", "error", "the run did not complete"))
            for key in ("toolExecutionNotifications", "toolConfigurationNotifications"):
                for note in invocation.get(key, []):
                    level = note.get("level", "warning")
                    if level in ("warning", "error"):
                        ident = note.get("descriptor", {}).get("id", "") or key
                        message = " ".join(note.get("message", {}).get("text", "").split())
                        scan.notices.append(Notice(name, scan.tool, ident, level, message))
    return scan


def _rule_levels(run: dict[str, Any]) -> dict[str, str]:
    """Each rule's default level: CodeQL leaves a result's own out."""
    components = [run["tool"]["driver"], *run["tool"].get("extensions", [])]
    levels: dict[str, str] = {}
    for component in components:
        for rule in component.get("rules", []):
            levels[rule["id"]] = rule.get("defaultConfiguration", {}).get("level", "warning")
    return levels


def _finding(scan: Scan, result: dict[str, Any], rules: dict[str, str], source_root: Path) -> Finding:
    rule = result.get("ruleId") or result.get("rule", {}).get("id", "")
    locations: list[dict[str, Any]] = result.get("locations") or [{}]
    physical: dict[str, Any] = locations[0].get("physicalLocation", {})
    uri: str = physical.get("artifactLocation", {}).get("uri", "")
    path = unquote(uri).replace("\\", "/").removeprefix("file://").removeprefix("./")
    region: dict[str, Any] = physical.get("region", {})
    line = int(region.get("startLine", 0))
    snippet: str = region.get("snippet", {}).get("text", "")
    text = _source_line(source_root / path, line) or snippet
    return Finding(
        scan.name,
        scan.tool,
        rule,
        result.get("level") or rules.get(rule, "warning"),
        path,
        line,
        text.strip(),
        " ".join(result.get("message", {}).get("text", "").split()),
    )


def _source_line(path: Path, line: int) -> str:
    """The flagged line as the checkout holds it, which the accepted list
    quotes; an accepted place follows code that moves but not code that
    changes."""
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeDecodeError):
        return ""
    return lines[line - 1] if 0 < line <= len(lines) else ""


def check(
    scans: list[Scan],
    accepted: list[AcceptedFinding],
    accepted_notices: list[AcceptedNotice],
    required: list[str],
) -> Outcome:
    outcome = Outcome(
        scans,
        {entry: [] for entry in accepted},
        {entry: [] for entry in accepted_notices},
        [],
        [],
        [],
        [name for name in required if name not in {scan.name for scan in scans}],
    )
    used: set[tuple[AcceptedFinding, Place]] = set()
    for scan in scans:
        for finding in scan.findings:
            match = _acceptance(accepted, finding)
            if match is None:
                outcome.unexpected.append(finding)
            else:
                outcome.accepted[match[0]].append(finding)
                used.add(match)
        for notice in scan.notices:
            entry = next((e for e in accepted_notices if e.matches(notice)), None)
            if entry is None:
                outcome.unexpected_notices.append(notice)
            else:
                outcome.accepted_notices[entry].append(notice)
    for entry in accepted:
        for place in entry.where:
            if (entry, place) not in used:
                outcome.stale.append(f"{entry.tool} {entry.rule} at {place.path}: `{place.line}`")
    for entry, matched in outcome.accepted_notices.items():
        if not matched:
            outcome.stale.append(f"{entry.tool} notice {entry.ident!r} mentioning {entry.contains!r}")
    return outcome


def _acceptance(accepted: list[AcceptedFinding], finding: Finding) -> tuple[AcceptedFinding, Place] | None:
    """The accepted entry and place that account for `finding`, if any."""
    for entry in accepted:
        place = entry.matches(finding)
        if place is not None:
            return entry, place
    return None


def report(outcome: Outcome, provenance: str) -> str:
    """The report a release carries, in Markdown."""
    counts = {scan.name: (len(scan.findings), 0) for scan in outcome.scans}
    for finding in outcome.unexpected:
        total, unexpected = counts[finding.scan]
        counts[finding.scan] = (total, unexpected + 1)
    lines = ["# pyOpenVBA security report", "", provenance, ""]
    if outcome.passed:
        accepted = sum(len(found) for found in outcome.accepted.values())
        lines.append(f"**Passed.** Every finding is one of the {accepted} accepted below, each with its reason.")
    else:
        lines.append("**Failed.** Each item under Unexpected needs a fix, or a change to the accepted list.")
    lines += ["", "| Scan | Tool | Findings | Unexpected |", "| --- | --- | ---: | ---: |"]
    for scan in sorted(outcome.scans, key=lambda s: s.name):
        total, unexpected = counts[scan.name]
        lines.append(f"| {scan.name} | {scan.tool} {scan.version} | {total} | {unexpected} |")
    lines += ["", "## Unexpected", ""]
    problems = [
        *(f"- No SARIF from the `{name}` scan, which is required." for name in outcome.missing),
        *(
            f"- {f.scan}: `{f.rule}` ({f.level}) at {f.path}:{f.line}: `{f.text}`. {f.message}"
            for f in outcome.unexpected
        ),
        *(f"- {n.scan}: {n.level} `{n.ident}`: {n.message}" for n in outcome.unexpected_notices),
        *(f"- Accepted, but no longer found: {entry}." for entry in outcome.stale),
    ]
    lines += problems or ["None."]
    lines += ["", "## Accepted", ""]
    if not outcome.accepted and not outcome.accepted_notices:
        lines.append("None.")
    for entry, found in outcome.accepted.items():
        lines.append(f"**{entry.rule}** ({entry.tool})")
        lines.append("")
        lines += [f"- {f.path}:{f.line}: `{f.text}`" for f in found]
        lines += ["", entry.reason, ""]
    for entry, noticed in outcome.accepted_notices.items():
        lines += [f"**{entry.ident}** ({entry.tool}), {len(noticed)} notices mentioning `{entry.contains}`", ""]
        lines += [entry.reason, ""]
    return "\n".join(lines).rstrip() + "\n"


def provenance(now: dt.datetime) -> str:
    """Which commit was scanned, when, and by which workflow run."""
    sha = os.environ.get("GITHUB_SHA", "")
    ref = os.environ.get("GITHUB_REF_NAME", "")
    when = now.strftime("%Y-%m-%d %H:%M UTC")
    if not sha:
        return f"A local run, {when}."
    run = "{}/{}/actions/runs/{}".format(
        os.environ.get("GITHUB_SERVER_URL", "https://github.com"),
        os.environ.get("GITHUB_REPOSITORY", ""),
        os.environ.get("GITHUB_RUN_ID", ""),
    )
    return f"Commit `{sha}` ({ref}), scanned {when} by {run}."


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Hold the security scans to the accepted list.")
    parser.add_argument("sarif_dir", type=Path)
    parser.add_argument("--accepted", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--require", action="append", default=[])
    parser.add_argument("--source-root", type=Path, default=Path())
    args = parser.parse_args(argv)

    accepted, accepted_notices = load_accepted(args.accepted.read_text(encoding="utf-8"))
    scans = [
        read_scan(path.stem, json.loads(path.read_text(encoding="utf-8")), args.source_root)
        for path in sorted(args.sarif_dir.glob("*.sarif"))
    ]
    outcome = check(scans, accepted, accepted_notices, args.require)
    text = report(outcome, provenance(dt.datetime.now(dt.timezone.utc)))
    args.out.write_text(text, encoding="utf-8")
    sys.stdout.write(text)
    return 0 if outcome.passed else 1


if __name__ == "__main__":
    sys.exit(main())
