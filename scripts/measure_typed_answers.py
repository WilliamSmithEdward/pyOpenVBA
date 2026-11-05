"""Add the types to fixtures that kept answers as text: their own macros, run again in live Excel with each CStr(x) also giving TypeName(x).

The rewrite is written in line, as a helper would clear Err when it returned, and evaluates x twice. With the types taken out, an answer has to be the text the
fixture holds; where it is not, Excel or the fixture has drifted, and that is recorded as `drifted` and never written over the old fixture. The old fixtures stay as they are.

    python scripts/measure_typed_answers.py currency_writes sheet_quoting

writes tests/fixtures/typed_answers/<name>.json, which tests/test_excel_typed_answers.py replays.
"""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'tests/fixtures'
TYPES = ('String', 'Long', 'Integer', 'Double', 'Single', 'Currency', 'Date', 'Boolean', 'Byte', 'Decimal', 'Empty', 'Null', 'Error', 'Variant', 'Object', 'Range', 'Nothing')
TAG = re.compile(r'\b(?:' + '|'.join(TYPES) + r'):')


def typed_source(source: str) -> str:
    """Each CStr(x) as (TypeName(x) & ":" & CStr(x)), by the parentheses, leaving strings alone."""
    out: list[str] = []
    at = 0
    while True:
        found = re.compile(r'\bCStr\(', re.IGNORECASE).search(source, at)
        if found is None:
            return ''.join(out) + source[at:]
        depth, here, quoted = 1, found.end(), False
        while depth:
            char = source[here]
            if char == '"':
                quoted = not quoted
            elif not quoted and char == '(':
                depth += 1
            elif not quoted and char == ')':
                depth -= 1
            here += 1
        inner = source[found.end():here - 1]
        before = source[at:found.start()]
        # Inside a string it is text, not a call.
        if before.count('"') % 2 or ''.join(out).count('"') % 2:
            out.append(source[at:here])
        else:
            out.append(f'{before}(TypeName({inner}) & ":" & CStr({inner}))')
        at = here


def probes_of(name: str) -> list[dict[str, str]]:
    """A fixture's probes: each a name, a macro with the routine to run, and the old answer."""
    data = json.loads((FIXTURES / f'{name}.json').read_text(encoding='utf-8'))
    if name == 'shapes/checkbox_links':
        sys.path.insert(0, str(ROOT / 'scripts'))
        from measure_checkbox_links import report_source  # pyright: ignore[reportMissingImports]  # noqa: PLC0415

        return [{'name': one['name'], 'source': report_source(one['statement']), 'routine': 'Report', 'old': one['reported']} for one in data]
    return [{'name': one['name'], 'source': one.get('source') or one['macro'], 'routine': 'Probe', 'old': str(one.get('result'))} for one in data['probes']]


def untyped(answer: str) -> str:
    return TAG.sub('', answer)


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # pyright: ignore[reportAttributeAccessIssue]

    for name in sys.argv[1:]:
        cases = []
        with ExcelSession(HarnessConfig(exclusive=False)) as excel:
            for case in probes_of(name):
                source = case['source']
                if 'CStr(' not in source:
                    continue
                excel.new_document()
                done = excel.run_vba(typed_source(source), case['routine'], line_numbers=False, timeout=60.0)
                typed = str(done.value) if done.ok else f'failed {done.outcome} {[one.message[:80] for one in done.dialogs]}'
                old = case['old']
                cases.append({'name': case['name'], 'typed': typed, 'drifted': '' if untyped(typed) == old else old})
                print(name, case['name'], typed[:90], 'DRIFT' if cases[-1]['drifted'] else '', flush=True)
        target = FIXTURES / 'typed_answers' / f'{name.replace("/", "__")}.json'
        target.parent.mkdir(exist_ok=True)
        target.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
