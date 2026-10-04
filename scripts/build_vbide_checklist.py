"""Generate the VBIDE member coverage table from the pinned local typelib dump."""
from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))


def main() -> None:
    from pyopenvba.apps.excel._vbide import VBProject, VBComponents, VBComponent, CodeModule

    classes = {cls.vba_type_name: cls for cls in (VBProject, VBComponents, VBComponent, CodeModule)}
    measured = {
        'VBProject': {'name', 'vbcomponents'},
        'VBComponents': {'count', 'item', 'add', 'remove'},
        'VBComponent': {'name', 'type', 'codemodule'},
        'CodeModule': {'lines', 'countoflines', 'countofdeclarationlines', 'addfromstring', 'insertlines',
                       'deletelines', 'replaceline', 'procstartline', 'procbodyline', 'proccountlines', 'procofline'},
    }
    reference = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT.parent / 'pyVBAReference'
    rows: list[dict[str, str]] = []
    for path in sorted((reference / 'reference/vbide/json').glob('*.json')):
        data = json.loads(path.read_text(encoding='utf-8'))
        name = data.get('name', path.stem)
        cls = classes.get(name)
        for category in ('properties', 'methods', 'events'):
            for entry in data.get(category, []):
                member_name = entry['name']
                implemented = cls is not None and member_name.lower() in cls._vba_members
                covered = member_name.lower() in measured.get(name, set())
                rows.append({'object': name, 'member': member_name, 'kind': entry.get('kind', category[:-1]),
                             'status': 'PARTIAL' if implemented and covered else 'UNASSESSED' if implemented else 'MISSING',
                             'evidence': 'tests/fixtures/vbide_projects.json; tests/test_excel_vbide.py' if implemented and covered else '',
                             'note': 'Measured cases only; wider input, lifetime and source-format conformance remains incomplete.' if implemented and covered else ''})
    output = ROOT / 'docs/vbide_checklist.csv'
    with output.open('w', newline='', encoding='utf-8') as stream:
        writer = csv.DictWriter(stream, fieldnames=('object', 'member', 'kind', 'status', 'evidence', 'note'))
        writer.writeheader()
        writer.writerows(rows)
    print(f'{len(rows)} VBIDE members written to {output}')


if __name__ == '__main__':
    main()
