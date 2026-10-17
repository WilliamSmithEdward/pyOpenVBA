"""Recheck typed pivot records and native default aggregation after materialization."""
from __future__ import annotations

import datetime as dt
import json
import tempfile
from pathlib import Path
from zipfile import ZipFile

from pyvbaharness import ExcelSession
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._pivots import NS

INSPECT = '''Function InspectCache() As String
Dim p As PivotTable, f As PivotField
Set p=Worksheets(1).PivotTables(1)
Set f=p.AddDataField(p.PivotFields("Amount"))
InspectCache=CStr(f.Function) & "|" & CStr(p.DataBodyRange.Cells(1,1).Value2)
End Function'''


def main() -> None:
    path = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivots/record_types.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    with tempfile.TemporaryDirectory() as folder, ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\n'
                               'VersionBuild=Application.Version & "|" & Application.Build\nEnd Function',
                               'VersionBuild', line_numbers=False)
        assert version.ok and not version.dialogs, version
        record['excel_version_build'] = version.value
        for case in record['probes']:
            excel.new_document()
            saved = Path(folder) / (case['name'] + '.xlsx')
            source = case['source'].replace('Probe=CStr(c.RecordCount)',
                       'Probe=CStr(c.RecordCount)\nThisWorkbook.SaveAs "' + str(saved) + '",xlOpenXMLWorkbook')
            result = excel.run_vba(source, 'Probe', line_numbers=False)
            assert result.ok and result.value == case['result'] and not result.dialogs, (case['name'], result)
            expected = case['record']
            if expected is not None:
                with ZipFile(saved) as package:
                    part = next(name for name in package.namelist() if name.startswith(
                        ('xl/pivotCache/pivotCacheRecords', 'xl/pivotCache/richPivotRecords')) and name.endswith('.xml'))
                    root = xml_tree(package.read(part))
                    row = root.find(NS + 'r')
                    assert row is not None
                    item = row[1]
                    assert item.tag == NS + expected['kind']
                    assert ('richPivotRecords' in part) == expected['rich']
                    observed, wanted = item.get('v', ''), expected['attributes'].get('v', '')
                    if expected['kind'] == 'n':
                        assert float(observed) == float(wanted), case['name']
                    elif expected['kind'] == 'd':
                        assert dt.datetime.fromisoformat(observed) == dt.datetime.fromisoformat(wanted), case['name']
                    else:
                        assert observed == wanted, case['name']
                inspected = excel.run_vba(INSPECT, 'InspectCache', line_numbers=False)
                assert inspected.ok and not inspected.dialogs, inspected
                if 'inspection' in case:
                    assert inspected.value == case['inspection'], (case['name'], inspected)
                case['inspection'] = inspected.value
            print(case['name'], 'matched', flush=True)
    record['measured_at'] = dt.datetime.now(dt.timezone.utc).isoformat()
    path.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
