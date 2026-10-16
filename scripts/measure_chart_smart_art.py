"""Create native SmartArt evidence that a graphic frame need not be a chart."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/charts'

def main() -> None:
    OUT.mkdir(exist_ok=True)
    target = OUT / 'smart_art.xlsx'
    source = ('Function Probe() As String\nApplication.DisplayAlerts = False\nDim item As Shape\n'
              'Set item = ActiveSheet.Shapes.AddSmartArt(Application.SmartArtLayouts(1), 0, 0, 200, 100)\n'
              'item.Name = "Diagram"\nProbe = ActiveSheet.ChartObjects.Count & ":" & item.HasChart & ":" & item.Type\n'
              f'ActiveWorkbook.SaveAs "{target}", xlOpenXMLWorkbook\nEnd Function')
    with ExcelSession() as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        result = excel.run_vba(source, 'Probe', timeout=60)
        assert result.ok and not result.dialogs, result
        (OUT / 'smart_art.json').write_text(json.dumps({'macro': source, 'result': result.value,
            'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat()}, indent=2) + '\n', encoding='utf-8')
        print(result.value, flush=True)

if __name__ == '__main__':
    main()
