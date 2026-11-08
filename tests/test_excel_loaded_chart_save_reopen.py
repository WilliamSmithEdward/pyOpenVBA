"""Set, read, save, reopen, read again, on the data labels, fills and lines of charts loaded from a file Excel saved (see set_save_reopen). Whether a kept
value matches Excel is tests/test_excel_loaded_chart_label_edits.py's and tests/test_excel_loaded_chart_look_edits.py's to say; this says the
file keeps what the model claims."""
from __future__ import annotations

from pathlib import Path

import pytest
from set_save_reopen import Case, Target, cases, check

FIXTURES = Path(__file__).parent / 'fixtures'
_SERIES = 'Set s = ws.ChartObjects(1).Chart.SeriesCollection({})'

TARGETS = [Target(f'{book} series {number}', _SERIES.format(number), _SERIES.format(number), FIXTURES / f'loaded_chart_{book}.xlsx')
           for book in ('plain', 'labelled', 'dressed', 'styled') for number in (1, 2)]
SETTINGS = {'s.HasDataLabels': ['True', 'False'], 's.Points(1).HasDataLabel': ['True', 'False'], 's.Points(1).DataLabel.Text': ['"P"'],
            's.Points(2).DataLabel.Position': ['4'], 's.Points(1).DataLabel.Font.Size': ['12'], 's.Points(2).DataLabel.ShowCategoryName': ['True'],
            's.DataLabels.ShowCategoryName': ['True'], 's.DataLabels.ShowSeriesName': ['True'], 's.DataLabels.ShowValue': ['False'],
            's.DataLabels.Position': ['3'], 's.DataLabels.NumberFormat': ['"0.0"'], 's.DataLabels.Separator': ['"; "'],
            's.Format.Fill.ForeColor.RGB': ['255'], 's.Format.Fill.Visible': ['0'], 's.Format.Fill.Transparency': ['0.5'],
            's.Format.Line.ForeColor.RGB': ['255'], 's.Format.Line.Weight': ['2'], 's.Format.Line.Visible': ['0'],
            's.Format.Line.DashStyle': ['4'], 's.Interior.Color': ['255'], 's.Points(1).Format.Fill.ForeColor.RGB': ['255'],
            's.Points(2).Format.Line.ForeColor.RGB': ['255'], 's.Values': ['"=One!$B$3:$C$3"'], 's.XValues': ['"=One!$A$2:$A$3"']}
#: None yet: a list that only shrinks.
ACCEPTED: dict[str, str] = {}


@pytest.mark.parametrize('case', cases(TARGETS, SETTINGS), ids=str)
def test_each_label_setting_comes_back(case: Case, tmp_path: Path) -> None:
    check(case, tmp_path, ACCEPTED, {})
