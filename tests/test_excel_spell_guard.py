"""No formula a macro writes reaches spelled_formula's guard, "a formula the model cannot spell is not implemented"
(apps/excel/_model.py). formula/_spell.py refuses a structured token run straight into a ref, name or structured token, but
the parser refuses each such shape first, as error 1004, and a formula naming another workbook makes its link (see
_new_links). The guard stays for what a sync of the formula code from pyOfficeEditor may bring back.

tests/fixtures/verified_gaps_accepted.json accepts the guard on these grounds. If a shape here starts reaching it, this
test fails, and that entry has to be reconsidered rather than cover it.
"""
from __future__ import annotations

import pytest

from pyopenvba.apps.excel import ExcelApplication

#: Shapes the parser refuses: Agent 03's after #566, and more of the same kind.
REFUSED = ["=[Other]A1", "=[x]Name", "=Tbl[Col]A1", "=[#Data]A1", "=[Other.xlsx]Rate", "=[Other.xlsx]'My Sheet'!A1", "=SUM([x]A1:A2)",
           "=[Col][Col2]", "=[]Sheet1!A1", "=[a]b", "=[x][y]z", "=t[c]n", "=[x]'s'!A1", "=[@c]A1", "=[#All]A1", "=x[y]A1", "=[1]A1",
           "=[x]A1:B2", "=[x]Name+1", "=1+[x]A1", "=-[x]A1", "=([x]A1)", "=[x]A1&[x]B1", "=[x]#REF!", "=[x]R1C1", "=[[x]]A1"]
#: Shapes read as references to another workbook, which make its link.
LINKED = ["=[x]S!A1+[y]T!B1", "=[x]Sheet1:Sheet2!A1"]
MEMBERS = ["Formula", "Formula2", "FormulaLocal", "Value", "Value2"]


def _written(member: str, shape: str) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    quoted = shape.replace('"', '""')
    app.add_module(f'Public Function Probe() As String\nOn Error Resume Next\nActiveSheet.Range("B2").{member} = "{quoted}"\n'
                   'Probe = CStr(Err.Number)\nEnd Function\n', name="M")
    return str(app.run("M.Probe"))


@pytest.mark.parametrize("member", MEMBERS)
@pytest.mark.parametrize("shape", REFUSED)
def test_the_parser_refuses_first(member: str, shape: str) -> None:
    assert _written(member, shape) == "1004"


@pytest.mark.parametrize("member", MEMBERS)
@pytest.mark.parametrize("shape", LINKED)
def test_a_reference_to_another_workbook_makes_its_link(member: str, shape: str) -> None:
    assert _written(member, shape) == "0"
