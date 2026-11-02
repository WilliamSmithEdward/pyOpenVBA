"""HTTP from VBA, replayed from live Excel (scripts/measure_http_requests.py, tests/fixtures/http_requests.json) against
the same server on 127.0.0.1 (tests/local_http_server.py)."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from local_http_server import local_server
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter import VBASandboxError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/http_requests.json').read_text(encoding='utf-8'))
INET = ('MSXML2.XMLHTTP', 'MSXML2.XMLHTTP.6.0')
#: Cases the replay leaves out, each with its reason.
#: Excel hangs on it; the model refuses it.
HANGS = {'redirect put'}
#: XMLHTTP opens these schemes and fails on no server here; the model refuses them.
SCHEMES = {'open ftp', 'open file'}
#: A timing race in Excel: how far the request got in the time the reads took.
RACES = {(program, name) for program in ('MSXML2.ServerXMLHTTP.6.0', 'WinHttp.WinHttpRequest.5.1')
         for name in ('async spun', 'async at once')}
#: StrConv to a Variant holds a String of bytes in Excel, and a Byte array in the model: that gap is not HTTP's.
STRCONV = {'post strconv'}


def skipped(program: str, name: str) -> bool:
    return (program in INET and name in HANGS) or (program in INET + ('MSXML2.XMLHTTP',) and name in SCHEMES) or (
        (program, name) in RACES) or name in STRCONV


KEPT = [index for index, (program, name, _) in enumerate(RECORD['cases']) if not skipped(program, name)]


def masked(answer: str, base: str, refused: str) -> str:
    return re.sub(r'\?c=\d+', '?c=N', answer.replace(base, 'http://SERVER').replace(
        base.removeprefix('http://'), 'SERVER').replace(refused.removeprefix('http://'), 'REFUSED'))


@pytest.fixture(scope='module')
def answers() -> dict[int, str]:
    """Every case in the order Excel ran them, since XMLHTTP's cookies and logins last from one to the next."""
    with local_server() as (base, closed):
        refused = f'http://127.0.0.1:{closed}'
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(RECORD['source'], name='HttpProbe')
        found: dict[int, str] = {}
        for index in KEPT:
            try:
                found[index] = masked(str(app.run('HttpProbe.Probe', base, refused, index, index)), base, refused)
            except VBAUnsupportedError as refusal:
                found[index] = f'refused: {refusal}'
        return found


@pytest.mark.parametrize('index', KEPT, ids=[f'{RECORD["cases"][index][0]}-{RECORD["cases"][index][1]}' for index in KEPT])
def test_each_case_as_excel_answered(answers: dict[int, str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][2]


@pytest.mark.parametrize('program', ['MSXML2.XMLHTTP', 'MSXML2.XMLHTTP.6.0', 'MSXML2.ServerXMLHTTP.6.0',
                                     'WinHttp.WinHttpRequest.5.1'])
def test_the_sandbox_refuses_it_before_anything_goes_out(program: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    Set h = CreateObject("{program}")\nEnd Sub\n', name='Probe')
    with pytest.raises(VBASandboxError):
        app.run('Probe.Go')


@pytest.mark.parametrize('statements', [
    'h.Open "PUT", gBase & "/moved302", False: h.Send "abc"',
    'h.Open "GET", "file:///c:/nothing.txt", False: h.Send',
    'h.Open "GET", "ftp://127.0.0.1/x", False: h.Send',
    'h.Open "NO SPACE", gBase & "/text", False',
])
def test_what_was_not_measured_is_refused(statements: str) -> None:
    with local_server() as (base, _):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(f'Public Sub Go(gBase)\n    Set h = CreateObject("MSXML2.XMLHTTP")\n    {statements}\nEnd Sub\n',
                       name='Probe')
        with pytest.raises(VBAUnsupportedError):
            app.run('Probe.Go', base)


def test_an_asynchronous_request_waits_for_messages() -> None:
    """The usual loop: it ends at DoEvents, and with nothing but polling it would not."""
    source = ('Public Function Go(url)\n    Set h = CreateObject("MSXML2.XMLHTTP")\n    h.Open "GET", url, True\n'
              '    h.Send\n    Go = h.readyState\n    Do While h.readyState <> 4\n        DoEvents\n    Loop\n'
              '    Go = Go & " " & h.status & " " & h.responseText\nEnd Function\n')
    with local_server() as (base, _):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source, name='Probe')
        assert app.run('Probe.Go', f'{base}/json') == '1 200 {"a": 1}'
