"""HTTP from VBA: MSXML2.XMLHTTP, its later versions, ServerXMLHTTP and WinHttpRequest against a server on 127.0.0.1
(tests/local_http_server.py). Nothing leaves the machine.

Each case makes the object, runs its statements, each followed by any error's number and description, and reads READS
at the points the statements name. Text goes through Esc, since the harness's answer changes control characters; a
server's port is masked in the answers.

    python scripts/measure_http_requests.py

writes tests/fixtures/http_requests.json, which tests/test_vba_http.py replays.
"""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tests'))
from local_http_server import local_server  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/http_requests.json'
MAIN = ['MSXML2.XMLHTTP', 'MSXML2.XMLHTTP.6.0', 'MSXML2.ServerXMLHTTP.6.0', 'WinHttp.WinHttpRequest.5.1']
ALIASES = ['Microsoft.XMLHTTP', 'MSXML2.XMLHTTP.3.0', 'MSXML2.ServerXMLHTTP', 'MSXML2.ServerXMLHTTP.3.0']
READS = ['status', 'statusText', 'responseText', 'responseBody', 'getAllResponseHeaders()',
         'getResponseHeader("Content-Type")', 'getResponseHeader("content-type")', 'getResponseHeader("X-Missing")',
         'readyState']
SNAP = '@SNAP'


def get(path: str) -> list[str]:
    return [f'h.Open "GET", U("{path}"), False', 'h.Send', SNAP]


def send(method: str, body: str = '', *headers: str) -> list[str]:
    return [f'h.Open "{method}", U("/echo"), False', *headers, f'h.Send{" " + body if body else ""}', SNAP]


BODIES: dict[str, list[str]] = {
    'made': [SNAP],
    'opened': ['h.Open "GET", U("/text"), False', SNAP],
    **{f'get {path}': get(path) for path in (
        '/text', '/latin1', '/plain-utf8', '/plain-latin1', '/untyped', '/bom', '/utf16', '/utf16-bare', '/json',
        '/binary', '/empty', '/zero', '/missing', '/broken', '/teapot', '/twice', '/gzip', '/moved', '/round',
        '/nowhere', '/echo')},
    'head': ['h.Open "HEAD", U("/text"), False', 'h.Send', SNAP],
    'get lower': ['h.Open "get", U("/echo"), False', 'h.Send', SNAP],
    'get with a body': send('GET', '"abc"'),
    'post plain': send('POST'),
    'post text': send('POST', '"abc"'),
    'post empty text': send('POST', '""'),
    'post accented': send('POST', 'ChrW(233) & ChrW(8364)'),
    'post number': send('POST', '12'),
    'post bytes': ['Dim b(2) As Byte', 'b(0) = 1: b(1) = 200: b(2) = 255', 'h.Open "POST", U("/echo"), False', 'h.Send b',
                   SNAP],
    'post typed': send('POST', '"abc"', 'h.setRequestHeader "Content-Type", "application/json"'),
    'post typed charset': send('POST', 'ChrW(233)', 'h.setRequestHeader "Content-Type", "text/plain; charset=iso-8859-1"'),
    'put': send('PUT', '"abc"'),
    'delete': send('DELETE'),
    'patch': send('PATCH', '"abc"'),
    'options': send('OPTIONS'),
    'custom method': send('PROBE', '"abc"'),
    'header': send('GET', '', 'h.setRequestHeader "X-One", "1"'),
    'header twice': send('GET', '', 'h.setRequestHeader "X-One", "1"', 'h.setRequestHeader "X-One", "2"'),
    'header empty': send('GET', '', 'h.setRequestHeader "X-One", ""'),
    'header accept': send('GET', '', 'h.setRequestHeader "Accept", "text/plain"'),
    'header agent': send('GET', '', 'h.setRequestHeader "User-Agent", "probe"'),
    'header host': send('GET', '', 'h.setRequestHeader "Host", "example.invalid"'),
    'header length': send('POST', '"abc"', 'h.setRequestHeader "Content-Length", "3"'),
    'header bad name': send('GET', '', 'h.setRequestHeader "X One", "1"'),
    'header empty name': send('GET', '', 'h.setRequestHeader "", "1"'),
    'header before open': ['h.setRequestHeader "X-One", "1"', SNAP],
    'header after send': ['h.Open "GET", U("/text"), False', 'h.Send', 'h.setRequestHeader "X-One", "1"', SNAP],
    'send before open': ['h.Send', SNAP],
    'send twice': ['h.Open "GET", U("/text"), False', 'h.Send', 'h.Send', SNAP],
    'open again': ['h.Open "GET", U("/text"), False', 'h.Send', SNAP, 'h.Open "GET", U("/json"), False', SNAP, 'h.Send',
                   SNAP],
    'abort': ['h.Open "GET", U("/text"), False', 'h.Send', 'h.Abort', SNAP],
    'abort before': ['h.Abort', SNAP],
    'open no url': ['h.Open "GET", "", False', SNAP, 'h.Send', SNAP],
    'open bad url': ['h.Open "GET", "not a url", False', SNAP, 'h.Send', SNAP],
    'open relative': ['h.Open "GET", "/text", False', SNAP, 'h.Send', SNAP],
    'open ftp': ['h.Open "GET", "ftp://127.0.0.1/x", False', SNAP, 'h.Send', SNAP],
    'open file': ['h.Open "GET", "file:///c:/nothing.txt", False', SNAP, 'h.Send', SNAP],
    'open empty method': ['h.Open "", U("/text"), False', SNAP, 'h.Send', SNAP],
    'refused': ['h.Open "GET", gRefused & "/text", False', SNAP, 'h.Send', SNAP],
    'unresolved': ['h.Open "GET", "http://name.invalid/text", False', SNAP, 'h.Send', SNAP],
    'path': ['h.Open "GET", gBase & "/path?a=b c&d=" & ChrW(233) & "&e=%41#frag", False', 'h.Send', SNAP],
    'path spaced': ['h.Open "GET", gBase & "/path/a b", False', 'h.Send', SNAP],
    'async': ['h.Open "GET", U("/text"), True', 'h.Send', 'Settle h', SNAP],
    'async left out': ['h.Open "GET", U("/text")', 'h.Send', 'Settle h', SNAP],
    'async at once': ['h.Open "GET", U("/text"), True', 'h.Send', SNAP, 'Settle h', SNAP],
    'async spun': ['h.Open "GET", U("/text"), True', 'h.Send', 'Spin h', SNAP, 'Settle h', SNAP],
    'post bytes in a variant': ['Dim b(2) As Byte', 'Dim v As Variant', 'b(0) = 1: b(1) = 200: b(2) = 255', 'v = b',
                                'h.Open "POST", U("/echo"), False', 'h.Send v', SNAP],
    'post strconv': ['Dim v As Variant', 'v = StrConv("abc", vbFromUnicode)', 'h.Open "POST", U("/echo"), False',
                     'h.Send v', SNAP],
    'auth given': ['h.Open "GET", U("/auth"), False, "user", "word"', 'h.Send', SNAP],
    **{f'get {path}': get(path) for path in ('/many', '/chunked', '/odd-charset', '/quoted-charset')},
    'post dynamic bytes': ['Dim b() As Byte', 'ReDim b(2)', 'b(0) = 1: b(1) = 200: b(2) = 255',
                           'h.Open "POST", U("/echo"), False', 'h.Send b', SNAP],
    'post strconv bytes': ['Dim b() As Byte', 'b = StrConv("abc", vbFromUnicode)', 'h.Open "POST", U("/echo"), False',
                           'h.Send b', SNAP],
    'post ascii made': send('POST', 'Chr(97) & "bc"'),
    'post variant text': ['Dim v As Variant', 'v = "abc"', 'h.Open "POST", U("/echo"), False', 'h.Send v', SNAP],
    'post latin': send('POST', 'ChrW(233)'),
    'post typed accented': send('POST', 'ChrW(233)', 'h.setRequestHeader "Content-Type", "application/json"'),
    'delete with a body': send('DELETE', '"abc"'),
    'head with a body': send('HEAD', '"abc"'),
    'post header': send('POST', '"abc"', 'h.setRequestHeader "X-One", "1"'),
    'post header typed': send('POST', '"abc"', 'h.setRequestHeader "X-One", "1"',
                              'h.setRequestHeader "Content-Type", "text/csv"'),
    'header language': send('GET', '', 'h.setRequestHeader "Accept-Language", "fr"'),
    'header encoding': send('GET', '', 'h.setRequestHeader "Accept-Encoding", "identity"'),
    'header connection': send('GET', '', 'h.setRequestHeader "Connection", "close"'),
    'header cookie': send('GET', '', 'h.setRequestHeader "Cookie", "given=1"'),
    'header case': send('GET', '', 'h.setRequestHeader "accept", "text/plain"', 'h.setRequestHeader "x-low", "1"'),
    'read twice': ['h.Open "GET", U("/twice"), False', 'h.Send', 's = s & "<" & Esc(h.getResponseHeader("X-Twice")) & ">"'],
    'async refused': ['h.Open "GET", gRefused & "/text", True', 'h.Send', 'Settle h', SNAP],
    'wait on it': ['h.Open "GET", U("/text"), False', 'h.Send', 's = s & "<" & h.waitForResponse(1) & ">"'],
    'set timeouts': ['h.setTimeouts 1000, 1000, 1000, 1000', 'h.Open "GET", U("/text"), False', 'h.Send', SNAP],
    'option': ['s = s & "<" & h.Option(0) & ">"'],
    'path with a space': ['h.Open "GET", gBase & "/path/a b#c", False', 'h.Send', SNAP],
    'do events': ['s = s & "<" & DoEvents & ">"'],
    'cookie kept': ['h.Open "GET", U("/cookie"), False', 'h.Send', 'h.Open "GET", U("/echo"), False', 'h.Send', SNAP],
    'credentials': ['h.Open "GET", U("/echo"), False, "user", "word"', 'h.Send', SNAP],
    **{f'redirect post {code}': ['h.Open "POST", U("/moved' + code + '"), False', 'h.Send "abc"', SNAP]
       for code in ('301', '302', '303', '307', '308')},
    'redirect put': ['h.Open "PUT", U("/moved302"), False', 'h.Send "abc"', SNAP],
    'status names': ['s = s & "<" & TypeName(h) & ">"', 'h.Open "GET", U("/text"), False', 'h.Send',
                     's = s & "<" & TypeName(h.status) & TypeName(h.statusText) & TypeName(h.responseText) & ">"'],
}
WINHTTP_ONLY: dict[str, list[str]] = {
    'too slow': ['h.SetTimeouts 500, 500, 500, 500', 'h.Open "GET", U("/slow"), False', 'h.Send', SNAP],
    'set credentials server': ['h.Open "GET", U("/auth"), False', 'h.SetCredentials "user", "word", 0', 'h.Send', SNAP],
    'set credentials before': ['h.SetCredentials "user", "word", 0', SNAP],
    'timeouts': ['h.SetTimeouts 5000, 5000, 5000, 5000', 'h.Open "GET", U("/text"), False', 'h.Send', SNAP],
    'wait': ['h.Open "GET", U("/text"), True', 'h.Send', 's = s & "<" & h.WaitForResponse(5) & ">"', SNAP],
    'wait sync': ['h.Open "GET", U("/text"), False', 'h.Send', 's = s & "<" & h.WaitForResponse(5) & ">"', SNAP],
    'wait none': ['s = s & "<" & h.WaitForResponse(1) & ">"', SNAP],
    'option agent': ['s = s & "<" & h.Option(0) & ">"', 'h.Option(0) = "probe"', 'h.Open "GET", U("/echo"), False', 'h.Send', SNAP],
    'option url': ['h.Open "GET", U("/moved"), False', 'h.Send', 's = s & "<" & Esc(h.Option(1)) & ">"'],
    'option page': ['s = s & "<" & h.Option(2) & ">"'],
    'option redirects': ['s = s & "<" & h.Option(6) & ">"', 'h.Option(6) = False', 'h.Open "GET", U("/moved"), False', 'h.Send', SNAP],
}
SERVER_ONLY: dict[str, list[str]] = {
    'too slow': ['h.setTimeouts 500, 500, 500, 500', 'h.Open "GET", U("/slow"), False', 'h.Send', SNAP],
    'timeouts': ['h.setTimeouts 5000, 5000, 5000, 5000', 'h.Open "GET", U("/text"), False', 'h.Send', SNAP],
    'wait': ['h.Open "GET", U("/text"), True', 'h.Send', 's = s & "<" & h.waitForResponse(5) & ">"', SNAP],
}
CASES = [(program, name, body) for program in MAIN for name, body in BODIES.items()]
CASES += [(program, 'get /text', BODIES['get /text']) for program in ALIASES]
CASES += [(program, 'get /echo', BODIES['get /echo']) for program in ALIASES]
CASES += [('WinHttp.WinHttpRequest.5.1', name, body) for name, body in WINHTTP_ONLY.items()]
CASES += [('MSXML2.ServerXMLHTTP.6.0', name, body) for name, body in SERVER_ONLY.items()]

_HELPERS = r'''
Private gBase As String, gRefused As String, gCase As Long

Private Function U(ByVal path As String) As String
    U = gBase & path & "?c=" & gCase
End Function

Private Function Esc(ByVal text As Variant) As String
    Dim i As Long, code As Long, out As String
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1)) And &HFFFF&
        If code = 13 Then
            out = out & "\r"
        ElseIf code = 10 Then
            out = out & "\n"
        ElseIf code < 32 Or code > 126 Or code = 92 Or code = 124 Or code = 64 Then
            out = out & "\u" & Right$("000" & Hex$(code), 4)
        Else
            out = out & Mid$(text, i, 1)
        End If
    Next
    Esc = out
End Function

Private Function Bytes(ByVal value As Variant) As String
    Dim i As Long, out As String, first As Long
    If Not IsArray(value) Then
        Bytes = TypeName(value)
        Exit Function
    End If
    On Error Resume Next
    first = LBound(value)
    If Err.Number Then
        Bytes = TypeName(value) & " unbound"
        Exit Function
    End If
    On Error GoTo 0
    out = TypeName(value) & " " & LBound(value) & " " & UBound(value) & " "
    For i = LBound(value) To UBound(value)
        out = out & Right$("0" & Hex$(value(i)), 2)
    Next
    Bytes = out
End Function

Private Function Failure(ByVal number As Long, ByVal description As String) As String
    Failure = "E" & number & ":" & Esc(description)
    Err.Clear
End Function

Private Sub Settle(ByVal h As Object)
    Dim start As Single
    On Error Resume Next
    start = Timer
    Do While Timer - start < 3
        DoEvents
        h.WaitForResponse 1
        Err.Clear
        If h.readyState = 4 Then Exit Sub
        If Err.Number Then Exit Sub
    Loop
End Sub

Private Sub Spin(ByVal h As Object)
    Dim start As Single
    On Error Resume Next
    start = Timer
    Do While Timer - start < 2
        If h.readyState = 4 Then Exit Sub
        If Err.Number Then Exit Sub
    Loop
End Sub

Private Function Snap(ByVal h As Object) As String
    Dim s As String, v As Variant
    On Error Resume Next
    Err.Clear: v = h.status: If Err.Number Then v = Failure(Err.Number, Err.Description)
    s = "|status=" & v
    Err.Clear: v = h.statusText: If Err.Number Then v = Failure(Err.Number, Err.Description) Else v = Esc(v)
    s = s & "|text=" & v
    Err.Clear: v = h.responseText: If Err.Number Then v = Failure(Err.Number, Err.Description) Else v = Esc(v)
    s = s & "|response=" & v
    Err.Clear: v = h.responseBody: If Err.Number Then v = Failure(Err.Number, Err.Description) Else v = Bytes(v)
    s = s & "|body=" & v
    Err.Clear: v = h.getAllResponseHeaders(): If Err.Number Then v = Failure(Err.Number, Err.Description) Else v = Esc(v)
    s = s & "|all=" & v
    Err.Clear: v = h.getResponseHeader("Content-Type"): If Err.Number Then v = Failure(Err.Number, Err.Description) Else v = Esc(v)
    s = s & "|type=" & v
    Err.Clear: v = h.getResponseHeader("content-type"): If Err.Number Then v = Failure(Err.Number, Err.Description) Else v = Esc(v)
    s = s & "|lower=" & v
    Err.Clear: v = h.getResponseHeader("X-Missing"): If Err.Number Then v = Failure(Err.Number, Err.Description) Else v = Esc(v)
    s = s & "|missing=" & v
    Err.Clear: v = h.readyState: If Err.Number Then v = Failure(Err.Number, Err.Description)
    s = s & "|state=" & v
    Err.Clear
    Snap = s
End Function
'''


def _statement(text: str) -> list[str]:
    if text == SNAP:
        return ['    s = s & Snap(h)']
    if text.startswith('Dim '):
        return [f'    {text}']
    return [f'    Err.Clear: {text}: If Err.Number Then s = s & "[" & Failure(Err.Number, Err.Description) & "]"']


def _probe() -> str:
    lines = [_HELPERS]
    for index, (program, _, body) in enumerate(CASES):
        lines += [f'Private Function Case{index}() As String', '    Dim h As Object, s As String',
                  '    On Error Resume Next', f'    Set h = CreateObject("{program}")',
                  '    If Err.Number Then', '        Case' + str(index) + ' = "cannot make: " & Failure(Err.Number, Err.Description)',
                  '        Exit Function', '    End If']
        for text in body:
            lines += _statement(text)
        lines += [f'    Case{index} = s', 'End Function', '']
    lines += ['Private Function Tried(ByVal which As Long) As String', '    gCase = which', '    Select Case which']
    for index in range(len(CASES)):
        lines += [f'        Case {index}: Tried = Case{index}()']
    lines += ['    End Select', 'End Function', '',
              'Public Function Probe(ByVal base As String, ByVal refused As String, ByVal first As Long, '
              'ByVal last As Long) As String', '    Dim i As Long, s As String', '    gBase = base: gRefused = refused',
              '    For i = first To last', '        s = s & "@@@" & Tried(i)', '    Next', '    Probe = Mid$(s, 4)',
              'End Function']
    return '\n'.join(lines) + '\n'


def masked(answer: str, base: str, refused: str) -> str:
    """The answer with what changes between runs taken out: the server's port and the case's query."""
    return re.sub(r'\?c=\d+', '?c=N', answer.replace(base, 'http://SERVER').replace(
        base.removeprefix('http://'), 'SERVER').replace(refused.removeprefix('http://'), 'REFUSED'))


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with local_server() as (base, closed), ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        refused = f'http://127.0.0.1:{closed}'
        excel.new_document()
        for index, (program, name, _) in enumerate(CASES):
            # One at a time: a case that hangs Excel is reported by name and answered as hung, and the rest still run.
            result = excel.run_vba(source, 'Probe', (base, refused, index, index), timeout=40.0)
            if not result.ok or result.dialogs:
                print(f'HUNG {index} {program} / {name}: {result.outcome} {result.dialogs}', flush=True)
                answers.append('hung')
                excel.new_document()
                continue
            answers.append(masked(str(result.value), base, refused))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[program, name, answer] for (program, name, _), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for program, name, answer in record['cases']:
        print(f'{program} / {name}\n   {answer}')


if __name__ == '__main__':
    main()
