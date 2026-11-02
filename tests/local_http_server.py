"""A web server on 127.0.0.1 with fixed answers, which the HTTP probe and its replay both ask: no clock, no server name,
and nothing that changes between runs."""
from __future__ import annotations

import gzip
import socket
import threading
import time
from collections.abc import Generator
from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

TEXT = 'héllo wörld'
#: path: status, headers, body.
ANSWERS: dict[str, tuple[int, list[tuple[str, str]], bytes]] = {
    '/text': (200, [('Content-Type', 'text/plain; charset=utf-8'), ('X-Probe', 'yes')], TEXT.encode('utf-8')),
    '/latin1': (200, [('Content-Type', 'text/plain; charset=iso-8859-1')], TEXT.encode('latin-1')),
    '/plain-utf8': (200, [('Content-Type', 'text/plain')], TEXT.encode('utf-8')),
    '/plain-latin1': (200, [('Content-Type', 'text/plain')], TEXT.encode('latin-1')),
    '/untyped': (200, [], TEXT.encode('utf-8')),
    '/bom': (200, [('Content-Type', 'text/plain')], b'\xef\xbb\xbf' + TEXT.encode('utf-8')),
    '/utf16': (200, [('Content-Type', 'text/plain; charset=utf-16')], b'\xff\xfe' + TEXT.encode('utf-16-le')),
    '/utf16-bare': (200, [('Content-Type', 'text/plain')], b'\xff\xfe' + TEXT.encode('utf-16-le')),
    '/json': (200, [('Content-Type', 'application/json')], b'{"a": 1}'),
    '/binary': (200, [('Content-Type', 'application/octet-stream')], bytes(range(10))),
    '/empty': (204, [], b''),
    '/zero': (200, [('Content-Type', 'text/plain')], b''),
    '/missing': (404, [('Content-Type', 'text/plain')], b'nothing here'),
    '/broken': (500, [('Content-Type', 'text/plain')], b'it broke'),
    '/teapot': (418, [], b''),
    '/twice': (200, [('Content-Type', 'text/plain'), ('X-Twice', 'one'), ('X-Twice', 'two')], b'ok'),
    '/cookie': (200, [('Content-Type', 'text/plain'), ('Set-Cookie', 'probe=1; Path=/')], b'ok'),
    '/gzip': (200, [('Content-Type', 'text/plain'), ('Content-Encoding', 'gzip')], gzip.compress(b'zipped', mtime=0)),
    '/moved': (302, [('Location', '/text')], b'moved'),
    '/moved301': (301, [('Location', '/echo')], b'moved'),
    '/moved302': (302, [('Location', '/echo')], b'moved'),
    '/moved303': (303, [('Location', '/echo')], b'moved'),
    '/moved307': (307, [('Location', '/echo')], b'moved'),
    '/moved308': (308, [('Location', '/echo')], b'moved'),
    '/round': (302, [('Location', '/round')], b'again'),
    '/nowhere': (302, [], b'no location'),
    '/odd-charset': (200, [('Content-Type', 'text/plain; charset=nonsense')], TEXT.encode('utf-8')),
    '/quoted-charset': (200, [('Content-Type', 'text/plain; charset="iso-8859-1"')], TEXT.encode('latin-1')),
    '/many': (200, [('Server', 'probe'), ('Date', 'Mon, 02 Nov 2026 00:00:00 GMT'),
                    ('Set-Cookie', 'many=1; Path=/many'),
                    ('ETag', '"x"'), ('Cache-Control', 'no-store'), ('Expires', '0'), ('Vary', 'Accept'),
                    ('Last-Modified', 'Mon, 02 Nov 2026 00:00:00 GMT'), ('X-Zed', 'z'), ('Content-Language', 'en'),
                    ('Accept-Ranges', 'bytes'), ('Pragma', 'no-cache'), ('X-Alpha', 'a'),
                    ('Content-Type', 'text/plain'),
                    ('Connection', 'keep-alive')], b'ok'),
}


class _Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def log_message(self, format: str, *args: object) -> None:
        del format, args

    def _answer(self) -> None:
        length = int(self.headers.get('Content-Length') or 0)
        sent = self.rfile.read(length) if length else b''
        path = self.path.split('?', 1)[0]
        if path == '/echo':
            lines = [f'{self.command} {path} {self.request_version}']
            own = f'127.0.0.1:{self.connection.getsockname()[1]}'
            lines += [f'{name}: {value}'.replace(own, 'SERVER') for name, value in self.headers.items()]
            lines.append(f'BODY={sent.hex()}')
            status, headers, body = 200, [('Content-Type', 'text/plain')], '\n'.join(lines).encode('utf-8')
        elif path == '/auth':
            challenge = [('WWW-Authenticate', 'Basic realm="probe"'), ('Content-Type', 'text/plain')]
            given = self.headers.get('Authorization', '')
            status, headers, body = (200, [('Content-Type', 'text/plain')], given.encode()) if given else (
                401, challenge, b'who are you')
        elif path == '/slow':
            time.sleep(1.5)
            status, headers, body = 200, [('Content-Type', 'text/plain')], b'late'
        elif path == '/chunked':
            self.send_response_only(200)
            self.send_header('Content-Type', 'text/plain')
            self.send_header('Transfer-Encoding', 'chunked')
            self.end_headers()
            self.wfile.write(b'3\r\nabc\r\n2\r\nde\r\n0\r\n\r\n')
            return
        elif path.startswith('/path'):
            status, headers, body = 200, [('Content-Type', 'text/plain')], self.path.encode('latin-1')
        else:
            status, headers, body = ANSWERS.get(path, (404, [], b''))
        self.send_response_only(status)
        for name, value in headers:
            self.send_header(name, value)
        if status != 204:
            self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)

    do_GET = do_POST = do_PUT = do_DELETE = do_HEAD = do_PATCH = do_OPTIONS = do_PROBE = _answer


class _Server(ThreadingHTTPServer):
    daemon_threads = True

    def handle_error(self, request: object, client_address: object) -> None:
        """A client that hangs up on a kept-alive connection is no fault."""


@contextmanager
def local_server() -> Generator[tuple[str, int], None, None]:
    """The server's address, and a port on which nothing listens."""
    server = _Server(('127.0.0.1', 0), _Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    with socket.socket() as spare:
        spare.bind(('127.0.0.1', 0))
        closed = spare.getsockname()[1]
    try:
        yield f'http://127.0.0.1:{server.server_address[1]}', closed
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=5)
