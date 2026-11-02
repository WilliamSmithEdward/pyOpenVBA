"""HTTP from VBA without COM: MSXML2.XMLHTTP, MSXML2.ServerXMLHTTP and WinHttp.WinHttpRequest on Python's own sockets, as
tests/fixtures/http_requests.json measured them against a server on 127.0.0.1.

They are three stacks underneath, and it shows. XMLHTTP is WinINet: it adds a browser's headers, unpacks gzip, answers a
204 with status 1223, and shares its cookies and its Basic logins across every object in the process. ServerXMLHTTP and
WinHttpRequest are WinHTTP: their headers go out and come back in WinHTTP's own order, and each object keeps its own
cookies. WinHttpRequest alone decodes text with no charset as Latin-1, and adds a Content-Type to plain text.

An asynchronous request finishes when VBA pumps messages, at DoEvents or WaitForResponse: a loop that only polls never
sees it end, as in Excel. A CreateObject of any of them is refused with the sandbox on.

Not modelled: WinINet's disk cache, proxies, Digest and Windows logins. The measurements did not reach TLS failures, so
those raise VBAUnsupportedError, as do the file: and ftp: schemes XMLHTTP would open and PUT through a redirect, which
hangs XMLHTTP in Excel."""
from __future__ import annotations

import base64
import codecs
import gzip
import http.client
import socket
import ssl
import zlib
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Final
from urllib.parse import urljoin, urlsplit

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NULL,
    VBAArray,
    VBAInt,
    error,
    to_bool,
    to_integer,
    to_number,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter

_UNSPECIFIED = (-2147467259, "Unspecified error\r\n")
_PENDING = (-2147483638, "The data necessary to complete this operation is not yet available.\r\n")
_PARAMETER = (-2147024809, "The parameter is incorrect.\r\n")
_ARGUMENT = (5, "Invalid procedure call or argument")
_XML_OPEN = (-1072954818, "This method cannot be called until the open method has been called.\r\n")
_XML_SEND = (-1072954817, "This method cannot be called until the send method has been called.\r\n")
_XML_SENT = (-1072954816, "This method cannot be called after the send method has been called.\r\n")
_HTTP_OPEN = (-2147012796, "This method cannot be called until the Open method has been called\r\n")
_HTTP_SEND = (-2147012795, "This method cannot be called until the Send method has been called\r\n")
_HTTP_SENT = (-2147012794, "This method cannot be called after the Send method has been called\r\n")
_NO_HEADER = (-2147012746, "The requested header was not found\r\n")
_WRONG_STATE = (-2147012877, "The handle is in the wrong state for the requested operation\r\n")
_REDIRECT = (-2147012740, "The HTTP redirect request failed\r\n")
_REFUSED = (-2147012867, "A connection with the server could not be established\r\n")
_UNRESOLVED = (-2147012889, "The server name or address could not be resolved\r\n")
_TIMED_OUT = (-2147012894, "The operation timed out\r\n")
_INET_RESOURCE = (-2146697211, "The system cannot locate the resource specified.\r\n")
_INET_REDIRECT = (-2146697196, "A redirection problem occurred.\r\n")
_INET_DOWNLOAD = (-2146697208, "The download of the specified resource has failed.\r\n")
_INET_ABORTED = (-2147467260, "Operation aborted\r\n")

_Failure = tuple[int, str]


@dataclass(frozen=True)
class _Stack:
    """What one program answers: its type's name, the errors of a step out of turn, and how it treats a request."""

    type_name: str
    inet: bool
    #: Before Open, after Open, after Send started, after Send failed: the error of a status, of the response itself,
    #: and of the headers.
    new: tuple[_Failure, _Failure, _Failure]
    opened: tuple[_Failure, _Failure, _Failure]
    waiting: tuple[_Failure, _Failure, _Failure]
    needs_open: _Failure
    sent_twice: _Failure | None
    header_after_send: _Failure
    open_blank: _Failure
    open_bad_url: _Failure
    bad_header: _Failure
    bad_charset: _Failure
    has_ready_state: bool = True
    async_by_default: bool = False


_XML30 = _Stack("IServerXMLHTTPRequest2", True, (_UNSPECIFIED, _PENDING, _UNSPECIFIED),
                (_UNSPECIFIED, _UNSPECIFIED, _UNSPECIFIED), (_UNSPECIFIED, _PENDING, _UNSPECIFIED), _UNSPECIFIED,
                _UNSPECIFIED, _UNSPECIFIED, _ARGUMENT,
                (-2147012890, "Method 'open' of object 'IServerXMLHTTPRequest2' failed"), _PARAMETER,
                (-1072896658, "System does not support the specified encoding.\r\n"), async_by_default=True)
_XML60 = _Stack("IServerXMLHTTPRequest2", True, (_XML_SEND, _XML_SEND, _XML_SEND), (_XML_SEND, _XML_SEND, _XML_SEND),
                (_PENDING, _PENDING, _PENDING), _XML_OPEN, _XML_SENT, _XML_SENT, _PARAMETER,
                (-2147012890, "System error: -2147012890.\r\n"), _PARAMETER,
                (-1072896658, "System error: -1072896658.\r\n"), async_by_default=True)
_SERVER = _Stack("IServerXMLHTTPRequest2", False, (_XML_OPEN, _XML_SEND, _XML_OPEN), (_HTTP_SEND, _XML_SEND, _HTTP_SEND),
                 (_PENDING, _PENDING, _WRONG_STATE), _XML_OPEN, _XML_SENT, _XML_SENT, _PARAMETER,
                 (-2147012890, "System error: -2147012890.\r\n"), _PARAMETER,
                 (-1072896658, "System does not support the specified encoding.\r\n"))
_WINHTTP = _Stack("WinHttpRequest", False, (_HTTP_SEND, _HTTP_SEND, _HTTP_SEND), (_HTTP_SEND, _HTTP_SEND, _HTTP_SEND),
                  (_PENDING, _PENDING, _WRONG_STATE), _HTTP_OPEN, None, _HTTP_SENT, _ARGUMENT,
                  (-2147012890, "The URL does not use a recognized protocol\r\n"), _ARGUMENT,
                  (-2147023783, "No mapping for the Unicode character exists in the target multi-byte code page.\r\n"),
                  has_ready_state=False)

#: The programs measured, by CreateObject's name.
PROGRAMS: Final[dict[str, _Stack]] = {
    "microsoft.xmlhttp": _XML30, "msxml2.xmlhttp": _XML30, "msxml2.xmlhttp.3.0": _XML30, "msxml2.xmlhttp.6.0": _XML60,
    "msxml2.serverxmlhttp": _SERVER, "msxml2.serverxmlhttp.3.0": _SERVER, "msxml2.serverxmlhttp.6.0": _SERVER,
    "winhttp.winhttprequest.5.1": _WINHTTP,
}

_INET_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; Trident/7.0; rv:11.0) like Gecko"
_HTTP_AGENT = "Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)"
_LANGUAGE = "en-US,en;q=0.5"
_STANDARD_METHODS = {"get", "post", "put", "delete", "head", "options", "patch", "trace"}
#: The order WinHTTP lists a response's headers in; any other follows them as it came.
_WINHTTP_ORDER = ("cache-control", "connection", "date", "pragma", "transfer-encoding", "content-length", "content-type",
                  "content-encoding", "content-language", "expires", "last-modified", "accept-ranges", "etag",
                  "location", "server", "set-cookie", "vary", "www-authenticate")
_REDIRECTS = {301, 302, 303, 307, 308}
_HOPS = 10
#: Seconds, where the program's own timeout is not set or is 0: no call here waits for ever.
_CONNECT_WAIT = 60.0
_RECEIVE_WAIT = 30.0


@dataclass
class _Response:
    status: int
    reason: str
    headers: list[tuple[str, str]]
    body: bytes


@dataclass
class Sessions:
    """What WinINet keeps for the whole process: cookies and logins, by host."""

    cookies: dict[str, dict[str, tuple[str, str]]] = field(default_factory=lambda: {})
    logins: dict[str, str] = field(default_factory=lambda: {})


def _msxml_utf8(data: bytes) -> str:
    """UTF-8 as MSXML reads it: a lead byte takes as many bytes as it says whatever they are, and a sequence that does
    not hold together, or a byte that leads nothing, is the one character U+FFFF."""
    out: list[str] = []
    index = 0
    while index < len(data):
        byte = data[index]
        length = 1 if byte < 0x80 else 2 if 0xC0 <= byte < 0xE0 else 3 if 0xE0 <= byte < 0xF0 else 4 if 0xF0 <= byte else 0
        if length == 1:
            out.append(chr(byte))
            index += 1
        elif length == 0:
            out.append("￿")
            index += 1
        else:
            piece = data[index:index + length]
            try:
                out.append(piece.decode("utf-8"))
            except UnicodeDecodeError:
                out.append("￿")
            index += length
    return "".join(out)


def _charset(headers: list[tuple[str, str]]) -> str | None:
    kind = next((value for name, value in headers if name.lower() == "content-type"), "")
    for part in kind.split(";")[1:]:
        key, _, value = part.partition("=")
        if key.strip().lower() == "charset":
            return value.strip().strip('"')
    return None


def _codec(name: str) -> codecs.CodecInfo | None:
    try:
        return codecs.lookup(name)
    except LookupError:
        return None


def _keeps_cookie(path: str, wanted: str) -> bool:
    return wanted == path or wanted.startswith(path.rstrip("/") + "/")


class HttpRequest(VBAObject):
    """An HTTP request object of any of the programs in PROGRAMS, with the sandbox off."""

    vba_library = ""
    #: Send needs to know a byte array's variable: see _body.
    vba_slot_arguments = frozenset({"send"})

    def __init__(self, stack: _Stack, interpreter: Interpreter) -> None:
        self.stack, self.interpreter = stack, interpreter
        self.vba_type_name = stack.type_name
        self.stage = "new"
        self.response: _Response | None = None
        self.failed_headers: list[tuple[str, str]] = []
        self.method, self.url = "", ""
        self.is_async = False
        self.user: tuple[str, str] | None = None
        self.login: tuple[str, str] | None = None
        self.headers: list[tuple[str, str]] = []
        self.pending: tuple[bytes, bool] = (b"", False)
        self.cookies: dict[str, dict[str, tuple[str, str]]] = {}
        self.agent = _HTTP_AGENT
        self.follows = True
        self.final_url = ""
        self.connect_wait, self.receive_wait = _CONNECT_WAIT, _RECEIVE_WAIT

    # -- reads

    def _read(self, which: int) -> None:
        """The error of a read at this stage: 0 a status, 1 the response, 2 the headers."""
        stages = {"new": self.stack.new, "opened": self.stack.opened, "aborted": self.stack.new,
                  "waiting": self.stack.waiting, "failed": self.stack.waiting}
        if self.stage == "aborted" and self.stack is _XML30:
            raise error(*self.stack.opened[which])
        if self.stage in stages:
            raise error(*stages[self.stage][which])

    def _done(self) -> _Response:
        assert self.response is not None
        return self.response

    @member
    def Status(self) -> object:
        self._read(0)
        return VBAInt(self._done().status, "Long")

    @member
    def StatusText(self) -> object:
        self._read(0)
        return self._done().reason

    @member
    def ResponseBody(self) -> object:
        self._read(1)
        response = self._done()
        if response.status == 1223 or (self.stack.inet and response.status in (0, 2, 12007, 12029, 12150)):
            return EMPTY
        if not response.body and not self.stack.inet:
            return EMPTY
        return VBAArray([(0, len(response.body) - 1)], element_type="Byte",
                        items=[VBAInt(byte, "Byte") for byte in response.body])

    @member
    def ResponseText(self) -> object:
        self._read(1)
        response = self._done()
        charset = _charset(response.headers)
        return self._winhttp_text(response.body, charset) if self.stack is _WINHTTP else self._msxml_text(
            response.body, charset)

    def _bad_charset(self) -> Exception:
        return error(*self.stack.bad_charset)

    def _msxml_text(self, data: bytes, charset: str | None) -> str:
        if charset is not None:
            info = _codec(charset)
            if info is None:
                raise self._bad_charset()
            if info.name != "utf-8":
                return data.decode(info.name, errors="replace")
        if data.startswith(b"\xef\xbb\xbf"):
            return _msxml_utf8(data[3:])
        if data.startswith(b"\xff\xfe"):
            return data[2:].decode("utf-16-le", errors="replace")
        if data.startswith(b"\xfe\xff"):
            return data[2:].decode("utf-16-be", errors="replace")
        return _msxml_utf8(data)

    def _winhttp_text(self, data: bytes, charset: str | None) -> str:
        if charset is None:
            return data.decode("latin-1")
        info = _codec(charset)
        if info is None:
            raise self._bad_charset()
        # A byte order mark stays in the text.
        name = "utf-16-be" if info.name == "utf-16" and data.startswith(b"\xfe\xff") else (
            "utf-16-le" if info.name == "utf-16" else info.name)
        return data.decode(name, errors="replace")

    def _header_list(self) -> list[tuple[str, str]]:
        if self.stage == "failed":
            return self.failed_headers
        return self._done().headers

    def _read_headers(self) -> None:
        if self.stage == "failed" and self.failed_headers:
            return
        self._read(2)

    @method
    def GetAllResponseHeaders(self) -> object:
        self._read_headers()
        headers = self._header_list()
        if self.stack.inet:
            return "".join(f"{name}: {value}\r\n" for name, value in headers) + ("\r\n" if headers else "")
        known = {name: place for place, name in enumerate(_WINHTTP_ORDER)}
        ordered = sorted(headers, key=lambda one: known.get(one[0].lower(), len(known)))
        return "".join(f"{name}: {value}\r\n" for name, value in ordered) + "\r\n"

    @method
    def GetResponseHeader(self, Header: object = MISSING) -> object:
        self._read_headers()
        wanted = to_text(Header).lower()
        found = next((value for name, value in self._header_list() if name.lower() == wanted), None)
        if found is not None:
            return found
        if self.stack is _WINHTTP:
            raise error(*_NO_HEADER)
        return ""

    @member
    def ReadyState(self) -> object:
        if not self.stack.has_ready_state:
            raise error(438)
        state = {"new": 0, "aborted": 0, "opened": 1, "waiting": 1, "failed": 1}.get(self.stage, 4)
        return VBAInt(state, "Long")

    # -- the steps

    @method
    def Open(self, Method: object = MISSING, Url: object = MISSING, Async: object = MISSING, User: object = MISSING,
             Password: object = MISSING) -> object:
        verb, address = to_text(Method), to_text(Url)
        if not verb or not address:
            self._forget()
            raise error(*self.stack.open_blank)
        if not verb.isascii() or " " in verb:
            raise VBAUnsupportedError("An HTTP method that is not a single word is not implemented")
        scheme = address.partition("://")[0].lower() if "://" in address else ""
        if scheme not in ("http", "https", "ftp", "file") or (not self.stack.inet and scheme not in ("http", "https")):
            self._forget()
            raise error(*self.stack.open_bad_url)
        self._drop_pending()
        self.method = verb.upper() if verb.lower() in _STANDARD_METHODS else verb
        self.url = address
        self.is_async = self.stack.async_by_default if Async is MISSING else to_bool(Async)
        self.user = None
        if User is not MISSING and self.stack is not _WINHTTP:
            self.user = (to_text(User), "" if Password is MISSING else to_text(Password))
        self.login = None
        self.headers = []
        self.response, self.failed_headers = None, []
        self.stage = "opened"
        return EMPTY

    def _forget(self) -> None:
        self._drop_pending()
        self.stage, self.response, self.failed_headers = "new", None, []

    def _drop_pending(self) -> None:
        if self._pump in self.interpreter.pumped:
            self.interpreter.pumped.remove(self._pump)

    @method
    def SetRequestHeader(self, Header: object = MISSING, Value: object = MISSING) -> object:
        name, value = to_text(Header), to_text(Value)
        if self.stage in ("new", "aborted"):
            raise error(*self.stack.needs_open)
        if self.stage != "opened":
            raise error(*self.stack.header_after_send)
        if "\r" in name or "\n" in name or "\r" in value or "\n" in value or not (name + value).isascii():
            raise VBAUnsupportedError("An HTTP header with a line break or beyond ASCII is not implemented")
        if not name or (" " in name and not self.stack.inet):
            raise error(*self.stack.bad_header)
        if not value and not self.stack.inet:
            raise error(*_NO_HEADER)
        self.headers.append((name, value))
        return EMPTY

    @method
    def SetTimeouts(self, ResolveTimeout: object = MISSING, ConnectTimeout: object = MISSING,
                    SendTimeout: object = MISSING, ReceiveTimeout: object = MISSING) -> object:
        if self.stack.inet:
            raise error(438)
        connect, receive = float(to_number(ConnectTimeout)) / 1000, float(to_number(ReceiveTimeout)) / 1000
        del ResolveTimeout, SendTimeout
        # WinHttpRequest set this way before Open went on to wait out 1.5 seconds on 0.5 (tests/fixtures).
        if self.stack is _WINHTTP and self.stage == "new":
            return EMPTY
        self.connect_wait = connect if connect > 0 else _CONNECT_WAIT
        self.receive_wait = receive if receive > 0 else _RECEIVE_WAIT
        return EMPTY

    @method
    def SetCredentials(self, UserName: object = MISSING, Password: object = MISSING, Flags: object = MISSING) -> object:
        if self.stack is not _WINHTTP:
            raise error(438)
        if self.stage in ("new", "aborted"):
            raise error(*self.stack.needs_open)
        if int(to_integer(Flags, "Long")) != 0:
            raise VBAUnsupportedError("WinHttpRequest.SetCredentials for a proxy is not implemented")
        self.user = (to_text(UserName), to_text(Password))
        return EMPTY

    @method
    def Send(self, Body: object = MISSING) -> object:
        if self.stage in ("new", "aborted"):
            raise error(*self.stack.needs_open)
        if self.stage != "opened" and self.stack.sent_twice is not None:
            raise error(*self.stack.sent_twice)
        body, is_text = self._body(Body)
        self._check_scheme()
        self.response, self.failed_headers = None, []
        if self.is_async:
            self.stage = "waiting"
            self.pending = (body, is_text)
            self.interpreter.pumped.append(self._pump)
            return EMPTY
        failure = self._finish(body, is_text)
        if failure is not None:
            raise error(*failure)
        return EMPTY

    def _check_scheme(self) -> None:
        scheme = urlsplit(self.url).scheme.lower()
        if scheme not in ("http", "https"):
            raise VBAUnsupportedError(f"XMLHTTP over {scheme}: is not implemented")

    def _body(self, given: object) -> tuple[bytes, bool]:
        """The body and whether it is plain text. A byte array in a Variant goes; in its own Byte variable only
        WinHttpRequest takes it, and the rest raise the parameter's error, as ``h.Send b`` does in Excel."""
        from pyopenvba.interpreter._runtime import Slot

        direct = False
        if isinstance(given, Slot):
            direct = bool(given.array_of) or given.declared != "Variant"
            given = given.get()
        if given is MISSING or given is EMPTY:
            return b"", False
        if isinstance(given, VBAArray):
            if given.element_type != "Byte":
                raise VBAUnsupportedError("HTTP Send of an array that is not bytes is not implemented")
            if direct and self.stack is not _WINHTTP:
                raise error(*_PARAMETER)
            return bytes(int(to_integer(one, "Byte")) for one in given.items), False
        if given is NULL or isinstance(given, VBAObject):
            raise VBAUnsupportedError("HTTP Send of Null or of an object is not implemented")
        return to_text(given).encode("utf-8", errors="surrogatepass"), True

    def _pump(self) -> None:
        """An asynchronous request's end: nothing is raised, and the reads report any failure."""
        self._drop_pending()
        self._finish(*self.pending)

    @method
    def Abort(self) -> object:
        self._drop_pending()
        self.stage = "aborted" if self.stage not in ("new", "aborted") and self.response is not None else "new"
        self.response, self.failed_headers = None, []
        return EMPTY

    @method
    def WaitForResponse(self, Timeout: object = MISSING) -> object:
        del Timeout
        if self.stack.inet:
            raise error(438)
        if self.stage in ("new", "aborted", "opened"):
            raise error(*(self.stack.opened[0] if self.stage == "opened" else self.stack.new[0]))
        if self.stage == "waiting":
            self._pump()
        return True

    @member
    def Option(self, OptionIndex: object = MISSING) -> object:
        index = int(to_integer(OptionIndex, "Long"))
        if self.stack is not _WINHTTP:
            raise error(438)
        if index == 0:
            return self.agent
        if index == 1 and self.stage == "done":
            return self.final_url
        if index == 2:
            return VBAInt(65001, "Long")
        if index == 6:
            return self.follows
        raise VBAUnsupportedError(f"WinHttpRequest.Option({index}) here is not implemented")

    @setter("Option")
    def _option(self, OptionIndex: object, value: object) -> None:
        index = int(to_integer(OptionIndex, "Long"))
        if self.stack is not _WINHTTP:
            raise error(438)
        if index == 0:
            self.agent = to_text(value)
        elif index == 6:
            self.follows = to_bool(value)
        else:
            raise VBAUnsupportedError(f"Setting WinHttpRequest.Option({index}) is not implemented")

    # -- the exchange

    def _sessions(self) -> Sessions:
        return self.interpreter.internet_sessions

    def _jar(self) -> dict[str, dict[str, tuple[str, str]]]:
        return self._sessions().cookies if self.stack.inet else self.cookies

    def _finish(self, body: bytes, is_text: bool) -> _Failure | None:
        """Run the request to its end; the failure Send raises, with the object left as Excel leaves it."""
        method_, url = self.method, self.url
        tried_login = False
        for _ in range(_HOPS + 1):
            try:
                answer = self._exchange(method_, url, body, is_text)
            except socket.gaierror:
                return self._failed(_INET_RESOURCE, 12007, _UNRESOLVED)
            except TimeoutError:
                if self.stack.inet:
                    raise VBAUnsupportedError("An XMLHTTP request that times out is not implemented") from None
                return self._failed(_TIMED_OUT, 0, _TIMED_OUT)
            except ConnectionError:
                return self._failed(_INET_RESOURCE, 12029, _REFUSED)
            except OSError as problem:
                raise VBAUnsupportedError(f"An HTTP request that fails with {problem!r} is not implemented") from None
            except http.client.HTTPException as problem:
                raise VBAUnsupportedError(f"An HTTP answer that reads as {problem!r} is not implemented") from None
            self.final_url = url
            if answer.status == 401 and self.user is not None and not tried_login and self._challenged(answer):
                tried_login = True
                self.login = self.user
                continue
            if answer.status not in _REDIRECTS or not self.follows:
                if tried_login and answer.status != 401 and self.stack.inet and self.login is not None:
                    self._sessions().logins[self._host(url)] = self._basic(self.login)
                return self._finished(answer)
            location = next((value for name, value in answer.headers if name.lower() == "location"), None)
            if location is None:
                self.failed_headers = [one for one in answer.headers if one[0].lower() != "content-type"]
                return self._failed(_INET_DOWNLOAD, 12150, _REDIRECT)
            target = urljoin(url, location)
            if urlsplit(target).scheme.lower() != urlsplit(url).scheme.lower():
                raise VBAUnsupportedError("An HTTP redirect to another scheme is not implemented")
            if answer.status in (301, 302, 303) and method_ not in ("GET", "HEAD"):
                if method_ != "POST":
                    if self.stack.inet:
                        raise VBAUnsupportedError(f"{method_} through a redirect hangs XMLHTTP in Excel")
                else:
                    method_, body, is_text = "GET", b"", False
            url = target
        return self._failed(_INET_REDIRECT, 0, _REDIRECT)

    def _challenged(self, answer: _Response) -> bool:
        return any(name.lower() == "www-authenticate" and value.lower().startswith("basic")
                   for name, value in answer.headers)

    def _basic(self, login: tuple[str, str]) -> str:
        return "Basic " + base64.b64encode(f"{login[0]}:{login[1]}".encode()).decode("ascii")

    def _host(self, url: str) -> str:
        return (urlsplit(url).hostname or "").lower()

    def _failed(self, inet: _Failure, status: int, http_failure: _Failure) -> _Failure:
        if self.stack.inet:
            self.response, self.stage, self.failed_headers = _Response(status, "", [], b""), "done", []
            return inet
        self.stage = "failed"
        return http_failure

    def _finished(self, answer: _Response) -> _Failure | None:
        self.stage = "done"
        if self.stack.inet and answer.status == 204:
            self.response = _Response(1223, "", [], b"")
            return _INET_ABORTED
        self.response = answer
        return None

    def _exchange(self, verb: str, url: str, body: bytes, is_text: bool) -> _Response:
        self.interpreter.outside("An HTTP request")
        parts = urlsplit(url)
        host = parts.hostname or ""
        secure = parts.scheme.lower() == "https"
        port = parts.port or (443 if secure else 80)
        address = host if parts.port is None else f"{host}:{parts.port}"
        path = self._target(url)
        lines = self._inet_lines(verb, address, path, body) if self.stack.inet else self._http_lines(
            verb, address, path, body, is_text)
        send = b"" if verb == "GET" and self.stack is not _WINHTTP else body
        request = "\r\n".join(lines).encode("latin-1") + b"\r\n\r\n" + send
        connection = socket.create_connection((host, port), timeout=self.connect_wait)
        try:
            connection.settimeout(self.receive_wait)
            stream = ssl.create_default_context().wrap_socket(connection, server_hostname=host) if secure else connection
            stream.sendall(request)
            reply = http.client.HTTPResponse(stream, method=verb)
            reply.begin()
            data = reply.read()
            headers = reply.getheaders()
            answer = _Response(reply.status, reply.reason, headers, data)
        finally:
            connection.close()
        self._keep_cookies(host, headers)
        if self.stack.inet:
            self._unpack(answer)
        return answer

    def _unpack(self, answer: _Response) -> None:
        """WinINet asks for gzip and deflate, and hands back what they unpack to with no word of either."""
        kind = next((value.lower() for name, value in answer.headers if name.lower() == "content-encoding"), "")
        if kind not in ("gzip", "deflate"):
            return
        answer.body = gzip.decompress(answer.body) if kind == "gzip" else zlib.decompress(answer.body)
        answer.headers = [one for one in answer.headers if one[0].lower() not in ("content-encoding", "content-length")]

    def _target(self, url: str) -> bytes | str:
        """What the request line asks for: spaces escaped; a # goes as %23 in WinINet and as itself in WinHTTP; beyond
        ASCII goes as Latin-1 in WinINet and as UTF-8 in WinHTTP. A later % escape is left as it is."""
        after = url.partition("://")[2]
        slash = after.find("/")
        query = after.find("?")
        start = min(place for place in (slash, query, len(after)) if place >= 0)
        path = after[start:]
        if path.startswith("?"):
            path = "/" + path
        path = path or "/"
        if self.stack.inet:
            question = path.find("?")
            if question >= 0:
                path = path[:question] + path[question:].replace("#", "%23")
            else:
                path = path.partition("#")[0]
            path = path.replace(" ", "%20")
            return "".join(char if ord(char) < 256 else "?" for char in path)
        path = path.replace(" ", "%20")
        return path.encode("utf-8", errors="surrogatepass").decode("latin-1")

    def _cookie_line(self, host: str, url_path: str) -> str:
        wanted = url_path.partition("?")[0].partition("#")[0] or "/"
        kept = self._jar().get(host, {})
        return "; ".join(f"{name}={value}" for name, (value, path) in kept.items() if _keeps_cookie(path, wanted))

    def _keep_cookies(self, host: str, headers: list[tuple[str, str]]) -> None:
        for name, value in headers:
            if name.lower() != "set-cookie":
                continue
            first, *rest = value.split(";")
            key, _, kept = first.partition("=")
            attributes = {part.partition("=")[0].strip().lower(): part.partition("=")[2].strip() for part in rest}
            jar = self._jar().setdefault(host, {})
            if attributes.get("max-age") == "0":
                jar.pop(key.strip(), None)
            else:
                jar[key.strip()] = (kept.strip(), attributes.get("path") or "/")

    def _inet_lines(self, verb: str, address: str, path: object, body: bytes) -> list[str]:
        """WinINet's order: Accept, what was set, its own, and the connection's. A second value joins the first."""
        joined: dict[str, tuple[str, str]] = {}
        for name, value in self.headers:
            if not value:
                continue
            key = name.lower()
            joined[key] = (joined[key][0], f"{joined[key][1]}, {value}") if key in joined else (name, value)
        lines = [f"{verb} {path} HTTP/1.1", f"Accept: {joined.get('accept', ('', '*/*'))[1]}"]
        ignored = {"accept", "accept-language", "accept-encoding", "connection", "cookie", "host", "content-length"}
        lines += [f"{name}: {value}" for key, (name, value) in joined.items() if key not in ignored]
        lines += [f"Accept-Language: {joined.get('accept-language', ('', _LANGUAGE))[1]}", "UA-CPU: AMD64",
                  "Accept-Encoding: gzip, deflate"]
        if "user-agent" not in joined:
            lines.append(f"User-Agent: {_INET_AGENT}")
        lines.append(f"Host: {address}")
        if verb != "GET":
            lines.append(f"Content-Length: {len(body)}")
        connection = joined.get("connection", ("", "Keep-Alive"))[1]
        lines.append(f"Connection: {connection.capitalize() if connection.lower() == 'close' else connection}")
        if verb != "GET":
            lines.append("Cache-Control: no-cache")
        login = self._basic(self.login) if self.login else self._sessions().logins.get(self._host(self.url))
        if login:
            lines.append(f"Authorization: {login}")
        cookies = self._cookie_line(self._host(self.url), str(path))
        if cookies:
            lines.append(f"Cookie: {cookies}")
        return lines

    def _http_lines(self, verb: str, address: str, path: object, body: bytes, is_text: bool) -> list[str]:
        """WinHTTP's order, whichever order they were set in. WinHttpRequest adds a Content-Type to plain ASCII text,
        and a charset to one that was set without, and sends a GET's body, which ServerXMLHTTP drops."""
        winhttp = self.stack is _WINHTTP
        user: dict[str, str] = {}
        rest: list[tuple[str, str]] = []
        for name, value in self.headers:
            key = name.lower()
            if key in ("accept", "accept-language", "accept-encoding", "connection", "cookie", "host", "content-type",
                       "user-agent", "content-length"):
                user[key] = value
            else:
                rest.append((name, value))
        sent = body if verb != "GET" or winhttp else b""
        ascii_text = winhttp and is_text and sent.isascii() and (verb != "GET" or bool(sent))
        lines = [f"{verb} {path} HTTP/1.1", f"Connection: {user.get('connection', 'Keep-Alive')}"]
        kind = user.get("content-type")
        if ascii_text:
            if kind is None:
                kind = "text/plain; Charset=UTF-8"
            elif "charset" not in kind.lower():
                kind += "; Charset=UTF-8"
        if kind is not None:
            lines.append(f"Content-Type: {kind}")
        lines.append(f"Accept: {user.get('accept', '*/*')}")
        if "accept-encoding" in user:
            lines.append(f"Accept-Encoding: {user['accept-encoding']}")
        if "accept-language" in user:
            lines.append(f"Accept-Language: {user['accept-language']}")
        elif not winhttp:
            lines.append(f"Accept-Language: {_LANGUAGE}")
        if "cookie" in user:
            lines.append(f"Cookie: {user['cookie']}")
        if winhttp and "host" in user:
            lines.append(f"Host: {user['host']}")
        lines.append(f"User-Agent: {user.get('user-agent', self.agent)}")
        lines += [f"{name}: {value}" for name, value in rest]
        if verb != "GET" or (winhttp and is_text and sent):
            lines.append(f"Content-Length: {len(sent)}")
        if not (winhttp and "host" in user):
            lines.append(f"Host: {address}")
        if self.login:
            lines.append(f"Authorization: {self._basic(self.login)}")
        cookies = self._cookie_line(self._host(self.url), str(path))
        if cookies:
            lines.append(f"Cookie: {cookies}")
        return lines
