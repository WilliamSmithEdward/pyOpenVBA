"""Hooks: what the engine cannot do itself, done by a function the caller registers (ExcelApplication.set_hook).

The model answers what it can measure, a macro's argument errors and its events, and checks any path the macro gave. Then it calls the function with a HookRequest, which
holds values only: nothing in it is acted on by the model, and it opens no viewer and prints nothing. What the function makes is its own, and is not claimed to match Excel.

A macro chooses what is in a request, so a function must not trust it blindly. `path` is the file as the model resolved it, and has passed the gate. `arguments` hold what the macro
wrote: a printer's name and a number of copies are the host's to honour or not.

The gate, for a file named while a macro runs under the sandbox: only inside a folder named with set_hook, with links, junctions and .. resolved, and UNC and device paths refused
unless the folder named is one. With no folder named, no path passes. With sandbox=False there is no limit, as a macro can then write anywhere anyway. A hook with no file to write
runs under the sandbox: registering it is the caller's grant.

What a function raises: a VBAError goes on as it is, so a function can give any error number, and one that stops a run still stops it. Any other Exception is a trappable 1004 with
the original as its __cause__. That is chosen, on the 1004 Excel gives when it cannot write a file (tests/fixtures/output_members.json). A function that asks for its own
member again in the same run is refused rather than recursing.
"""
from __future__ import annotations

import os
import tempfile
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import Path
from types import MappingProxyType
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.interpreter._files import native_path, spelled
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Workbook
    from pyopenvba.interpreter._runtime import Interpreter

#: The members with a hook. Each is added here when it is built, so that a name mistyped raises rather than never being called.
HOOK_NAMES = frozenset({'ExportAsFixedFormat', 'PrintOut', 'Refresh'})
#: Names Windows gives to devices rather than files, whatever follows the dot.
_DEVICES = frozenset({'CON', 'PRN', 'AUX', 'NUL', 'CONIN$', 'CONOUT$', *(f'COM{at}' for at in range(1, 10)), *(f'LPT{at}' for at in range(1, 10))})


@dataclass(frozen=True)
class ExternalSource:
    """What a Refresh asks for: the kind of source, OLEDB or ODBC, its connection string as the macro set it with that kind's prefix
    taken off, the command, the command's kind as Excel names it (xlCmdSql, xlCmdTable, xlCmdDefault), the parameters given, and
    whether the field names are wanted. The model lays out the rows; the function only fetches them.

    The connection string goes to the function whole, as it needs it to connect. Wherever the model shows one, in this object's repr and
    in every error, a password is written ****."""

    kind: str
    connection: str
    command_text: str
    command_type: str = 'xlCmdSql'
    parameters: tuple[tuple[str, object], ...] = ()
    fields_wanted: bool = True

    def __repr__(self) -> str:
        from pyopenvba.apps.excel._query_databases import masked

        return (f'ExternalSource(kind={self.kind!r}, connection={masked(self.connection)!r}, command_text={self.command_text!r}, '
                f'command_type={self.command_type!r}, parameters={self.parameters!r}, fields_wanted={self.fields_wanted!r})')


@dataclass(frozen=True)
class HookRequest:
    """One call of a member with a hook.

    ``member`` is its name, and ``owner`` the class it was called on: Worksheet, Workbook, Range, Chart or Sheets. ``target`` and ``workbook`` are the live objects. ``arguments``
    holds only what the macro gave, under Excel's names, as read-only values. ``path`` is the file to write, past the gate, with Excel's name rules applied, or None if there is none.
    """

    member: str
    owner: str
    target: object
    workbook: Workbook
    arguments: Mapping[str, object]
    path: Path | None = None
    #: For Refresh: the source to fetch rows from, as the macro named it.
    source: ExternalSource | None = None

    def saved_bytes(self) -> bytes:
        """What SaveCopyAs would write now, for a tool that converts files. Nothing about the workbook changes: not Saved, not its name or its path, and no event runs. A workbook never
        saved gives a new file in its own format. One that links to other files reports itself, as SaveCopyAs elsewhere does."""
        from pyopenvba.apps.excel._io import save_workbook, workbook_format
        from pyopenvba.apps.excel._save_copy import refuse_moved_links

        refuse_moved_links(self.workbook)
        suffix = '.xlsm' if workbook_format(self.workbook) == 52 else '.xlsx'
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            target = Path(folder) / f'copy{suffix}'
            save_workbook(self.workbook, target, fresh_suffix=suffix)
            return target.read_bytes()


@dataclass(frozen=True)
class Hook:
    function: Callable[[HookRequest], object]
    folders: tuple[str, ...] = ()


def set_hook(hooks: dict[str, object], name: str, function: object, folders: object = ()) -> None:
    if name not in HOOK_NAMES:
        raise ValueError(f'{name!r} has no hook; those that do are {sorted(HOOK_NAMES)}')
    if function is None:
        hooks.pop(name, None)
        return
    if not callable(function):
        raise TypeError('a hook is a function taking a HookRequest, or None to remove it')
    if isinstance(folders, (str, os.PathLike)):
        raise TypeError('folders is a list of folders')
    hooks[name] = Hook(function, tuple(os.fspath(one) for one in folders))  # pyright: ignore[reportArgumentType, reportUnknownVariableType, reportUnknownArgumentType, reportGeneralTypeIssues]


def hook_of(application: Application, member: str) -> Hook | None:
    found = application.interpreter.hooks.get(member) if application.interpreter is not None else None
    return found if isinstance(found, Hook) else None


def final_path(path: str) -> str:
    """The file a macro's path names, as this system opens it: a backslash a separator off Windows too, each part in the case the disk
    spells it, then links and .. resolved. The gate checks this and the read or write uses this, so nothing turns the path round between
    them: off Windows a backslash is a character of a name, and "grant/..\\away" checked as given is one file inside grant that a read
    turning its backslashes round would find outside."""
    return os.path.realpath(spelled(native_path(path)))


def _real(path: str) -> Path:
    return Path(final_path(path))


def is_inside(path: Path | str, folders: tuple[str, ...]) -> bool:
    """Whether a file lies in one of the folders once links and .. are resolved. Parts compare whole, so C:\\out2 is not in C:\\out; a UNC or device path has an anchor of its own, so it
    is in nothing but a folder of that kind; and the names by which Windows reaches a device or a file's other streams are in nothing."""
    real = _real(os.fspath(path))
    if os.name == 'nt' and (':' in real.name or real.name.split('.')[0].rstrip(' ').upper() in _DEVICES):
        return False
    return any(real != base and real.is_relative_to(base) for base in (_real(one) for one in folders))


def gate(application: Application, hook: Hook, member: str, path: Path) -> None:
    """Refuse a macro's file before the function hears of it."""
    interpreter = application.interpreter
    if interpreter is None or not (interpreter.sandbox and interpreter.frames):
        return
    if not hook.folders:
        application.outside(f'{member} writing a file')
    _refuse_outside(application, hook.folders, path, f'{member} may not write there: it is outside the folders named with set_hook')


def _refuse_outside(application: Application, folders: tuple[str, ...], path: Path | str, message: str) -> None:
    """Refuse a path that resolves outside ``folders``, saying nothing of the path: the macro learns nothing of what lies outside."""
    if not is_inside(path, folders):
        from pyopenvba.exceptions import VBASandboxError

        interpreter = application.interpreter
        raise VBASandboxError(message, where=interpreter.current_where() if interpreter is not None else "")


def set_link_folders(interpreter: Interpreter, folders: object) -> None:
    """The folders a sandboxed macro may read linked workbooks' files in (ExcelApplication.allow_link_reads); () for none."""
    if isinstance(folders, (str, os.PathLike)):
        raise TypeError('folders is a list of folders')
    interpreter.link_folders = tuple(os.fspath(one) for one in folders)  # pyright: ignore[reportArgumentType, reportUnknownVariableType, reportUnknownArgumentType, reportGeneralTypeIssues]


class LinkReadPassed:
    """What link_gate hands back for a file it checked, the one thing load_workbook takes in place of its own gate: no other
    code makes one, so nothing skips the gate on opening a workbook by accident."""

    __slots__ = ()


_PASSED = LinkReadPassed()


def passed(token: object) -> bool:
    return token is _PASSED


def link_gate(application: Application, what: str, path: str | None = None) -> tuple[LinkReadPassed, str]:
    """The sandbox's gate on reading a linked workbook's file. With no folders granted it is the interpreter's own gate, shut, before
    anything on disk is looked at; with folders, a member passes, and each file it reads, ``path``, must lie in one. It gives back the
    token that lets load_workbook read the file, and the file's final_path, the one string checked: the read opens that and nothing else."""
    interpreter = application.interpreter
    if interpreter is not None and interpreter.sandbox and interpreter.frames and not interpreter.link_folders:
        application.outside(what)
    final = '' if path is None else final_path(path)
    if interpreter is not None and interpreter.sandbox and interpreter.frames and path is not None:
        _refuse_outside(application, interpreter.link_folders, final,
                        f'{what} may not read there: it is outside the folders named with allow_link_reads')
    return _PASSED, final


def call(application: Application, hook: Hook, member: str, owner: str, target: object, workbook: Workbook, arguments: dict[str, object], path: Path | None = None,
         source: ExternalSource | None = None) -> object:
    """Call the function, as the docstring above says, and give back what it returns: nothing for an output hook, the rows for Refresh."""
    interpreter = application.interpreter
    assert interpreter is not None
    running: list[str] = interpreter.hooks_running
    if member in running:
        raise VBAUnsupportedError(f'{member} called from its own hook is not implemented')
    running.append(member)
    try:
        return hook.function(HookRequest(member, owner, target, workbook, MappingProxyType(dict(arguments)), path, source))
    except VBAError:
        raise
    except Exception as failed:
        # A source's function may echo its connection string, password and all, so its words stay in __cause__ alone.
        said = type(failed).__name__ if source is not None else failed
        raise error(1004, f'The function set for {member} failed: {said}') from failed
    finally:
        running.remove(member)
