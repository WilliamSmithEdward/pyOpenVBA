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
from pyopenvba.interpreter._files import spelled
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Workbook
    from pyopenvba.interpreter._runtime import Interpreter

#: The members with a hook. Each is added here when it is built, so that a name mistyped raises rather than never being called.
HOOK_NAMES = frozenset({'ExportAsFixedFormat', 'PrintOut'})
#: Names Windows gives to devices rather than files, whatever follows the dot.
_DEVICES = frozenset({'CON', 'PRN', 'AUX', 'NUL', 'CONIN$', 'CONOUT$', *(f'COM{at}' for at in range(1, 10)), *(f'LPT{at}' for at in range(1, 10))})


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


def _real(path: str) -> Path:
    return Path(os.path.realpath(spelled(path)))


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


def link_gate(application: Application, what: str, path: str | None = None) -> None:
    """The sandbox's gate on reading a linked workbook's file, before anything on disk is looked at. With no folders granted it
    is the interpreter's own gate, shut; with folders, a member passes, and each file it reads, ``path``, must lie in one."""
    interpreter = application.interpreter
    if interpreter is None or not (interpreter.sandbox and interpreter.frames):
        return
    if not interpreter.link_folders:
        application.outside(what)
    if path is not None:
        _refuse_outside(application, interpreter.link_folders, path,
                        f'{what} may not read there: it is outside the folders named with allow_link_reads')


def call(application: Application, hook: Hook, member: str, owner: str, target: object, workbook: Workbook, arguments: dict[str, object], path: Path | None = None) -> None:
    """Call the function, as the docstring above says."""
    interpreter = application.interpreter
    assert interpreter is not None
    running: list[str] = interpreter.hooks_running
    if member in running:
        raise VBAUnsupportedError(f'{member} called from its own hook is not implemented')
    running.append(member)
    try:
        hook.function(HookRequest(member, owner, target, workbook, MappingProxyType(dict(arguments)), path))
    except VBAError:
        raise
    except Exception as failed:
        raise error(1004, f'The function set for {member} failed: {failed}') from failed
    finally:
        running.remove(member)
