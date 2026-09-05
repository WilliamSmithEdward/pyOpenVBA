"""Compact a database with DAO while the engine's clock stands still.

    python frozen_clock.py <source.accdb> <destination.accdb>

The engine keys its encoding of owner and permission SIDs to the file's
creation date, and ``CompactDatabase`` stamps the destination with the
time it runs.  Holding that time at the source's own creation date makes
the destination keep the source's SIDs byte for byte, which is what lets
pyOpenVBA's compaction -- which keeps the creation date -- be compared
with the engine's output page by page.  Every other stamp the compaction
writes is that one instant too.

The engine is hosted in this process through pywin32, and the clock
imports of its DLLs and of the C runtimes they use are pointed at the
frozen instant by rewriting their import address tables.  Nothing outside
this process is touched.  Prints the frozen serial (days since 1899-12-30)
on success.  Test tooling only: the library itself never uses COM.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import datetime as dt
import os
import struct
import sys
from pathlib import Path
from typing import Any

if __name__ == "__main__" and str(Path(__file__).resolve().parents[2] / "src") not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from pyopenvba.access._pages import OFFSET_CREATION_DATE, toggle_definition_mask  # noqa: E402

#: ``ctypes`` seen as untyped: its Windows-only names (``WinDLL``,
#: ``WINFUNCTYPE``, ``get_last_error``) are not in the stubs off Windows,
#: and this tool is type-checked everywhere the tests are.
_windows: Any = ctypes
PAGE_READWRITE = 0x04
EPOCH = dt.datetime(1601, 1, 1)
ZERO_DAY = dt.datetime(1899, 12, 30)
PE32_PLUS = 0x20B
#: The DLLs whose clock imports are redirected: the engine's, Office's
#: shared code, and the C runtimes any of them call through.
PATCHED_DLLS = ("\\ACE", "\\MSO", "VCRUNTIME", "MSVCP", "MSVCR", "UCRTBASE")


class FILETIME(ctypes.Structure):
    _fields_ = [("dwLowDateTime", wt.DWORD), ("dwHighDateTime", wt.DWORD)]


class SYSTEMTIME(ctypes.Structure):
    _fields_ = [
        ("wYear", wt.WORD),
        ("wMonth", wt.WORD),
        ("wDayOfWeek", wt.WORD),
        ("wDay", wt.WORD),
        ("wHour", wt.WORD),
        ("wMinute", wt.WORD),
        ("wSecond", wt.WORD),
        ("wMilliseconds", wt.WORD),
    ]


_HOOKS: list[object] = []


def _read32(address: int) -> int:
    return ctypes.c_uint32.from_address(address).value


def _read64(address: int) -> int:
    return ctypes.c_uint64.from_address(address).value


def freeze(instant: dt.datetime) -> dict[str, list[str]]:
    """Point the loaded engine and runtime DLLs' clock imports at
    ``instant`` (a local time).  Returns what was patched, per DLL."""
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32process  # pyright: ignore[reportMissingModuleSource]

    k32 = _windows.WinDLL("kernel32", use_last_error=True)
    k32.VirtualProtect.argtypes = [ctypes.c_void_p, ctypes.c_size_t, wt.DWORD, ctypes.POINTER(wt.DWORD)]
    ticks = int(round((instant - EPOCH).total_seconds() * 10_000_000))

    def as_filetime(address: int | None) -> None:
        if address:
            filetime = FILETIME.from_address(address)
            filetime.dwLowDateTime = ticks & 0xFFFFFFFF
            filetime.dwHighDateTime = ticks >> 32

    def as_systemtime(address: int | None) -> None:
        if address:
            systemtime = SYSTEMTIME.from_address(address)
            systemtime.wYear, systemtime.wMonth, systemtime.wDay = instant.year, instant.month, instant.day
            systemtime.wDayOfWeek = (instant.weekday() + 1) % 7
            systemtime.wHour, systemtime.wMinute, systemtime.wSecond = instant.hour, instant.minute, instant.second
            systemtime.wMilliseconds = instant.microsecond // 1000

    prototype = _windows.WINFUNCTYPE(None, ctypes.c_void_p)
    hooks = {
        b"GetSystemTimeAsFileTime": prototype(as_filetime),
        b"GetSystemTimePreciseAsFileTime": prototype(as_filetime),
        b"GetLocalTime": prototype(as_systemtime),
        b"GetSystemTime": prototype(as_systemtime),
    }
    _HOOKS.extend(hooks.values())
    addresses = {name: ctypes.cast(hook, ctypes.c_void_p).value or 0 for name, hook in hooks.items()}

    def patch(base: int) -> list[str]:
        nt_headers = base + _read32(base + 0x3C)
        optional = nt_headers + 24
        if ctypes.c_uint16.from_address(optional).value != PE32_PLUS:
            return []
        import_rva = _read32(optional + 112 + 8)
        if not import_rva:
            return []
        done: list[str] = []
        descriptor = base + import_rva
        while True:
            original_thunks, name_rva, first_thunk = _read32(descriptor), _read32(descriptor + 12), _read32(descriptor + 16)
            if not name_rva:
                break
            dll = ctypes.string_at(base + name_rva).decode(errors="replace")
            index = 0
            while original_thunks:
                thunk = _read64(base + original_thunks + 8 * index)
                if thunk == 0:
                    break
                if not thunk >> 63:
                    function = ctypes.string_at(base + (thunk & 0xFFFFFFFF) + 2)
                    if function in addresses:
                        slot = base + first_thunk + 8 * index
                        old = wt.DWORD()
                        if not k32.VirtualProtect(ctypes.c_void_p(slot), 8, PAGE_READWRITE, ctypes.byref(old)):
                            raise OSError(_windows.get_last_error())
                        ctypes.c_uint64.from_address(slot).value = addresses[function]
                        k32.VirtualProtect(ctypes.c_void_p(slot), 8, old.value, ctypes.byref(old))
                        done.append(f"{dll}!{function.decode()}")
                index += 1
            descriptor += 20
        return done

    # pywin32 ships no stubs for these, so they are looked up by name.
    enumerate_modules = getattr(win32process, "EnumProcessModules")
    module_file_name = getattr(win32process, "GetModuleFileNameEx")
    process = win32api.GetCurrentProcess()
    patched: dict[str, list[str]] = {}
    for module in enumerate_modules(process):
        path = str(module_file_name(process, module))
        if any(marker in path.upper() for marker in PATCHED_DLLS):
            found = patch(int(module))
            if found:
                patched[Path(path).name] = found
    return patched


def creation_instant(path: Path) -> tuple[float, dt.datetime]:
    """The creation date on page 0, as its serial and as the local datetime
    the engine turns back into that serial."""
    plain = toggle_definition_mask(path.read_bytes()[:4096])
    serial = struct.unpack_from("<d", plain, OFFSET_CREATION_DATE)[0]
    days = int(serial)
    instant = ZERO_DAY + dt.timedelta(days=days) + dt.timedelta(microseconds=round((serial - days) * 86400e6))
    return serial, instant


def main() -> None:
    import win32com.client  # pyright: ignore[reportMissingModuleSource]

    source, destination = Path(sys.argv[1]), Path(sys.argv[2])
    serial, instant = creation_instant(source)
    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    # Opening once loads the engine's DLLs, so their imports can be patched.
    engine.OpenDatabase(str(source)).Close()
    patched = freeze(instant)
    if not any(name.upper().startswith("ACECORE") for name in patched):
        print(f"the engine's clock was not hooked: {patched}", file=sys.stderr)
        sys.stderr.flush()
        os._exit(2)
    if destination.exists():
        destination.unlink()
    engine.CompactDatabase(str(source), str(destination))
    print(repr(serial))
    sys.stdout.flush()
    os._exit(0)


if __name__ == "__main__":
    main()
