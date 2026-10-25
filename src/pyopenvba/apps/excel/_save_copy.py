"""Workbook.SaveCopyAs: the file a save would write, written under another name, the workbook left as it is.

Measured in live Excel (scripts/measure_save_copy_as.py, tests/fixtures/save_copy_as.json):

- The copy holds what Save would write at that moment, a workbook under
  manual calculation calculated first, as a save calculates it. It is in
  the workbook's own format whatever the name's extension: a .xlsm name
  gets a macro-free workbook's file and a .xlsx name a macro-enabled one's.
  A name without an extension is written without one.
- No save event is raised. The workbook keeps its name, its path and its
  Saved flag, and its Last save time and Last author move as a save's do.
- A file already there is replaced without a prompt, alerts on or off.
- An empty name writes nothing and raises nothing. The workbook is still
  calculated, and its Last save time and Last author still move.
- Null is error 1004, and so is a name in a folder that does not exist,
  the workbook's own file or another open workbook's file.
- Copied into another folder, a workbook's hyperlink to a file is written
  relative to the copy's folder, and the workbook's own link moves there
  too, which every later save then writes wrong. That copy is refused
  here, and so is one into another folder of a workbook whose formulas
  link to other workbooks, which is not measured.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, NULL, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook

#: Excel's message for a copy it cannot write, the file's path in the quotes.
CANNOT_ACCESS = ("Microsoft Excel cannot access the file '{}'. There are several possible reasons:\n\n"
                 "• The file name or path does not exist.\n"
                 "• The file is being used by another program.\n"
                 "• The workbook you are trying to save has the same name as a currently open workbook.")


def save_copy(book: Workbook, filename: object) -> None:
    """Write ``book`` to ``filename`` as SaveCopyAs writes it (see the module's notes)."""
    from pyopenvba.apps.excel._document_properties import stamp
    from pyopenvba.apps.excel._io import calculate_for_save, save_workbook, workbook_format

    if filename is MISSING:
        raise VBAUnsupportedError("SaveCopyAs without a filename is not implemented; supply Filename")
    if filename is NULL:
        raise error(1004, "Method 'SaveCopyAs' of object '_Workbook' failed")
    if not isinstance(filename, str):
        raise VBAUnsupportedError("SaveCopyAs with a filename that is not text is not implemented")
    if not filename:
        calculate_for_save(book)
        stamp(book)
        return
    target = Path(os.path.abspath(filename))
    if not target.parent.is_dir() or _open_file(book, target):
        raise error(1004, CANNOT_ACCESS.format(target))
    if str(target.resolve().parent) != book.links_folder and _links_to_files(book):
        raise VBAUnsupportedError("SaveCopyAs into another folder of a workbook that links to files is not "
                                  "implemented: Excel moves the workbook's own links to the copy's folder")
    try:
        save_workbook(book, target, fresh_suffix=".xlsm" if workbook_format(book) == 52 else ".xlsx")
    except OSError as failure:
        raise error(1004, CANNOT_ACCESS.format(target)) from failure


def _open_file(book: Workbook, target: Path) -> bool:
    """Whether ``target`` is the file of a workbook open beside ``book``, or of ``book`` itself."""
    wanted = os.path.normcase(target.resolve())
    return any(other.path and os.path.normcase(Path(other.path, other.name).resolve()) == wanted
               for other in book.application.workbooks_.books)


def _links_to_files(book: Workbook) -> bool:
    """Whether the workbook links to a file by where it is: a hyperlink to one, or a formula to another workbook."""
    from pyopenvba.apps.excel._hyperlinks_file import names_file

    return bool(book.linked_books) or any(names_file(link.address) for sheet in book.sheets_
                                          for link in sheet.hyperlinks)
