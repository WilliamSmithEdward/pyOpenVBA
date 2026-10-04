"""Worksheet document identities and code transfer between projects."""
from __future__ import annotations

import re

from pyopenvba.apps.excel._model import Workbook, Worksheet
from pyopenvba.vba import split_attribute_header


def materialize(book: Workbook) -> None:
    """Allocate pending code names in tab order, as VBProject access does."""
    if not book.code_name:
        book.code_name = "ThisWorkbook"
    reserved = book.used_code_names
    reserved.add(book.code_name.casefold())
    reserved.update(sheet.code_name.casefold() for sheet in book.sheets_ if sheet.code_name)
    if book.project_runtime is not None:
        reserved.update(book.project_runtime.modules)
    for sheet in book.sheets_:
        if not sheet.code_name:
            number = 1
            while f"sheet{number}" in reserved:
                number += 1
            sheet.code_name = f"Sheet{number}"
            reserved.add(sheet.code_name.casefold())


def remove_document(sheet: Worksheet) -> None:
    book = sheet.book
    if sheet.code_name:
        book.deleted_document_names.add(sheet.code_name)
        book.used_code_names.add(sheet.code_name.casefold())
        if book.project_runtime is not None:
            book.project_runtime.modules.pop(sheet.code_name.casefold(), None)
        book.pending_document_sources.pop(sheet.code_name, None)
    sheet.vba_document = None


def copy_document(source: Worksheet, copied: Worksheet) -> None:
    from pyopenvba.apps.excel._projects import attach_project

    project = attach_project(copied.book)
    if not source.code_name:
        return
    reserved = copied.book.used_code_names
    reserved.update(sheet.code_name.casefold() for sheet in copied.book.sheets_ if sheet is not copied and sheet.code_name)
    if project is not None:
        reserved.update(project.modules)
    name = source.code_name
    match = re.fullmatch(r"(.*?)(\d+)", name)
    prefix, number = (match.group(1), int(match.group(2)) + 1) if match else (name, 1)
    while name.casefold() in reserved:
        name = f"{prefix}{number}"
        number += 1
    copied.code_name = name
    reserved.add(name.casefold())
    original = source.vba_document
    if original is not None and project is not None:
        # The new document gets fresh fields/statics; only source transfers.
        body = split_attribute_header(original.module.parsed.source)[1]
        from pyopenvba.interpreter._runtime import Interpreter

        Interpreter.add_module(project, body, name=name, kind="document")
        project.bind_document(name, copied)
    elif source.code_name in source.book.pending_document_sources:
        copied.book.pending_document_sources[name] = source.book.pending_document_sources[source.code_name]
    elif source.book.package is not None and source.book.package.has("xl/vbaProject.bin"):
        from pyopenvba.apps.excel._vba_save import MemoryExcelFile

        # with_vba=False preserves code without parsing or executing it.
        with MemoryExcelFile(source.book.package.serialize()) as host:
            body = split_attribute_header(host.get_module(source.code_name))[1]
        copied.book.pending_document_sources[name] = body
