"""Power Query (Get and Transform) inside Excel workbooks, in pure Python.

A workbook keeps its queries in a custom XML part as a base64 blob: an
OPC package holding an M section document, a metadata document beside it,
and a permission list.  This package reads all of that, changes it, and
writes it back, with no Excel and no dependencies.

    from pyopenvba import PowerQueryWorkbook

    with PowerQueryWorkbook("orders.xlsx") as book:
        for query in book.queries():
            print(query.name, query.load_target)
        book.query("Orders").formula = "let Source = #table({}, {}) in Source"
        book.save()

Every rule the writer follows was measured, either against live Excel or
against Microsoft's own packaging assemblies; :mod:`_metadata` and
:mod:`_section` say which, case by case.
"""

from pyopenvba.powerquery._files import pull_queries, push_queries
from pyopenvba.powerquery._mashup import Mashup
from pyopenvba.powerquery._metadata import Entry, Item, Metadata, QueryGroup
from pyopenvba.powerquery._package import Package
from pyopenvba.powerquery._section import Section
from pyopenvba.powerquery.workbook import (
    LOAD_CONNECTION_ONLY,
    LOAD_PIVOT_TABLE,
    LOAD_TABLE,
    PowerQuery,
    PowerQueryWorkbook,
)

__all__ = [
    "LOAD_CONNECTION_ONLY",
    "LOAD_PIVOT_TABLE",
    "LOAD_TABLE",
    "Entry",
    "Item",
    "Mashup",
    "Metadata",
    "Package",
    "PowerQuery",
    "PowerQueryWorkbook",
    "QueryGroup",
    "Section",
    "pull_queries",
    "push_queries",
]
