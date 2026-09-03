"""The Jet 4 / ACE storage engine behind ``.accdb`` and ``.mdb`` files.

Pure Python, no Office required.  ``AccessDatabase`` opens a database and
reads its tables; the private modules beneath it own one layer each:
pages and usage maps, table definitions, row and value codecs, long
values.
"""

from pyopenvba.access._schema import ColumnSpec, IndexSpec
from pyopenvba.access._props import PropertyValue
from pyopenvba.access._queries import QueryRow, SavedQuery
from pyopenvba.access._vba import VBAModule
from pyopenvba.access.database import AccessDatabase, CatalogEntry, Index, LinkedTable, Relationship, RowId, Table

__all__ = ["AccessDatabase", "CatalogEntry", "ColumnSpec", "Index", "IndexSpec", "LinkedTable", "PropertyValue", "QueryRow", "Relationship", "RowId", "SavedQuery", "Table", "VBAModule"]
