"""The Jet storage engine behind ``.accdb`` and ``.mdb`` files.

The scope is what the Access application opens today: Jet 4 and ACE.
Jet 3 (Access 97, 2 KiB pages) is refused.

Pure Python, no Office required.  ``AccessDatabase`` opens a database and
reads its tables; the private modules beneath it own one layer each:
pages and usage maps, table definitions, row and value codecs, long
values.
"""

from pyopenvba.access._schema import ColumnSpec, IndexSpec
from pyopenvba.access._complex import Attachment, ComplexColumn
from pyopenvba.access._designs import AccessDesign, DesignObject, DesignRecord
from pyopenvba.access._facade import AccessControl, AccessForm, AccessVBAProject
from pyopenvba.access._macros import Macro, MacroAction
from pyopenvba.access._props import PropertyValue
from pyopenvba.access._queries import QueryRow, SavedQuery
from pyopenvba.access._vba import VBAModule
from pyopenvba.access.database import AccessDatabase, CatalogEntry, Index, LinkedTable, Relationship, RowId, Table

__all__ = ["AccessControl", "AccessDatabase", "AccessDesign", "AccessForm", "AccessVBAProject", "Attachment", "CatalogEntry", "ColumnSpec", "ComplexColumn", "DesignObject", "DesignRecord", "Index", "IndexSpec", "LinkedTable", "Macro", "MacroAction", "PropertyValue", "QueryRow", "Relationship", "RowId", "SavedQuery", "Table", "VBAModule"]
