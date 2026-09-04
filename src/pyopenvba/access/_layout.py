"""What changes between Jet 3 and Jet 4, in one place.

Jet 4 (Access 2000 and later, and ACE after it) grew the page from 2 KiB
to 4 KiB, widened every count in a row from a byte to a word, moved text
from the database code page to UTF-16, and added fields to the table
definition -- so every offset past the first few moved.  Rather than
sprinkle version checks through the reader, the numbers live here and a
``Layout`` travels with the store and with each table definition.

Every offset and size below was measured, not assumed: a definition is
parsed and the bytes consumed must equal the length the page declares,
and the rows the reader decodes must equal what the engine reports for
the same file.  The Jet 3 numbers came from files DAO 3.6 wrote through
Jet 4.0, which still creates and reads Access 97 databases even though
Access itself dropped the format in 2013.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Layout:
    """The version-dependent shape of a database file."""

    name: str
    page_size: int

    #: Data page header.  Jet 4 has four more bytes before the row count.
    page_row_count: int
    page_row_table: int

    #: Table definition header.
    tdef_tag: int
    tdef_row_count: int
    tdef_next_autonumber: int
    tdef_table_type: int
    tdef_max_columns: int
    tdef_var_column_count: int
    tdef_column_count: int
    tdef_logical_index_count: int
    tdef_real_index_count: int
    tdef_owned_pages: int
    tdef_free_space_pages: int
    tdef_index_headers: int

    size_real_index_header: int
    size_column_header: int
    size_real_index: int
    size_logical_index: int

    #: Column header.
    column_number: int
    column_var_index: int
    column_sort_order: int
    column_sort_version: int
    column_flags: int
    column_fixed_offset: int
    column_length: int

    #: Real index definition.
    index_columns: int
    index_usage_map: int
    index_root_page: int

    #: A row's column count and each of its variable-column offsets: one
    #: byte in Jet 3, two in Jet 4.
    count_width: int
    #: A name in a definition is prefixed by its length the same way.
    name_length_width: int
    #: Jet 4 stores text as UTF-16; Jet 3 stores it in the code page named
    #: on page 0.
    unicode_text: bool

    @property
    def is_jet3(self) -> bool:
        return self.page_size == 2048


JET3 = Layout(
    name="Jet 3",
    page_size=2048,
    page_row_count=0x08,
    page_row_table=0x0A,
    tdef_tag=0x02,
    tdef_row_count=0x0C,
    tdef_next_autonumber=0x10,
    tdef_table_type=0x14,
    tdef_max_columns=0x15,
    tdef_var_column_count=0x17,
    tdef_column_count=0x19,
    tdef_logical_index_count=0x1B,
    tdef_real_index_count=0x1F,
    tdef_owned_pages=0x23,
    tdef_free_space_pages=0x27,
    tdef_index_headers=0x2B,
    size_real_index_header=8,
    size_column_header=18,
    size_real_index=39,
    size_logical_index=20,
    # Jet 4 opens a column header with the table tag and names the column
    # at 5; Jet 3 has no tag and names it at 1.  The word at 5 tracks the
    # column number on a table the engine has just built, which is why only
    # the catalog -- where it is zero throughout -- tells the two apart.
    column_number=1,
    column_var_index=3,
    column_sort_order=9,
    column_sort_version=11,
    column_flags=13,
    column_fixed_offset=14,
    column_length=16,
    index_columns=0,
    index_usage_map=30,
    index_root_page=34,
    count_width=1,
    name_length_width=1,
    unicode_text=False,
)

JET4 = Layout(
    name="Jet 4",
    page_size=4096,
    page_row_count=0x0C,
    page_row_table=0x0E,
    tdef_tag=0x0C,
    tdef_row_count=0x10,
    tdef_next_autonumber=0x14,
    tdef_table_type=0x28,
    tdef_max_columns=0x29,
    tdef_var_column_count=0x2B,
    tdef_column_count=0x2D,
    tdef_logical_index_count=0x2F,
    tdef_real_index_count=0x33,
    tdef_owned_pages=0x37,
    tdef_free_space_pages=0x3B,
    tdef_index_headers=0x3F,
    size_real_index_header=12,
    size_column_header=25,
    size_real_index=52,
    size_logical_index=28,
    column_number=5,
    column_var_index=7,
    column_sort_order=11,
    column_sort_version=13,
    column_flags=15,
    column_fixed_offset=21,
    column_length=23,
    index_columns=4,
    index_usage_map=34,
    index_root_page=38,
    count_width=2,
    name_length_width=2,
    unicode_text=True,
)
