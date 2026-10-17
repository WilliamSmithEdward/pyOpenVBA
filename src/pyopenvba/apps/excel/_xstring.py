"""Text as SpreadsheetML stores it, which is not quite as XML stores it.

XML cannot carry most control characters at all, and a carriage return or
a line break in an attribute does not survive being read. SpreadsheetML's
string type, ST_Xstring, spells such a character ``_xHHHH_`` instead, by
its UTF-16 code in hex, and spells the underscore of anything that would
read as an escape ``_x005F_``, so a cell holding ``_x0041_`` keeps it.

Adapted from William's pyOfficeEditor codec. The engine's native controls
cover literal escapes, control characters, Unicode cell/formula persistence
and pivot attributes. Excel writes element text and attributes differently:

- In element text the hex is uppercase. A tab and a line feed are written
  as they are, a line feed as CRLF, and a carriage return is escaped,
  since an XML reader turns a bare one into a line feed.
- In an attribute the hex is lowercase, and a tab, a line feed and a
  carriage return are all escaped, since an XML reader turns each into a
  space there.

Excel reads either spelling in either place, with the hex digits in either
case. ``_X0041_``, with a capital X, is not an escape, and neither is one
short a digit or its closing underscore. A NUL reads back as one, though
Excel's object model cuts text at a NUL when it is given one.
"""

from __future__ import annotations

import re

#: An escape, read as Excel reads one.
_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")

#: What escaping element text changes: the underscore of anything that
#: would read as an escape, and what XML cannot hold there. A tab and a
#: line feed stay as they are.
_TEXT_CHANGES = re.compile("_(?=x[0-9A-Fa-f]{4}_)|[\x00-\x08\x0b-\x1f\ud800-\udfff\ufffe\uffff]")
#: What writing an attribute changes: the same, and a tab and a line feed.
_ATTRIBUTE_CHANGES = re.compile("_(?=x[0-9A-Fa-f]{4}_)|[\x00-\x1f\ud800-\udfff\ufffe\uffff]")
#: The characters XML cannot hold at all, even as a reference.
_NOT_XML = re.compile("(?![\t\n\r])[\x00-\x1f\ud800-\udfff\ufffe\uffff]")


def decode(text: str) -> str:
    """Stored text as the characters it spells.

    Escapes naming the two halves of a surrogate pair make one character,
    as the UTF-16 Excel counts in does.
    """
    if "_x" not in text:
        return text
    decoded = _ESCAPE.sub(lambda found: chr(int(found.group(1), 16)), text)
    if any("\ud800" <= character <= "\udfff" for character in decoded):
        decoded = decoded.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")
    return decoded


def escape(text: str) -> str:
    """Text with what an element cannot hold spelled as Excel spells it,
    and its line feeds as they are: what an XML reader gives back of text
    Excel wrote. The formula code works on formulas in this form, since
    that is how it reads them from a part."""
    return _TEXT_CHANGES.sub(lambda found: f"_x{ord(found.group(0)):04X}_", text)


def encode_text(text: str) -> str:
    """Text as Excel writes it inside an element: escaped, and each line
    feed written as CRLF."""
    return escape(text).replace("\n", "\r\n")


def encode_attribute(text: str) -> str:
    """Text as Excel writes it in an attribute's value."""
    return _ATTRIBUTE_CHANGES.sub(lambda found: f"_x{ord(found.group(0)):04x}_", text)


def replace_unrepresentable(text: str) -> str:
    """Text with each character XML cannot hold replaced by U+FFFD.

    Where Excel replaces rather than escapes: in a table's headers, which
    it cleans this way when it makes the table, measured for every one of
    those characters. A tab, a line feed and a carriage return it keeps.
    """
    return _NOT_XML.sub("\ufffd", text)


__all__ = ["decode", "encode_attribute", "encode_text", "escape", "replace_unrepresentable"]
