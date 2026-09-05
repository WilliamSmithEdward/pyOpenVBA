"""The refresh settings behind a loaded query.

These are the boxes in Excel's Connection Properties dialog, under
Refresh control. None of them live in the Power Query package: they sit
on the workbook connection the query loads through, and two of them are
mirrored onto the query table beside it.

Where each one is written was measured by toggling it in Excel through
the object model and diffing the file:

| Dialog | Written as |
| --- | --- |
| Enable background refresh | ``connection/@background``, mirrored as ``queryTable/@backgroundRefresh`` |
| Refresh every N minutes | ``connection/@interval`` |
| Refresh data when opening the file | ``connection/@refreshOnLoad``, mirrored as ``queryTable/@refreshOnLoad`` |
| Remove data before saving | ``connection/@saveData``, which Excel drops when the box is ticked |
| Refresh this connection on Refresh All | ``x15:connection/@excludeFromRefreshAll`` in the connection's extension list |
| Enable refresh | ``queryTable/@disableRefresh`` |

Every one of them is absent by default, and absent means off, so a
setting turned off is written by taking the attribute away.

"Enable Fast Data Load", the last box in that group, is not here. Excel's
object model does not expose it, so there was no way to watch Excel write
it, and a setting written on a guess is worse than one left alone.
"""

from __future__ import annotations

import re

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery._opc import OpcFile

_CONNECTIONS = "xl/connections.xml"
#: The extension a connection carries when it is kept out of Refresh All.
_EXT_URI = "{DE250136-89BD-433C-8126-D09CA5730AF9}"
_X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"


def _escape(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _attribute(tag: str, name: str) -> str | None:
    match = re.search(rf'\b{re.escape(name)}="([^"]*)"', tag)
    return None if match is None else match.group(1)


def _with_attribute(tag: str, name: str, value: str | None) -> str:
    """The opening tag with an attribute set, or taken away when `value`
    is None.  A new attribute goes last, which is where Excel puts one."""
    pattern = rf'\s{re.escape(name)}="[^"]*"'
    stripped = re.sub(pattern, "", tag)
    if value is None:
        return stripped
    close = "/>" if stripped.rstrip().endswith("/>") else ">"
    body = stripped.rstrip()[: -len(close)].rstrip()
    return f'{body} {name}="{_escape(value)}"{close}'


class RefreshSettings:
    """Refresh control for one loaded query.

    Reading a property answers what the file says; setting one writes it
    through to the workbook straight away, the way the rest of this
    package works.
    """

    def __init__(self, package: OpcFile, query: str) -> None:
        self._package = package
        self._query = query
        if self._connection() is None:
            raise PowerQueryError(
                f"the query {query!r} loads nowhere, so it has no connection to refresh; "
                "load it to a sheet first"
            )

    # -- finding the two parts ---------------------------------------------

    def _connection(self) -> str | None:
        """The whole ``<connection>`` element for this query, or None."""
        if not self._package.has(_CONNECTIONS):
            return None
        raw = self._package.read(_CONNECTIONS).decode("utf-8")
        for block in re.findall(r"<connection\b.*?</connection>|<connection\b[^>]*/>", raw, re.S):
            if f"Location={_escape(self._query)};" in block:
                return block
        return None

    def _write_connection(self, block: str) -> None:
        raw = self._package.read(_CONNECTIONS).decode("utf-8")
        current = self._connection()
        if current is None:  # pragma: no cover - the constructor checked
            raise PowerQueryError(f"the query {self._query!r} has no connection")
        self._package.write(_CONNECTIONS, raw.replace(current, block, 1).encode("utf-8"))

    @property
    def _identifier(self) -> str | None:
        block = self._connection()
        return None if block is None else _attribute(block, "id")

    def _query_table_part(self) -> str | None:
        identifier = self._identifier
        if identifier is None:
            return None
        for name in self._package.names():
            if not re.match(r"xl/queryTables/queryTable\d+\.xml$", name):
                continue
            if f'connectionId="{identifier}"' in self._package.read(name).decode("utf-8"):
                return name
        return None

    def _query_table_flag(self, name: str) -> str | None:
        part = self._query_table_part()
        if part is None:
            return None
        raw = self._package.read(part).decode("utf-8")
        opening = re.search(r"<queryTable\b[^>]*>", raw)
        return None if opening is None else _attribute(opening.group(0), name)

    def _set_query_table_flag(self, name: str, value: str | None) -> None:
        part = self._query_table_part()
        if part is None:
            return
        raw = self._package.read(part).decode("utf-8")
        opening = re.search(r"<queryTable\b[^>]*>", raw)
        if opening is None:  # pragma: no cover - written by us or by Excel
            return
        fresh = _with_attribute(opening.group(0), name, value)
        self._package.write(part, raw.replace(opening.group(0), fresh, 1).encode("utf-8"))

    def _connection_flag(self, name: str) -> str | None:
        block = self._connection()
        if block is None:  # pragma: no cover - the constructor checked
            return None
        opening = re.search(r"<connection\b[^>]*>", block)
        return None if opening is None else _attribute(opening.group(0), name)

    def _set_connection_flag(self, name: str, value: str | None) -> None:
        block = self._connection()
        if block is None:  # pragma: no cover - the constructor checked
            return
        opening = re.search(r"<connection\b[^>]*>", block)
        if opening is None:  # pragma: no cover
            return
        self._write_connection(block.replace(opening.group(0), _with_attribute(opening.group(0), name, value), 1))

    # -- the settings -------------------------------------------------------

    @property
    def background(self) -> bool:
        """Whether a refresh runs in the background."""
        return self._connection_flag("background") == "1"

    @background.setter
    def background(self, value: bool) -> None:  # noqa: FBT001 - mirrors the dialog
        self._set_connection_flag("background", "1" if value else None)
        self._set_query_table_flag("backgroundRefresh", None if value else "0")

    @property
    def interval_minutes(self) -> int | None:
        """Minutes between timed refreshes, or None when there are none."""
        value = self._connection_flag("interval")
        return None if value in (None, "0") else int(str(value))

    @interval_minutes.setter
    def interval_minutes(self, value: int | None) -> None:
        if value is not None and value <= 0:
            raise PowerQueryError("a refresh interval is a positive number of minutes, or None")
        self._set_connection_flag("interval", None if value is None else str(value))

    @property
    def on_open(self) -> bool:
        """Whether the query refreshes when the workbook opens."""
        return self._connection_flag("refreshOnLoad") == "1"

    @on_open.setter
    def on_open(self, value: bool) -> None:  # noqa: FBT001 - mirrors the dialog
        self._set_connection_flag("refreshOnLoad", "1" if value else None)
        self._set_query_table_flag("refreshOnLoad", "1" if value else None)

    @property
    def keep_data(self) -> bool:
        """Whether the result is saved with the workbook.

        The dialog puts this the other way round: its box removes the data
        before saving, so it is ticked when this is False.  What Excel
        reads is ``queryTable/@removeDataOnSave``; the connection's
        ``saveData`` follows it, which is how Excel writes the pair.
        """
        return self._query_table_flag("removeDataOnSave") != "1"

    @keep_data.setter
    def keep_data(self, value: bool) -> None:  # noqa: FBT001 - mirrors the dialog
        self._set_query_table_flag("removeDataOnSave", None if value else "1")
        self._set_connection_flag("saveData", "1" if value else None)

    @property
    def in_refresh_all(self) -> bool:
        """Whether Refresh All refreshes this query."""
        block = self._connection()
        return not (block is not None and 'excludeFromRefreshAll="1"' in block)

    @in_refresh_all.setter
    def in_refresh_all(self, value: bool) -> None:  # noqa: FBT001 - mirrors the dialog
        block = self._connection()
        if block is None:  # pragma: no cover - the constructor checked
            return
        without = re.sub(r"<extLst>.*?</extLst>", "", block, flags=re.S)
        if value:
            self._write_connection(without)
            return
        if without.rstrip().endswith("/>"):
            opening = re.search(r"<connection\b[^>]*/>", without)
            if opening is None:  # pragma: no cover
                return
            without = opening.group(0)[:-2].rstrip() + "></connection>"
        extension = (
            f'<extLst><ext uri="{_EXT_URI}" xmlns:x15="{_X15}">'
            '<x15:connection id="" excludeFromRefreshAll="1"/></ext></extLst>'
        )
        self._write_connection(without.replace("</connection>", extension + "</connection>", 1))

    @property
    def enabled(self) -> bool:
        """Whether the query may be refreshed at all."""
        return self._query_table_flag("disableRefresh") != "1"

    @enabled.setter
    def enabled(self, value: bool) -> None:  # noqa: FBT001 - mirrors the dialog
        self._set_query_table_flag("disableRefresh", None if value else "1")

    def __repr__(self) -> str:
        return (
            f"RefreshSettings({self._query!r}, background={self.background}, "
            f"interval_minutes={self.interval_minutes}, on_open={self.on_open}, "
            f"keep_data={self.keep_data}, in_refresh_all={self.in_refresh_all}, "
            f"enabled={self.enabled})"
        )
