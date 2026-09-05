"""Set the refresh control on a loaded Power Query.

    python examples/power_query_refresh.py [path.xlsx]

These are the boxes in Excel's Connection Properties dialog, under
Refresh control. They belong to the connection a loaded query goes
through, so a query that loads nowhere has none, and `query.refresh`
says so.

The workbook this writes gives three queries three refresh profiles:

* **Snapshot** refreshes when the workbook opens and keeps its rows in
  the file, so the numbers are there before anything runs.
* **Hourly** refreshes every sixty minutes, in the foreground, so a
  refresh finishes before the next line of a macro runs.
* **Transient** keeps no rows in the file at all. It fetches them when
  the workbook opens and drops them again on save, and it stays out of
  Refresh All because it has already refreshed itself.

Excel greys "Remove data from the external data range before saving the
workbook" out until "Refresh data when opening the file" is ticked,
which is why Transient sets both. Setting `keep_data` alone is legal and
Excel keeps it, but the dialog then shows a ticked box greyed out under
an unticked one.

Open the file, right-click a query in Queries & Connections, choose
Properties, and the boxes match the table this prints.
"""

from __future__ import annotations

import sys
from pathlib import Path

from pyopenvba import PowerQueryWorkbook

ROWS = '''let
    Source = Table.FromRecords({[Region = "North", Units = 120], [Region = "South", Units = 98]}),
    Typed = Table.TransformColumnTypes(Source, {{"Region", type text}, {"Units", Int64.Type}})
in
    Typed'''


def build(path: str | Path) -> Path:
    """Write the workbook at `path` and return the path."""
    with PowerQueryWorkbook.create_new(path) as book:
        for name, cell in (("Snapshot", "A1"), ("Hourly", "D1"), ("Transient", "G1")):
            book.add_query(name, ROWS)
            book.load_to_sheet(name, ["Region", "Units"], cell=cell)

        # Ready to read the moment the file opens, and still there offline.
        snapshot = book.query("Snapshot").refresh
        snapshot.on_open = True
        snapshot.keep_data = True

        # On a timer, and in the foreground so a macro can wait for it.
        hourly = book.query("Hourly").refresh
        hourly.interval_minutes = 60
        hourly.background = False

        # Fetched on open and dropped again on save, so the file holds no
        # rows.  It refreshes itself, so it stays out of Refresh All.
        transient = book.query("Transient").refresh
        transient.on_open = True
        transient.keep_data = False
        transient.in_refresh_all = False

        return book.save()


def report(path: str | Path) -> list[str]:
    """One line per query, saying how it refreshes."""
    book = PowerQueryWorkbook(path)
    lines: list[str] = []
    for query in book.queries():
        settings = query.refresh
        every = "never" if settings.interval_minutes is None else f"every {settings.interval_minutes} min"
        lines.append(
            f"{query.name:10} on open: {str(settings.on_open):5}  {every:14} "
            f"background: {str(settings.background):5}  keeps data: {str(settings.keep_data):5}  "
            f"in Refresh All: {str(settings.in_refresh_all):5}  enabled: {settings.enabled}"
        )
    return lines


if __name__ == "__main__":
    written = build(sys.argv[1] if len(sys.argv) > 1 else "power_query_refresh.xlsx")
    print(written)
    for line in report(written):
        print("  " + line)
