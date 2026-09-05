"""Build a demo workbook whose Power Queries call public web APIs.

Nothing here needs Excel, and nothing here needs the network: the queries
are M source, written straight into the workbook's Power Query package.
Excel runs them when it refreshes.

    python examples/power_query_demo.py [path.xlsx]

The workbook it writes holds four groups of queries:

* **Parameters** -- two parameter queries the others read.
* **Functions** -- a reusable fetcher over the PokeAPI, and a small text
  helper.
* **Web** -- one query per JSON shape, because reading JSON in M is
  mostly a question of what shape came back: a list of records, one
  request per row, a list of bare numbers, nested features, a record
  whose field names are the data, and a record of records.
* **Workbook** -- a table written in M and a summary grouped from it, so
  the file has something to show with no network at all.

Six of them are loaded onto the first sheet, side by side. Open the file
and press Refresh All. The first time Excel meets a host it may ask for
anonymous access; that permission belongs to Excel, not to the file.

Every query here was refreshed in Excel before it was committed. The
APIs are public and need no key:

* https://pokeapi.co
* https://earthquake.usgs.gov
* https://api.frankfurter.app
* https://hacker-news.firebaseio.com
"""

from __future__ import annotations

import sys
from pathlib import Path

from pyopenvba import PowerQueryWorkbook

# --- parameters ---------------------------------------------------------------

POKEMON_COUNT = """20 meta [
    IsParameterQuery = true,
    Type = "Number",
    IsParameterQueryRequired = true
]"""

STORY_COUNT = """10 meta [
    IsParameterQuery = true,
    Type = "Number",
    IsParameterQueryRequired = true
]"""

# --- functions ----------------------------------------------------------------

#: A base URL with a RelativePath is the form to reach for: the data
#: source stays one constant host, which is what lets Excel refresh a
#: query that builds its own request.
POKE_API = '''(relative as text) as any =>
    let
        Response = Web.Contents("https://pokeapi.co/api/v2/", [RelativePath = relative]),
        Parsed = Json.Document(Response)
    in
        Parsed'''

TITLE_CASE = '''(value as text) as text => Text.Proper(Text.Replace(value, "-", " "))'''

# --- one query per JSON shape --------------------------------------------------

#: {"results": [{"name": ..., "url": ...}, ...]} -- a list of records.
POKEDEX = '''let
    Source = GetFromPokeApi("pokemon?limit=" & Text.From(PokemonCount)),
    Results = Source[results],
    AsTable = Table.FromList(Results, Splitter.SplitByNothing(), {"Item"}),
    Expanded = Table.ExpandRecordColumn(AsTable, "Item", {"name", "url"}, {"Slug", "Url"}),
    Numbered = Table.AddIndexColumn(Expanded, "Number", 1, 1, Int64.Type),
    Named = Table.AddColumn(Numbered, "Name", each TitleCase([Slug]), type text),
    Kept = Table.SelectColumns(Named, {"Number", "Name", "Slug"})
in
    Kept'''

#: One request per row, and a nested list flattened into text.  The query
#: fetches its own list instead of reading Pokedex, because a query that
#: both names another query and reaches a data source is one Excel
#: refuses to refresh until its privacy levels are sorted out.
POKEMON_STATS = '''let
    Source = GetFromPokeApi("pokemon?limit=8"),
    Results = Source[results],
    AsTable = Table.FromList(Results, Splitter.SplitByNothing(), {"Item"}),
    Expanded = Table.ExpandRecordColumn(AsTable, "Item", {"name"}, {"Slug"}),
    Named = Table.AddColumn(Expanded, "Name", each TitleCase([Slug]), type text),
    Detail = Table.AddColumn(Named, "Detail", each GetFromPokeApi("pokemon/" & [Slug])),
    Height = Table.AddColumn(Detail, "Height (cm)", each [Detail][height] * 10, Int64.Type),
    Weight = Table.AddColumn(Height, "Weight (kg)", each [Detail][weight] / 10, type number),
    Types = Table.AddColumn(Weight, "Types", each
        Text.Combine(List.Transform([Detail][types], each TitleCase(_[type][name])), ", "), type text),
    Kept = Table.SelectColumns(Types, {"Name", "Height (cm)", "Weight (kg)", "Types"})
in
    Kept'''

#: [49570669, 49563355, ...] -- a list of bare numbers, no records at all.
TOP_STORY_IDS = '''let
    Source = Json.Document(Web.Contents("https://hacker-news.firebaseio.com/v0/topstories.json")),
    Top = List.FirstN(Source, StoryCount),
    AsTable = Table.FromList(Top, Splitter.SplitByNothing(), {"Story"}),
    Typed = Table.TransformColumnTypes(AsTable, {{"Story", Int64.Type}}),
    Ranked = Table.AddIndexColumn(Typed, "Rank", 1, 1, Int64.Type),
    Kept = Table.SelectColumns(Ranked, {"Rank", "Story"})
in
    Kept'''

#: GeoJSON: {"features": [{"properties": {...}, ...}]} -- a record inside
#: every row, expanded a level down.
EARTHQUAKES = '''let
    Source = Json.Document(Web.Contents(
        "https://earthquake.usgs.gov/earthquakes/feed/v1.0/summary/2.5_day.geojson")),
    Features = Source[features],
    AsTable = Table.FromList(Features, Splitter.SplitByNothing(), {"Feature"}),
    Properties = Table.AddColumn(AsTable, "Props", each [Feature][properties]),
    Expanded = Table.ExpandRecordColumn(Properties, "Props",
        {"place", "mag", "time"}, {"Place", "Magnitude", "Epoch"}),
    Moment = Table.AddColumn(Expanded, "When (UTC)",
        each #datetime(1970, 1, 1, 0, 0, 0) + #duration(0, 0, 0, [Epoch] / 1000), type datetime),
    Kept = Table.SelectColumns(Moment, {"Place", "Magnitude", "When (UTC)"}),
    Sorted = Table.Sort(Kept, {{"Magnitude", Order.Descending}})
in
    Sorted'''

#: {"rates": {"AUD": 1.38, "BGN": 1.71, ...}} -- the field names are the
#: data, so the record turns into rows.
RATES = '''let
    Source = Json.Document(Web.Contents("https://api.frankfurter.app/latest", [Query = [from = "USD"]])),
    Rates = Source[rates],
    AsTable = Record.ToTable(Rates),
    Renamed = Table.RenameColumns(AsTable, {{"Name", "Currency"}, {"Value", "Rate per USD"}}),
    Typed = Table.TransformColumnTypes(Renamed, {{"Currency", type text}, {"Rate per USD", type number}}),
    Sorted = Table.Sort(Typed, {{"Currency", Order.Ascending}})
in
    Sorted'''

#: {"rates": {"2025-01-02": {"EUR": ..., "GBP": ...}, ...}} -- a record of
#: records: rows from the outer one, columns from the inner.
RATE_HISTORY = '''let
    Source = Json.Document(Web.Contents("https://api.frankfurter.app/",
        [RelativePath = "2025-01-02..2025-01-10", Query = [from = "USD", to = "EUR,GBP,JPY"]])),
    Rates = Source[rates],
    AsTable = Record.ToTable(Rates),
    Renamed = Table.RenameColumns(AsTable, {{"Name", "Date"}}),
    Expanded = Table.ExpandRecordColumn(Renamed, "Value", {"EUR", "GBP", "JPY"}),
    Typed = Table.TransformColumnTypes(Expanded,
        {{"Date", type date}, {"EUR", type number}, {"GBP", type number}, {"JPY", type number}}),
    Sorted = Table.Sort(Typed, {{"Date", Order.Ascending}})
in
    Sorted'''

# --- the queries that need nothing but the workbook ----------------------------

ORDERS = '''let
    Source = Table.FromRecords({
        [Order = 1001, Product = "Widget", Quantity = 12, Unit = 9.99],
        [Order = 1002, Product = "Gadget", Quantity = 3, Unit = 24.5],
        [Order = 1003, Product = "Widget", Quantity = 7, Unit = 9.99],
        [Order = 1004, Product = "Doohickey", Quantity = 25, Unit = 1.75],
        [Order = 1005, Product = "Gadget", Quantity = 9, Unit = 24.5]
    }),
    Typed = Table.TransformColumnTypes(Source, {
        {"Order", Int64.Type},
        {"Product", type text},
        {"Quantity", Int64.Type},
        {"Unit", Currency.Type}}),
    Totalled = Table.AddColumn(Typed, "Total", each [Quantity] * [Unit], Currency.Type)
in
    Totalled'''

ORDER_SUMMARY = '''let
    Grouped = Table.Group(Orders, {"Product"}, {
        {"Orders", each Table.RowCount(_), Int64.Type},
        {"Units", each List.Sum([Quantity]), Int64.Type},
        {"Revenue", each List.Sum([Total]), Currency.Type}}),
    Sorted = Table.Sort(Grouped, {{"Revenue", Order.Descending}})
in
    Sorted'''


def build(path: str | Path) -> Path:
    """Write the demo workbook at `path` and return the path."""
    with PowerQueryWorkbook.create_new(path) as book:
        parameters = book.add_group("Parameters")
        functions = book.add_group("Functions")
        web = book.add_group("Web")
        workbook = book.add_group("Workbook")

        book.add_query(
            "PokemonCount", POKEMON_COUNT, group=parameters,
            description="How many Pokemon the Pokedex query asks for.",
        )
        book.add_query(
            "StoryCount", STORY_COUNT, group=parameters,
            description="How many Hacker News stories to keep.",
        )

        book.add_query(
            "GetFromPokeApi", POKE_API, group=functions,
            description="Fetch one path from pokeapi.co and parse the JSON.",
        )
        book.add_query(
            "TitleCase", TITLE_CASE, group=functions,
            description="charizard becomes Charizard, and mr-mime becomes Mr Mime.",
        )

        book.add_query(
            "Pokedex", POKEDEX, group=web,
            description="A JSON list of records, and a parameter in the URL.",
        )
        book.add_query(
            "PokemonStats", POKEMON_STATS, group=web,
            description="One request per row, then a nested list flattened into text.",
        )
        book.add_query(
            "TopStoryIds", TOP_STORY_IDS, group=web,
            description="A JSON list of bare numbers, with no records in it.",
        )
        book.add_query(
            "Earthquakes", EARTHQUAKES, group=web,
            description="GeoJSON: a record of properties inside every feature.",
        )
        book.add_query(
            "Rates", RATES, group=web,
            description="A record whose field names are the data.",
        )
        book.add_query(
            "RateHistory", RATE_HISTORY, group=web,
            description="A record of records: rows from one, columns from the other.",
        )

        book.add_query(
            "Orders", ORDERS, group=workbook,
            description="A table written in M, so the workbook shows something offline.",
        )
        book.add_query(
            "OrderSummary", ORDER_SUMMARY, group=workbook,
            description="Grouped from Orders, which it names directly.",
        )

        # Six of them go onto the first sheet, side by side.  Excel settles
        # the columns and fills the rows on the first refresh.
        book.load_to_sheet("Pokedex", ["Number", "Name", "Slug"], cell="A1")
        book.load_to_sheet(
            "PokemonStats", ["Name", "Height (cm)", "Weight (kg)", "Types"], cell="E1"
        )
        book.load_to_sheet("Rates", ["Currency", "Rate per USD"], cell="K1")
        book.load_to_sheet("RateHistory", ["Date", "EUR", "GBP", "JPY"], cell="N1")
        book.load_to_sheet("OrderSummary", ["Product", "Orders", "Units", "Revenue"], cell="T1")
        book.load_to_sheet("Earthquakes", ["Place", "Magnitude", "When (UTC)"], cell="Z1")
        return book.save()


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "power_query_demo.xlsx"))
