let
    Source = #table({"A"}, {{1}, {2}}),
    Kept = Table.SelectRows(Source, each [A] > 1)
in
    Kept
