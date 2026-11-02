# The formula engine

Cells are worked out by `pyopenvba.formula._calc`, pyOfficeEditor's
formula engine (`pyofficeeditor.excel._calc`), copied in from its commit
4723383 with its imports pointed at pyOpenVBA. Its modules keep their
names, so the next copy is a diff of the same files:

```console
python scripts/copy_formula_engine.py [path to pyOfficeEditor]
```

prints how each module differs from pyOfficeEditor's committed one, and
`--write` copies them over. Read the differences first: the copy carries
the changes below, each made for one of pyOpenVBA's live measurements,
and a new copy has to keep them or show they are no longer needed. A
module both projects changed is merged three ways, pyOfficeEditor's
module at the commit last copied being the base, rewritten as the script
rewrites it:

```console
python scripts/copy_formula_engine.py --merge 4723383 [path to pyOfficeEditor]
```

merges every module so and leaves each conflict marked. The merge from
81481d2 to 4723383 brought in ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD,
GROUPBY, PIVOTBY and BAHTTEXT, YIELD, PMT, IPMT, PPMT, CUMIPMT and
CUMPRINC to the bit, Kahan's e^x - 1, functions passed by name
(`_xleta.SUM`), dates and times read from text as Excel reads them, and
reference-only parameters taking an array one item at a time. Its six
conflicts were the LAMBDA both projects had given optional parameters
(pyOfficeEditor's count of required ones is kept, with pyOpenVBA's `[y]`
read as one), TEXTJOIN, TRANSPOSE and an import. pyOfficeEditor's own
workbook driver learned circular references, data tables, linked
workbooks' cached values and error checking in those commits; pyOpenVBA
does not use that driver, so those wait for its own.
`tests/test_excel_formula_corpus.py` runs pyOfficeEditor's own corpus
through the copy, and `tests/test_formula_engine_sync.py` the tests
pyOfficeEditor wrote for the commits merged, so both sets of
measurements hold at once.

`cells.py` and `host.py` are pyOpenVBA's own: TEXT, DOLLAR and FIXED
format through the display engine `Range.Text` uses.
`apps/excel/_engine_book.py` gives the engine the model's workbook to
read, and hands it each formula spelled as a file spells it but with a
column's `@` kept escaped: a file writes a column named `@home` as
`T[@home]`, which the engine's lexer reads as this row. A name a LET or
a LAMBDA binds comes to it `_xlpm.x1`, as a file writes it, so that a
name looking like a cell is not read as the cell.

## pyOpenVBA's changes

1. **Legacy formulas.** Every formula but an array formula is one:
   Range.Formula writes them, and a file keeps a formula that needs
   dynamic arrays as `t="array"`. `Context.legacy` marks one, and
   `registry.function` declares where it differs (`tests/fixtures/formula/`):
   - `legacy_first`: an array given where one value is wanted gives its
     first item, even inside SUMPRODUCT: INDEX (1, 2, 3), VLOOKUP and
     HLOOKUP (0, 2, 3), XLOOKUP (0), IFERROR and IFNA (0, 1).
   - `legacy_as_cell`: IF, IFS, IFERROR, IFNA, SWITCH and CHOOSE work
     their arguments out as a cell does inside an argument worked out
     whole: `SUMPRODUCT(IF(A1:A3>1,1,0))` is #VALUE! in row 5, 1 in row 2.
   - `legacy_cell`: TRANSPOSE's argument is worked out as a cell works a
     value out, a range cut to the formula's row or column, even inside an
     argument worked out whole: `INDEX(TRANSPOSE(A1:A3),1,2)` is #VALUE! in
     row 6.
   - `legacy_corner`: `N(A1:A3)` is A1's value.
2. **Defined names.** A name's formula is worked out as an array formula,
   legacy rules off, and its last sum snaps as a cell's does:
   `SUM(dbl)` with dbl `=$A$1:$A$3*2` is 12 in any row, `=dbl` shows 2.
3. **Lookups compare to the bit.** MATCH, VLOOKUP, HLOOKUP, XLOOKUP and
   XMATCH, exact and approximate, where `=`, COUNTIF, SUMIF and SWITCH
   read fifteen digits (`tests/fixtures/zero_snap.json`).
4. **Tiny sums.** A `+` or `-` too small to be normal keeps its bits and
   is not snapped to zero; `^` and text still flush such a number.
   SUMIF and SUMIFS add in order with no last addition snapped.
5. **A number as text** loses mantissa digits until it fits twenty
   characters: `1.23456789012345E+100` is `1.2345678901235E+100`
   (`tests/fixtures/number_spelling.json`).
6. **Tables.** A special item a table does not show is left out, and
   only with nothing left is the reference #REF!; this row is the
   formula's row on any sheet (`tests/fixtures/structured_references/`).
7. **Functions.** RIGHT past the text's length is the whole text; MATCH
   over a single value is #N/A; CHOOSE with an array of positions picks
   item by item, lined up as an operator lines arrays up; INDEX's first
   argument, ROWS and COLUMNS are worked out whole; INDIRECT reads an R1C1
   range; AND, OR and XOR pass over a text argument that says neither
   TRUE nor FALSE, `AND(TRUE,"x")` being TRUE, and read an argument left
   out as FALSE, `AND(TRUE,)` and `OR(,)` being FALSE
   (`tests/fixtures/vba_worksheet_function_args.json`); ROW and COLUMN propagate error arguments, including #REF! after a CSE copy
   invalidates a reference (`tests/fixtures/cse_array_edits.json`);
   ROW() and COLUMN() with no
   argument are an array of one in an array formula; INDEX with one index
   into cells in rows and columns is #REF!, `INDEX(A1:B2,2,)` being row 2;
   BINOM.DIST and the functions built on it take 0 to the power 0 as 1
   (`tests/fixtures/formula/cells_functions.json`); the database functions
   refuse labels alone, a database or criteria of one row, with #VALUE!;
   ACCRINTM settled the day it is issued is 0; CELL with an info type that
   is not one of its words, `CELL(1)`, is #VALUE!; BINOM.INV and CRITBINOM
   want a chance and an alpha strictly between 0 and 1; XIRR of one
   payment is #N/A; TEXTJOIN takes 254 arguments, not 252; and BYROW and
   BYCOL given no function are #CALC!
   (`tests/fixtures/formula_refusals.json`); LAMBDA's parameter in
   brackets, `[y]` or a file's `_xlop.y`, may be left out of a call, and
   leaving out any other is #VALUE! (`tests/fixtures/bound_names.json`);
   EDATE keeps the 1900 calendar's day 0, so `EDATE(0,1)` is 31
   (`tests/fixtures/date1904/paths.json`); a date given as text in an
   array is read as a date in XIRR's and XNPV's dates and in the holidays
   of NETWORKDAYS, WORKDAY and their INTL forms, as
   `XIRR({..},{"1/1/2008",..})` and `NETWORKDAYS(..,{"11/22/2012"})` read
   it in a cell and VBA hands a Date in an array over
   (`tests/fixtures/worksheet_functions_more.json`).
8. **Values from VBA.** `nodes.Given` carries a value into a call from
   outside any formula, as WorksheetFunction hands one: a scalar, an array
   with blanks in it, or cells. `values.Omitted`, what an argument left
   empty evaluates to, is blank to every formula (`CHOOSE(1,)&"x"` is x)
   and reads back in VBA as 0, where a blank cell reads as Empty.
9. **A `+` sign** reads cells as values, as `-` does: `ISREF(+A1)` is
   FALSE, and `=+A1:A3` is the cell in the formula's row.
10. **A name in quotes.** `'A1'` and `Data!'A1'` are names, the way A1
    spells a name R1C1 read where A1 reads a cell; with no such name
    defined they are #NAME? (`tests/fixtures/formula_notation.json`).
11. **An error where only a range will do** is handed on: `ROW(#REF!)`
    and `LET(x,NA(),COUNTIF(x,1))` are the error, as Excel gives them for
    every reference-only parameter measured
    (`tests/fixtures/formula/reference_errors.json`); pyOfficeEditor makes
    them #VALUE!. An array there is still one #VALUE! an item.
12. **The workbook.** The `Book` protocol leaves out pyOfficeEditor's
    `pivot_reports`, since GETPIVOTDATA is pyOpenVBA's own
    (`functions/pivotdata.py`). The model's book answers `external_sheets`
    and `external_name` from the copy of a linked workbook's cells and
    names the file keeps (`apps/excel/_external_links.py`), and reports a
    workbook it keeps no copy of unsupported
    (`tests/fixtures/external_links.json`).
13. **3D references.** `Reference.three_d` marks cells on two or more
    sheets, `Jan:Mar!A1`, which are no value: where one is wanted, alone,
    with an operator or behind `@`, they are #REF!, and no union,
    intersection or range takes them. `registry.THREE_D_READ` lists the 62
    functions that read every sheet, SUM, LARGE, NPV, CONCAT, VSTACK and
    TOCOL among them, and `THREE_D_ARRAY` the 41 that give #REF!, the ones
    taking an array, SUMPRODUCT, MMULT, SORT, with ANCHORARRAY and SINGLE;
    any other function given one for a range, and N and T, give #VALUE!.
    VSTACK, HSTACK, TOCOL and TOROW take each area of a union or a 3D
    reference in turn, and CONCAT and TEXTJOIN each sheet's. ISREF calls
    one no reference, and INDIRECT's text naming one is #REF!. The lexer
    reads `Q1:Q4!A1`, a first sheet named like a cell, as the cell Q1 to
    `Q4!A1`, and in brackets of their own `(ThreeA:ThreeB!A1)` as the name
    ThreeA to `ThreeB!A1` (`tests/fixtures/three_d_references.json`,
    `three_d_functions.json`, every function that takes a range measured).
14. **Sheet names past ASCII.** The lexer reads a sheet's name written
    without quotes with any character past ASCII in it, as Excel writes
    `x–y!B3` and `ＡＢ!B3` (`tests/fixtures/sheet_quoting.json`), where
    pyOfficeEditor's reads only Unicode letters and digits.
15. **Cells read as Excel reads them.** A circular reference is the cells a
    formula reads as it is worked out, so what is never read closes no
    loop (`tests/fixtures/self_reference.json`). IFERROR and IFNA work their
    fallback out only when the first argument holds an error, or #N/A, in
    a legacy formula or where the fallback is one cell: `IFERROR(1,A5)` in
    A5 is 1. An exact MATCH, VLOOKUP or HLOOKUP over cells reads them in
    order and stops at its match, through the model's book handing cells
    over one at a time: `MATCH(3,A1:A10,0)` in A5 is 3.
16. **Combining accents in names.** The lexer reads a combining accent
    (U+0300 to U+036F) after a name's first character, as Excel's defined
    names hold one: `f2e` with U+0301 is a name
    (`tests/fixtures/name_equality.json`, marks), where pyOfficeEditor's
    stops before the accent.
17. **The Regional format.** `Context.locale` is the workbook's
    application's (`ExcelApplication(locale=...)`), en-US for a book that
    has none. TEXT, DOLLAR and FIXED pass it to the display engine, TEXT
    asking for its code as written. `values.text_to_number` and
    `values.read_date_time`, which DATEVALUE and TIMEVALUE now call for
    `dates.parse_date_time`, read text under it. Under ja-JP the currency
    sign is ¥, three numbers are the year, the month and the day and
    nothing else, and a month's name reports itself unsupported
    (`tests/fixtures/excel_locale_live.json`).
