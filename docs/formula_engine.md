# The formula engine

Cells are worked out by `pyopenvba.formula._calc`, pyOfficeEditor's
formula engine (`pyofficeeditor.excel._calc`), copied in from its commit
ed27bf2 with its imports pointed at pyOpenVBA. Its modules keep their
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
rewrites it: `git merge-file` took `functions/financial.py` from 098e441
to ed27bf2 so.
`tests/test_excel_formula_corpus.py` runs pyOfficeEditor's own corpus
through the copy, so both sets of measurements hold at once.

`cells.py` and `host.py` are pyOpenVBA's own: TEXT, DOLLAR and FIXED
format through the display engine `Range.Text` uses.
`apps/excel/_engine_book.py` gives the engine the model's workbook to
read, and hands it each formula spelled as a file spells it but with a
column's `@` kept escaped: a file writes a column named `@home` as
`T[@home]`, which the engine's lexer reads as this row.

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
   - `legacy_cell`: TRANSPOSE's argument is not worked out whole.
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
   TRUE nor FALSE, `AND(TRUE,"x")` being TRUE; ROW() and COLUMN() with no
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
   (`tests/fixtures/formula_refusals.json`).
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
