# Changelog

All notable changes to pyOpenVBA are documented here. This project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Cells are typed as Excel types them. A string written through `Value`,
  `Value2`, `Formula`, `FormulaR1C1` or an array brings the number format
  typing it gives -- `"5%"` 0%, `"$1,000.50"` a dollar format, `"1e3"`
  scientific, `"1 3/16"` a fraction, times with seconds, tenths and
  AM/PM, dates month first on Excel's 1900 calendar, a run of spaces
  between a date's or a time's parts read as one -- and a Date or
  Currency a macro computes brings its own. A cell keeps a format of its
  own unless it is General or a built-in one of another kind, reads
  `"1/2"` as a half when its format is a number's, and keeps every string
  as text when its format is Text, where a Date or Currency becomes the
  text Excel shows. `Value` reads a number as a Date or Currency through
  the cell's format, raising error 6 where one cannot hold it, and
  `Value2` as a Double. A number typed with more than fifteen significant
  digits keeps fifteen, the rest cut off rather than rounded. 3,947
  writes in live Excel and a workbook it saved pin the rules.
- `Range.PrefixCharacter`: a leading apostrophe keeps the rest as text
  and sets the cell's prefix flag, which its format keeps until `Clear` or
  `ClearFormats`, and which a save writes as Excel's `quotePrefix`.
- `Range.HasFormula`.
- `NumberFormat` keeps a code as Excel rewrites it -- a quoted `"$"` bare,
  an escaped `\-` or `\(` bare, codes in their case, a lower-case exponent
  refused with error 1004 -- and the file spells a code as Excel's does,
  with the `$` quoted, literal dashes, brackets and spaces escaped, and
  the currency and accounting built-ins (ids 5 to 8 and 41 to 44) written
  out. A file's own spelling reads back as Excel reads it.
- `Range.Text` under General shows what Excel's eleven characters show:
  the number in full while it fits, else rounded or given an exponent,
  whichever keeps more digits.
- Excel's number-format engine behind `Range.Text`, the `TEXT` function
  and the new `WorksheetFunction.Text`: sections and conditions, digit
  placeholders, grouping and scaling commas, percent, exponents,
  fractions, text sections, dates and times on Excel's 1900 calendar,
  and elapsed times. 156 codes through 28 values in live Excel are
  reproduced exactly. A time rounds to the second, or to the fraction of
  one shown, as Excel rounds it: half a unit is added to the serial before
  the hours, minutes and seconds are cut off it, which 152 times on the
  half second in live Excel pin. A value its format cannot show fills
  the cell with `#` as wide as the column, and makes `TEXT` an error. A
  format with a `*` fill reports itself unsupported in `Range.Text`, since
  what it pads the cell with depends on the width in pixels. Before, both went
  through Access's `Format`, which spelled `$#,##0_)` as `$5_)`, fractions
  as `2 ?/?` and `[h]:mm:ss` as `[1]:00:00`.
- AutoFilter: `Range.AutoFilter`, `Worksheet.AutoFilter`, `AutoFilterMode`,
  `FilterMode` and `ShowAllData`, and the AutoFilter, Filters and Filter
  objects a macro reads them through. A criterion with `=` or none matches
  the text a cell shows, ignoring case, with `*`, `?` and `~`, so `1`
  finds 1 but not 1 shown as 1.00, and a wildcard finds text only; `<>`,
  `>`, `>=`, `<` and `<=` compare numbers and dates by value and text as
  Sort orders it. xlAnd and xlOr, lists with xlFilterValues, top and
  bottom items and percentages, and above and below average work as
  Excel's do, down to how each criterion reads back (`>1/15/2020` is
  `>43845`, `>2.50` is `>2.5`) and the field a list returns. Filtering
  shows and hides every row under the header; ShowAllData and turning
  the filter off show them again. The file holds each filter as Excel
  writes it, with the hidden `_FilterDatabase` name, and a filter in a
  file reads back as Excel reads it: 113 layouts and 34 saved filters
  in live Excel pin it, and Excel opens the filters this writes. Colour,
  icon and date-group filters in a file are kept but not applied.
- Edits on a filtered sheet hold back to what shows, as Excel's do.
  Values, formulas, formats, borders, clearing, fills, Replace, Sort, row
  heights and hiding reach only the visible cells, and a range with no
  visible cell is edited whole. Copy takes the visible cells, closed up,
  and pastes their formulas as the values they show; a copy of several
  areas does the same anywhere. Deleting cells up deletes the visible
  rows whole, inserting cells is error 1004, and inserting whole rows
  inserts as many as the range shows. The filter follows the sheet: rows
  and columns inserted and deleted move and stretch it, deleting a
  filtered column filters again by the rest, and deleting or clearing
  the header row turns it off. `AutoFilter.Range` runs on over rows
  written below it, which filtering again takes in. 195 layouts in live
  Excel pin it; a two-dimensional array written to several areas, which
  Excel reads at another stride, reports itself unsupported.
- SUBTOTAL, in cells and through `WorksheetFunction.Subtotal`. 1 to 11
  pass over the rows a filter hid -- and while a sheet is in filter mode
  Excel takes every hidden row there for one -- 101 to 111 over every
  hidden row, and hidden columns count. A cell that holds SUBTOTAL or
  AGGREGATE is passed over. The function number is cut to a whole one,
  TRUE is 1, and anything outside 1 to 11 and 101 to 111 is `#VALUE!`; a
  value where a range belongs is error 1004 when the formula is written,
  as in Excel. A SUBTOTAL recalculates as rows are hidden, filtered or
  shown. 31 layouts in live Excel pin it, including references that
  OFFSET, INDIRECT, IF, INDEX, a name or a reference operator give.
- SUMSQ and DEVSQ, in cells and through `WorksheetFunction`, each
  Excel's double to the bit over 87 sets of numbers. DEVSQ of no numbers
  is `#NUM!`.
- `Application.Evaluate`, `[...]` and `Worksheet.Evaluate` work out an
  expression as Excel's Evaluate does. What comes to cells is a Range: a
  reference, a name for cells, INDEX, OFFSET, INDIRECT, CHOOSE, IF, IFS,
  SWITCH or XLOOKUP landing on cells, an intersection or a union.
  Anything else is worked out as an array formula, blocks whole:
  `Evaluate("A1:A3*2")` is three numbers, in an array counted from 1
  that is one-dimensional when it is one row. An error comes back as an
  error value. An expression Excel cannot read, an empty one, or one
  past 255 characters is Error 2015; ROW() and COLUMN() with no
  argument are an array of one. 91 cases in live Excel pin it. Before,
  text that was not a reference, a name or an array constant reported
  itself unsupported.
- Formulas work with references as Excel does. The range operator
  joins any two references, so `SUM(A1:INDEX(A:A,5))`,
  `SUM(INDEX(A1:A10,2):A5)` and `SUM(Block:A5)` add what they span.
  An intersection, `SUM(A1:A5 A2:B3)`, reads the cells two blocks share
  and is `#NULL!` when they share none. A union, `SUM((A1:A2,A4))`, is
  read whole by SUM, COUNT, COUNTA, AVERAGE, MAX, MIN, LARGE and
  SUBTOTAL, picked from by INDEX's fourth argument and counted by AREAS.
  COUNTIF refuses it with `#VALUE!`, and a cell holding one is
  `#VALUE!`. INDEX, OFFSET, INDIRECT, CHOOSE, IF, IFS, SWITCH and
  XLOOKUP give cells wherever a reference is wanted, for ROWS, ROW,
  OFFSET's start, a COUNTIF or a SUBTOTAL.
  INDIRECT reads a defined name and an R1C1 reference counted from the
  top left. 45 formulas in live Excel pin it. Before, each reported
  itself unsupported. Another function given a union, and a relative
  R1C1 reference in INDIRECT, still do.
- Sheet protection: `Worksheet.Protect`, `Unprotect`,
  `ProtectContents`, `ProtectDrawingObjects`, `ProtectScenarios`,
  `ProtectionMode`, `EnableSelection` and the Protection object. A
  protected sheet refuses a macro what Excel refuses it, with Excel's
  own error text: writes to locked cells (the unlocked ones still take
  the value), clears, formats unless allowed, inserting and deleting
  rows and columns unless allowed, merging, filling, copying onto locked
  cells, AutoFilter, RemoveDuplicates and sorting. UserInterfaceOnly
  frees macros. A second Protect, the password rules and the Allow
  options follow Excel's. 168 cases in live Excel pin it. A save writes
  the protection as Excel does, the password as a salted SHA-512 hash
  Excel opens, and a protected sheet in a file is protected when read,
  its password checked against Excel's hash or the legacy 16-bit one.
- `Scripting.Dictionary`, made by `CreateObject` or `New` without
  starting the Scripting Runtime. Keys match as the real dictionary
  matches them: 1, 1# and 1& are one key, a Date is its serial number,
  True is -1, Empty and "" are one key, an object is itself, and text
  goes case and all unless CompareMode is vbTextCompare. Reading a
  missing key adds it. Keys and Items are zero-based arrays in insertion
  order, For Each walks the keys, and the errors are the real one's. 47
  probes in live Excel pin it.
- `Range.CountLarge`, `Next`, `Previous`, `Calculate`, `AddressLocal`,
  `FormulaLocal`, `FormulaR1C1Local` and `NumberFormatLocal`, and
  `Worksheet.Next` and `Previous`. The Local spellings are the English
  ones, as an English Excel has them; `Calculate` answers Null.
- Array formulas: `Range.FormulaArray`, `HasArray` and `CurrentArray`.
  An array formula is worked out once as an array over its block, and
  each cell shows its item; every cell reports the formula, and what
  reads one follows the array. FormulaArray takes A1 or R1C1 and at most
  255 characters, rewrites a whole array from one of its cells, and is
  error 1004 over part of one. Value and Formula over part of an array
  change nothing and over all of it replace it; clearing part of one, or
  inserting or deleting rows or columns through it, is error 1004, as in
  Excel. A file keeps the formula in the block's first cell, read and
  saved as Excel does; before, the model read such a formula as an
  ordinary one, and a save of the sheet lost the array. 78 cases and 4
  workbooks in live Excel pin it. Copying part of an array, and filling,
  sorting, replacing or removing duplicates in one, report themselves.
- `CallByName`, a member named at run time: VbGet reads a property and
  passes its arguments, VbLet sets one, VbMethod calls a method, and
  VbMethod on a host object's property, VbSet with a value, or a call
  type with none of the four is error 438, as 24 probes in live Excel
  show. `Collection.Item` is a method, as VBA's own Collection has it.
- Windows: `Application.ActiveWindow`, `Application.Windows`,
  `Workbook.Windows`, `Application.Goto` and the Window object -- zoom,
  view, gridlines, headings, zeros and formulas, scrolling, frozen panes
  and the macro recorder's Freeze Top Row, SplitRow and SplitColumn,
  caption, tabs, tab ratio and scroll bars. Each sheet keeps its own view
  and selection, a copy of the sheet takes them along, and a save writes
  them as Excel does. Windows lists windows front to back as Excel does.
  113 workbooks in live Excel pin it. A split that is not frozen, which
  Excel keeps in twips of what the window shows, is answered for and not
  saved; freezing with the active cell out of view reports itself.
- Data validation: `Range.Validation` and the Validation object --
  `Add`, `Modify`, `Delete`, every setting and message, and `Value`,
  which checks the cell's value against its rule as Excel does. Rules
  move with inserted and deleted rows and columns, split round a Delete
  or a Clear, travel with a copy, and a rule like one the sheet has joins
  it; a save writes the dataValidations element as Excel writes it. 39
  workbooks in live Excel pin it. A list read from another sheet, which
  Excel keeps in the worksheet's x14 extension, is read and not made.
- Excel tables: `Worksheet.ListObjects`, `ListObjects.Add`, and the
  ListObject, ListColumns, ListColumn, ListRows and ListRow objects,
  with `Range.ListObject`. A file's tables are read, and Add makes one
  over a range as Excel does: headers from the first row turned to text,
  generated names for blank and repeated ones, a header row put in when
  the range has none, the first free TableN, TableStyleMedium2, and
  error 1004 over another table. A table's name, style and style options
  can be set. A save writes each new table in the parts Excel writes for
  it; a table read from a file keeps its part until it changes. 8
  workbooks and 15 questions in live Excel pin it. Editing a table's rows
  and columns comes later.
- A table's totals row: `ListObject.ShowTotals` and
  `ListColumn.TotalsCalculation`. Showing the row puts cells in under the
  table, those below moving down; the first time, the first column says
  Total and the last adds up, or counts where its values are not all
  numbers. Each function writes its SUBTOTAL, `SUBTOTAL(109,[Price])`;
  hiding the row deletes its cells and keeps what each column totals for
  the next time. A label, a SUBTOTAL or another formula written into the
  row is what the column totals from then on, and a number written there
  becomes its text. The table part keeps each label, function and custom
  formula as Excel does. 20 workbooks in live Excel pin it.
- Editing a table's rows and columns: `ListRows.Add`, `ListRow.Delete`,
  `ListColumns.Add`, `ListColumn.Delete`, `ListObject.Resize` and
  `ListObject.Delete`, each moving the cells around the table and the
  references to them as Excel does, a structured reference to a column
  that goes becoming #REF!. A formula written into an empty table column
  fills it and makes it calculated, and a row the table gains takes the
  formula. A value or formula a macro writes in the row under a table or
  the column right of it takes the table over it, and a value written
  under it stretches `SUM(C2:C4)` over the new row as Excel does. 28
  workbooks and 21 writes in live Excel pin it. `ListObject.Unlist`, and a
  copy or fill next to a table, report themselves.
- Structured references. A formula can name part of a table,
  `Table1[Qty]`, `Table1[#All]`, `Table1[[#Headers],[Qty]:[Price]]`,
  `Table1[@Qty]`, and the model spells it back, saves it, reads it from a
  file and works it out as Excel does: this row as `@` on screen and
  `[#This Row]` in the file, the whole table as its bare name and
  `Table1[]`, column names escaped and bracketed as Excel escapes them, a
  column read as this row's cell where one value is wanted, and the
  table's name left out of `[Qty]` and `[@Qty]` inside the table.
  `Range("Table1[Qty]")` and Evaluate read them. Renaming a table,
  setting `ListColumn.Name` or writing over a header renames every
  formula with it, and a header names its column with the text it shows.
  AutoFill across moves a reference to one column along the table. 80
  formulas, a sweep of 50 column names, copies, fills and renames in live
  Excel pin it.
- Events, and a sheet's own code. A sheet's module and ThisWorkbook
  are their objects' code: a code name reaches its sheet or workbook, as
  `Sheet1.Range("A1")`; inside the module `Me` is the object and its
  members are the module's by name, so Range there is that sheet's; and
  outside code reaches the module's Public members through the object,
  `Sheet2.MyMacro` or `Application.Run "Sheet2.MyMacro"`.
  `add_module(..., kind="document")` binds a module, and an xlsm's own
  modules are bound as it is read. `Worksheet.CodeName` and
  `Workbook.CodeName` come from the file. The modules hear Change,
  Calculate, SelectionChange, Activate, Deactivate and NewSheet, sheet
  first and workbook second, with Excel's Target and order, suppressed by
  EnableEvents, and nested when a handler edits a sheet. 35 actions and 7
  reach cases in live Excel pin it. Before, a `Worksheet_Change` handler
  never ran and `Sheet1` was not defined.
- Workbook protection: `Workbook.Protect`, `Unprotect`,
  `ProtectStructure` and `ProtectWindows`. A protected structure refuses
  adding, deleting, renaming, moving, copying and hiding sheets with
  Excel's own errors. A second Protect needs the password, and one with
  no Structure turns the protection off, as in Excel; Windows does
  nothing, as in Excel today. It is saved and read as Excel saves and
  reads it, the password as a salted SHA-512 hash Excel opens. 41 cases
  and 5 files in live Excel pin it.
- The constants of VBA's own type library that the reference dumps left
  out: the colours (`vbRed`), the key codes (`vbKeyReturn`), the system
  colours (`vbButtonFace`), `vbModeless` and the QueryClose modes.
  `scripts/dump_vba_typelib_constants.py` reads them from VBE7.DLL.
- Defined names are written in Excel's order: by name as the Name Manager
  shows it, ignoring case, a sheet's own before the workbook's.
- `Range.RemoveDuplicates`. A row goes when the columns asked for hold
  what an earlier row's hold, and the rows left close up from the top of
  the range, formats and all, formulas shifting as a copy's would, while
  nothing outside the range moves. Cells match when both are blank, the
  same error, text the same but for case, or numbers of equal value that
  show the same: 1 and 1.00 differ, and so do 0.3 and 0.1 + 0.2. TRUE is
  a number 1 that shows TRUE; text never matches a number. Header is xlNo
  when left out and xlGuess guesses as Sort does; a single cell works on
  its current region. Columns left out or 0 does nothing, one outside the
  range is error 1004, an array of them error 5 if any is outside. 98
  layouts in live Excel pin it; text beyond ASCII, which Windows compares
  with æ matching ae, reports itself unsupported.
- `Range.AutoFill`, with every fill type but Flash Fill. The source
  repeats down, up or across the destination one column or row at a
  time; runs of numbers, dates, times, text with a number in it, and day
  and month names carry on as series, and the rest repeats, formulas
  moved as a copy moves them. Numbers step by 1 alone and by their
  difference in pairs, each value the first plus the step times how far
  along, that product rounded to 15 digits and then the sum, as Excel
  does (1/3, 2/3 goes on 1, 1.33333333333333, 1.66666666666666). A number
  that is the whole source repeats, and so does a lone cell in a row
  filled down beside others of its kind. Dates step by days, by months
  under `mmm-yy` or where they share a day of the month or all end one,
  and by weekdays, months or years when asked; times by an hour. Text
  keeps what surrounds its number and the zeros in front of it, 1st goes
  to 2nd, Q4 and 4th Qtr wrap to the first quarter, and Mon and JANUARY
  keep their case. Numbers run together while their formats agree.
  `xlFillCopy`, `xlFillFormats` and `xlFillValues` repeat everything, the
  formats alone and the values alone. 208 layouts and 520 columns in live
  Excel pin it. What the model cannot match reports itself: a trend
  through three or more numbers that do not step evenly, which Excel
  takes from its LINEST arithmetic, a growth trend, dates whose times
  differ, and a destination that runs on two ways at once.
- Worksheet functions from VBA go through the formula engine:
  `WorksheetFunction.X` raises error 1004 on an error and `Application.X`,
  the late-bound form, hands it back as an error value, for every one of
  the 72 WorksheetFunction members the engine has, `Match`, `Index`,
  `Transpose`, `CountIf`, `SumIfs`, `XLookup` and `Round` among them. A
  range argument is its cells, a VBA array a row or rows, Empty an
  argument left out; an array answer is a VBA array from 1, one row
  one-dimensional unless INDEX, CHOOSE or XLOOKUP answered with part of a
  range. The hand-written `Sum`, `Average`, `Max`, `Min`, `Count`,
  `CountA`, `Trim`, `Text`, `Proper` and `VLookup` are gone; `VLookup`'s
  approximate match now works. 150 calls made both ways in live Excel pin
  it; a range or array given where one value is wanted is not yet run
  item by item.
- `Range.Sort`, and `Worksheet.Sort` with its `SortFields`, `SetRange`,
  `Header`, `MatchCase`, `Orientation` and `Apply` as a recorded macro
  uses them. Numbers sort before text, text before FALSE and TRUE, those
  before errors, and blanks go last either way; the sort is stable. Text
  sorts as Excel sorts it -- marks before digits before letters, case
  ignored unless asked for, hyphens and apostrophes passed over but for
  ties -- for ASCII, checked on 400 random strings; text beyond ASCII
  reports itself unsupported. Rows move whole, formulas shifting with
  them. Up to three keys, orders, orientation and text sorted as numbers
  work; a single cell sorts its current region; an omitted Header is
  xlNo, and xlGuess takes the first row for a header as Excel does. 58
  layouts and four orderings of 583 values in live Excel pin it.
- The clipboard, as recorded macros use it. `Range.Copy` and the new
  `Range.Cut` with no destination put the range on it, and
  `CutCopyMode` reads 1 or 2. The new `Worksheet.Paste` pastes it at a
  destination or the selection, and the new `Range.PasteSpecial` pastes
  everything, values, formulas, formats, number formats with values or
  formulas, all but borders, or column widths, with Excel's operations,
  skipped blanks and transposing. What is pasted is what the cells hold
  then, a copy lasts through any number of pastes, and a paste fills the
  destination with copies when it is a whole number of the source's
  size. `Cut` moves cells, and every reference in the workbook wholly
  inside them follows, names included, while one wholly inside the cells
  they land on becomes `#REF!`. 49 layouts in live Excel pin it.
- `Range.Replace`, as Excel's does it. It edits the text the formula bar
  shows -- 1234 in `#,##0` is `1234`, a date `1/2/2020`, 50% `50%` -- and
  types what is left again, so the number 123 with `2` replaced by `5` is
  the number 153, and even a Text cell takes a number or a formula. The
  wildcards are Find's, a `*` as short as will match except at the end.
  A formula it cannot enter stops it there. It works on several areas
  and whole sheets, always answers True, and leaves its settings for the
  next `Find`. 80 layouts in live Excel and the formula-bar text of 20
  numbers in 26 formats pin it.
- `Range.SpecialCells` for constants and formulas of each kind, blanks,
  the last cell and visible cells, with Excel's own quirks: a one-cell
  range searches from A1 to the last cell, a merged area comes whole, and
  what is found is split into the very areas Excel returns -- taken from
  the last cell back, each joining the one-column area below it or the
  one-row area to its right. 18 layouts of live Excel, eight of them
  random grids, pin it. Comments, validation and conditional formats
  report themselves unsupported.
- `Range.FillDown`, `FillUp`, `FillRight` and `FillLeft`: each area's
  first row or column, in the direction of the fill, copied over the rest
  as `Copy` copies it, blanks included; a one-row or one-column area fills
  from its neighbour, and one at the edge of the sheet is error 1004
  before anything changes. 16 layouts of live Excel pin them.
- `Range.CurrentRegion`, as Excel answers it: from the first area's
  top-left cell, grown while the ring around it holds a value or a
  formula, corners included, taking in merged areas whole. 34 layouts of
  live Excel pin it, eight of them random grids read from every other
  cell.
- Rows grow with their borders as Excel's do. A medium line along a row's
  bottom edge draws it a pixel taller; a thick or double line draws the
  row below a pixel taller too; and a save writes the `thickBot` and
  `thickTop` flags, the height and the deeper descent Excel writes, on
  the rows and on the sheet's standard row when column formats carry the
  lines. Where lines of both weights meet along one edge, the first cell
  along the row with one decides, as in Excel. 39 cases of live Excel pin
  the answers, one of them Excel's answer for a border along a hidden
  row, which the model reports it cannot tell.
- Row, column and sheet-wide formats as Excel keeps them. Formatting a
  range that spans every column formats its rows, one that spans every
  row its columns, and the whole sheet every column. Each row and column
  keeps a format of its own (a row's `s` with `customFormat`, a column's
  `style`), read from and saved to the file, and a position with no cell
  shows its row's format, else its column's. A formatted row and a
  formatted column crossing get a cell of their own, a new cell starts in
  its row's or column's format, and an empty cell goes once its format is
  the one it would inherit. `ClearFormats` and `Clear` put positions back
  to the default, inserted rows and columns take the formats beside them,
  and copies carry what each position showed. Borders on whole rows,
  columns and the sheet follow Excel's own rules for each edge. Reads over
  whole rows and columns answer Null exactly where Excel's do, without
  visiting every position. A row's height counts the font at every
  position, from a second table measured for each font on a row to itself,
  so a row or column format makes rows shorter or taller and moves
  `StandardHeight`. 149 probes of live Excel pin the answers; Excel's own
  saved workbook reads back; 16 files written the way another program
  might write them read and save as Excel's do; the model's saved rows,
  columns and cells are Excel's XML; and a live gate opens them in Excel.
- Rows grow with their fonts as Excel's do. A row that keeps no height of
  its own is as tall as its tallest font, read from a table measured in
  live Excel: every pixel size up to 409.5pt of Aptos, Calibri, Arial,
  Cambria, Consolas, Courier New, Georgia, Segoe UI, Tahoma, Times New
  Roman and Verdana, in all four styles, 25,662 rows in all. `AutoFit`
  sizes a row to it, merged cells over several rows and empty turned
  cells leave rows alone, and a save writes the height and descent Excel
  does. A row whose height rests on something not measured -- another
  font, fonts from two tables, super- or subscript, wrapped or turned
  text -- keeps the height its file recorded until it changes, and then
  reports itself unsupported rather than guess.
- Row heights, column widths and hidden rows and columns as Excel keeps
  them on a 96-DPI display. `Range` gains `Hidden`, `UseStandardHeight`,
  `UseStandardWidth`, `Height`, `Width`, `Left` and `Top`, and a sheet
  `StandardHeight` and `StandardWidth`; `RowHeight`, `ColumnWidth` and row
  `AutoFit` follow Excel, and all of it is read from and saved to the
  file. A height is rounded to twips and then to quarter pixels, a width
  to whole pixels, a zero size hides and keeps what it hid, and every row
  or column sized or hidden at once becomes the sheet's default. A read
  across several rows or columns answers Null exactly where Excel's does,
  which depends on where the sheet's cells are (459 live reads fit). 152
  probes of live Excel pin the answers, an Excel-sized workbook and one
  another program sized read back value for value, the model's saved
  rows, columns and sheet defaults are Excel's XML, and a live gate opens
  them in Excel. Heights and tints are spelled with Excel's rule for a
  measurement: 15 digits when the double is within 5/16 of an epsilon of
  them, relative to its size, and 17 otherwise.
- Cell formats as Excel keeps them. A workbook's stylesheet is read into
  one format per cell: font, fill, borders, alignment, protection and
  number format. `Font` gains `Underline`, `Strikethrough`,
  `Superscript`, `Subscript`, `FontStyle`, `ColorIndex`, `ThemeColor`
  and `TintAndShade`; `Interior` gains `Pattern`, the pattern colours,
  `ThemeColor` and `TintAndShade`; `Range` gains `Borders`,
  `BorderAround`, `ClearFormats`, `VerticalAlignment`, `WrapText`,
  `IndentLevel`, `Orientation`, `ShrinkToFit`, `AddIndent`,
  `ReadingOrder`, `Locked` and `FormulaHidden`. 161 probes of live Excel
  pin the answers, including Null for a mixed range, the palette and
  tint arithmetic, and where a border between two cells is stored. An
  Excel-formatted workbook reads back value for value, a live gate opens
  the model's own formatted workbook in Excel, and for the same edits the
  saved stylesheet is Excel's byte for byte except for the position of a
  red font.
- `docs/excel_checklist.csv`: a status for every member of every Excel
  class the model implements, with its signature from the type library,
  the fixtures that are its evidence and the gaps a registration visibly
  has. `scripts/build_excel_checklist.py` regenerates it, and
  `tests/test_excel_checklist.py` keeps it in step with the registry.
- VBA `Worksheet.Move Before:=...` / `After:=...`, with same-book reordering,
  cross-book transfer, and new-workbook moves; Python `SheetView.move` delegates
  to the same model implementation. Cross-book moves invalidate held VBA objects
  with error 424, matching native Excel; reacquire them from the destination.
  Same-book reordering preserves objects. External-link and unsupported content transfers
  fail before mutation; reserved local-name indices follow same-book reordering.
- Cross-workbook range copies import referenced global/local names and aliases.
  Python callers can choose `name_conflict="reuse"`, `"rename"`, or `"error"`;
  the default matches unattended Excel. Renaming preserves existing names and
  rewrites incoming formulas/dependencies, with eight native behavior probes
  and native save/reopen checks for global and local aliases.
- Multi-workbook Python APIs (`open_workbook`, `workbooks`, `activate_workbook`,
  explicit workbook selection for `sheet`/`save`) and `SheetView.copy`/
  `copy_range`. VBA `Worksheet.Copy` supports Before/After in the same or another
  workbook and creating a new workbook, including unique tab names, cells,
  merges and relevant scoped/global names. Seven native copy comparisons and
  two native persistence gates cover modeled copies; external-link copies and
  advanced worksheet content remain explicitly unsupported.
  Workbook numbering no longer reuses names after Close; named Filename and
  SaveChanges arguments work. Newly added/copied tabs are registered when saving
  existing packages, preserving tab order and part identities across repeated saves.
- Named-range CRUD across workbook and worksheet scopes: named-argument
  `Names.Add`, scoped lookup/enumeration, replacement, rename, RefersTo/Value
  retargeting, absolute R1C1 forms, Visible, Comment, Range.Name and Delete.
  Renames update dependent formulas, aliases and control bindings; changes
  invalidate calculation dependencies. Nineteen native cases and three native
  persistence gates cover creation, update and deletion. Shared Python APIs
  provide detached `NamedRange` snapshots and add/update/remove operations on
  `ExcelApplication` and `SheetView`.
  Local names now retain qualified identities after reopen; Power Query's
  ExternalData-name handling follows that identity. Formulas without cached
  results calculate on first read after reopening.
- Whole-row/column insertion and deletion now update A1 formulas throughout
  the workbook and defined names, including absolute references, expanding or
  shrinking ranges, and deleted targets becoming `#REF!`. Twelve native cases
  and four native save/reopen checks cover both axes. Qualified errors such as
  `Sheet1!#REF!` now calculate correctly. Inserts that would discard occupied
  edge cells fail before mutation. Merges/shapes, partial-cell reference
  rewriting and full formatting/metadata movement remain incomplete.
  Fresh workbooks now register additional worksheet parts in the workbook
  sheet list and relationships, preserving cross-sheet formulas on reopen.
- Corrected explicit-destination `Range.Copy`: snapshots protect overlapping
  copies, trailing blank source cells clear destination cells, exact destination
  multiples repeat the source block, and copied formulas shift per block.
  Seventeen native cases and five native save/reopen checks cover geometry,
  cross-sheet copying, formulas and formatting. Merged-cell/multi-area copies
  and destinations larger than 1,048,576 cells remain explicitly unsupported.
  Bold font state now loads and saves, including copied/blank cells and
  bold-only edits; other font and formatting persistence remains incomplete.
- A1 `Range.Formula` array assignment shares the R1C1 array writer, with
  native A1 reference shifting during singleton-axis expansion. Fifteen native
  comparison cases and four native save/reopen checks cover formulas, mixed
  values, errors, bounds, merged cells and multiple areas. Both formula getters
  now share array reads that retain requested trailing blank cells.
- Headless Excel `Range.FormulaR1C1` scalar/block assignment and scalar/array
  reads, including relative, absolute and mixed references, whole rows/columns,
  quoted text and worksheet-edge wrapping. Twelve native comparison cases,
  headless persistence and a native Excel save/reopen gate cover conversion.
  Array assignment additionally passes 15 native cases and four native
  save/reopen checks, covering custom lower bounds, singleton-axis expansion,
  errors, empty arrays, merged cells and multiple areas. Uncovered matrix cells
  receive `#N/A`, and higher-dimensional arrays raise error 13.
  Workbook loading now preserves numeric cells as Doubles instead of narrowing
  whole numbers to VBA Integer/Long; formula reads use Excel's boolean/error text.
- Headless Excel `Merge`, `UnMerge`, `MergeCells` and `MergeArea`, including
  ordinary overlap expansion, across-row merges, first-nonempty value/formula
  retention and merged-region persistence. Twenty-one native cases and four
  live save/reopen checks cover values, flags, partial-clear errors and unmerge.
  Range enumeration now includes explicitly requested blank cells outside the
  used range. Multi-area merges and across merges over existing multi-row
  merged regions remain unsupported.
- Headless Excel `Range.Find`, `FindNext` and `FindPrevious` for single-area
  ranges: values/formulas, whole/partial matching, wildcards and escapes,
  case sensitivity, row/column traversal, wraparound, blank cells and saved
  application search settings. Forty-eight native cases pin behavior;
  save/reopen and sparse whole-sheet checks cover persistence and scale.
  Multi-area, comments, format searches and MatchByte=True remain explicit
  unsupported operations.
- Absolute R1C1 addresses in `INDIRECT`-backed Excel control names, with
  eight native behavior cases and a save/reopen gate for cell-driven
  source and linked-cell addresses. Relative R1C1 remains unsupported.
- Shared `references()`, `add_reference()` and `remove_reference()` APIs
  across Excel, Word, PowerPoint and Access, with Office presets and custom
  registered libraries. GUID-based adds are idempotent; removing an absent
  reference returns false. Access's strict `drop_reference()` remains.
  Whole control/project reference groups are preserved or removed intact;
  MSForms removal is refused while UserForms exist. Reference-only edits
  honor save protection/signature rules and invalidate compiled caches.
  Native Excel, Word and PowerPoint gates check references before VBA
  injection, compile early-bound declarations, and verify removal.
- Conditional Forms control names using `CHOOSE` and `IF`, with lazy
  branch selection, cell-driven targets and empty sources for selected
  scalar values/errors. Sixteen native Excel cases and two live
  save/reopen gates cover lists and linked-cell writes.
- Single-area `INDEX` names for Forms control sources and links, including
  whole rows/columns, aliases, nested reference formulas and cell-driven
  targets. Fifteen native Excel cases verify indices, selected items and
  linked destinations; a live gate verifies dynamic bindings after saving.
- Forms control names backed by `OFFSET` and A1 `INDIRECT`, including
  nested offsets, aliases, cell-driven addresses and calculated list
  sizes. Twelve native Excel cases cover dynamic selection and invalid
  ranges; save/reopen gates verify dynamic sources and linked targets.
- Named Forms control links and list sources, including workbook/local
  names, aliases, missing names and changes to their targets. Twenty-four
  native Excel cases cover resolution and selection behavior; live
  save/reopen gates verify workbook and worksheet-local bindings.
  Newly created local names serialize with Excel's worksheet scope.
- Late overlapping-box additions, including nested-group link retention
  and identical-box link transfer to an unlinked sheet group. Eighty new
  native Excel cases cover link combinations and selection states.
- Radio ownership in existing partial, nested and identical overlapping
  boxes, plus group-box deletion and link transfer. Thirty-two native
  behavior cases cover creation order, deletion and late additions.
- Non-overlapping, interleaved radio groups and multi-box regrouping,
  replaying 44 native Excel cases. Membership is retained during movement
  and reconstructed on open, verified with an Excel-authored moved-radio
  fixture and two live save/reopen gates.
- Single-box radio regrouping: adding a box around a contiguous prefix,
  suffix or whole group, and deleting it to merge the groups. Thirty-six
  native Excel cases cover selection and linked-cell effects, with live
  save/reopen checks for both operations.
- Radio-group values, shared linked cells, persisted group boundaries and
  deletion behavior, supported by 70 native Excel behavior cases and an
  Excel save/reopen gate.
- `SheetView.add_form_control` creates all nine supported Forms control
  types with validation before mutation and detached return snapshots.
- Multi/extended list source and linked-cell rebinding, plus selection
  preservation when switching between those modes, replaying 28 native
  Excel measurements and verified by an Excel persistence gate.
- Spinner/scroll-bar values, linked cells, bounds and increments, replaying
  52 native Excel cases. Scroll bars support `LargeChange`; spinners
  reproduce Excel's error 438 for that member.
- Issue #25 regressions using Excel-authored control fixtures: radio state
  reading, Radio/Spin/Scroll/GBox type preservation, and numeric defaults
  and bounds. Unchecked controls omit the `checked` attribute as Excel does.
- A public Python worksheet shape API through `ExcelApplication.sheet()`:
  `shapes`, `shape`, `add_shape`, `add_textbox`, `add_button`,
  `update_shape` and `remove_shape`. Reads return detached snapshots
  with control details; edits share the VBA model's state. Names and
  geometry are validated before mutation.
- `SheetView.update_control` edits or clears saved linked-cell and list-range
  bindings, updating control properties and VML together. It validates A1
  references before mutation and leaves existing cell values untouched.
- Checkbox values through `SheetView.set_control_value` and VBA's
  `Shape.ControlFormat.Value`, plus the checkbox `ControlFormat.LinkedCell`
  property. Direct cell writes and formula calculations synchronize linked
  checkboxes across sheets. Thirty live Excel probes cover defaults,
  invalid values, unchanged assignments, text/errors, formulas and rebinding.
- Single-selection dropdown/list-box values through the same Python/VBA
  API, with A1 range sources, numeric linked-cell synchronization, and
  VBA `ListFillRange` and `ListCount`. Fifty live Excel probes cover
  bounds, errors, coercions, formulas and source removal/restoration.
- Inline list read/add/insert/update/remove/clear APIs in Python and VBA,
  plus multi/extended list-box selection, `ControlFormat.MultiSelect`,
  and `DrawingObject.Selected(index)`. Fifty-four live Excel probes cover
  item edits, selection shifts, mode transitions, native constants and
  source conversion. Inline items and selected indexes persist through
  the properties and VML parts, with live save/reopen gates for all modes.
- Whole-list VBA reads and assignments, plus Python `set_control_items`.
  Ninety-two live cases cover detached one-based arrays, Null/empty lists,
  one- and two-dimensional assignments, array bounds, partial failures,
  and range-backed edits. Item additions/replacements disconnect an A1
  source and start an inline list without changing source cells; individual
  removal from a range source raises Excel error 1004.

### Changed

- A cell holds every number as a Double, a date included; its format
  decides what `Value` reads. A query's refreshed rows keep their text as
  text rather than typing it.
- The in-memory Excel is Excel on a 96-DPI display (Windows at 100%): a
  standard row is 15pt where the model used 14.5pt, the height Excel gives
  it on a 144-DPI display. A shape's cell anchor is read and written
  against the same rows and columns, so a shape nobody moved stays where
  it was when its drawing is saved again.
- A sheet the model adds is written as Excel writes a new sheet, with its
  view, row defaults and page margins, and a rewritten sheet's row
  defaults are this display's.
- Setting any part of a format marks that part as applied, as Excel does,
  even when the value does not change. An xf's apply flags now say what a
  macro set rather than what differs from the cell style, so
  `Font.Bold = False` on an untouched cell leaves a formatted empty cell
  in the used range and in the file, as in Excel. A file's apply flags
  are not read, since Excel takes a format that differs only in them for
  the same one.
- A loaded sheet that Excel tidies on opening is written again on save as
  Excel writes it, instead of keeping its bytes. That covers an empty cell
  in the format its row or column gives it, a row format every column
  shows, a column with no width, and a stale dimension.

### Fixed

- A formula whose answer is empty text, which Excel saves as `<v/>`,
  reads back as "" rather than as an empty cell.
- An `e` in a number format with no sign after it shows the year, as
  Excel's era year does in this locale: `TEXT(0,"0;-0;zero")` is
  `z1900ro`, and `Range.Text` shows a cell the same way.
- A workbook holding a formula the parser cannot read, such as a LAMBDA
  called where it is made, works out every other formula as before; that
  cell keeps the value its file gave it. The model stopped at the first
  such formula, so writing any cell of the workbook failed.
- `Range.Select` and `Range.Activate` on a sheet that is not the active
  one are error 1004, as in Excel; the model made the sheet active.
  `Activate` on a cell inside the selection moves the active cell and
  keeps the selection; the model selected the cell alone.
- A workbook opens on the sheet its file had active, with that sheet's
  selection, and a save moves the selected tab to the active sheet. The
  model opened every workbook on its first sheet, and a save left the
  tab where the file had it.
- Closing the active workbook activates the window behind it, as Excel
  does; the model activated the workbook opened last.
- A save works out every formula nothing has read since it was written,
  and saves it with its value, as Excel's cells are always up to date
  under automatic calculation. The model saved such a formula with no
  value. One the model cannot work out is still saved without one, and
  under manual calculation a cell keeps the value it last had.
- `Err` clears where VBA clears it: when a procedure of the macro starts,
  on every `On Error` statement, and on leaving a procedure from its
  error handler. It keeps what a procedure left in it after the procedure
  returns, and a built-in function or an object's member leaves it
  alone. The model never cleared it on a call or an `On Error`
  statement, so an error handled in one function was still in `Err` in
  the next one. 23 programs in live Excel pin it.
- A table part names a column holding a tab or a line break with Excel's
  `_x0009_` escapes; the model wrote the characters themselves, which a
  reader of the file takes for spaces.
- CHOOSE with an array of indexes chooses for each item, as Excel does
  in a cell: `CHOOSE({1,2},A1:A3,B1:B3)` is the two columns side by
  side, so `VLOOKUP(x,CHOOSE({1,2},B:B,A:A),2,FALSE)` looks to the left.
  The model chose once, by the first index. 7 formulas in live Excel pin
  it.
- `Me` in a class module is the instance the code runs for. The model
  reported it as an undefined variable.
- Arguments given to a property that takes none apply to what it
  answers, as in VBA: `Range("C1:C3").Formula(2, 1)` is the second
  formula of the array Formula answers. The model raised error 450.
- A workbook Excel saved with a formula filled down, copied or written
  to a block reads with every cell's formula. Excel stores such a block
  as one shared formula whose other cells only point at the first; the
  model read those cells as the values they last showed, so they never
  changed again. A save now writes shared formulas as Excel does: over
  the block a macro writes, fills or copies a formula to, kept through
  edits, inserts and deletes as Excel keeps them, and numbered as Excel
  numbers them. A formula's text and value keep a quote as it is, as
  Excel writes them. 11 workbooks saved by Excel pin it.
- A formula written to a General cell takes the number format Excel
  gives it, as it is written: `=A1+30` on a date is a date, so its Value
  reads back as a Date rather than a Double. A reference brings its
  top-left cell's format, `+` and `-` the first side's but nothing for
  two dates, `*`, `/` and `&` nothing, and SUM, MAX, MIN, INT, ROUND,
  ROUNDDOWN, ROUNDUP, TRUNC and MOD their first formatted argument's.
  TIME brings h:mm AM/PM. A range written at once takes its first cell's
  format. The model formatted only a formula that was itself DATE, TODAY
  or NOW, and only when it was first worked out. 111 cases in live Excel
  pin it.
- A range or an array where a formula wants one value is handled as
  Excel handles it in a cell. Cells given to a function that wants one
  value, beside an operator, or as the whole formula are cut to the one
  on the formula's own row or column, as Excel's `@` does: `=LEN(A1:A3)`
  in row 2 is `LEN(A2)` and in row 5 `#VALUE!`, and `SUM(A1:A3*2)` in
  row 2 is 4. An array given there is run through item by item, so
  `SUM(LEN({"a","bb"}))` is 3, except by INDEX, VLOOKUP, HLOOKUP,
  XLOOKUP, IFERROR and IFNA, which take its first item. IF takes a
  branch for each item of an array of conditions. SUMPRODUCT's
  arguments, INDEX's first and those of ROWS and COLUMNS are worked out
  as arrays even in a cell, except inside an IF, CHOOSE, IFERROR, IFNA,
  IFS or SWITCH, and a defined name's formula as an array formula. ROW
  and COLUMN give every row or column only in an array. The model read
  the first cell of a block in a cell and an operator's whole block
  everywhere, so a formula written before dynamic arrays could come out
  wrong without saying so. 232 formulas in live Excel pin it.
- A value where the -IF and -IFS functions, COUNTBLANK, OFFSET, ROW,
  COLUMN or AREAS read cells is error 1004 when the formula is written,
  as it already was for SUBTOTAL, and a function that answers with a
  value counts as one: `SUMIF(LEN(A1:A3),1)` is refused, as is
  `SUBTOTAL(9,LEN(A1))`, which the model took. A name there that comes
  to a value is `#VALUE!`.
- An argument that comes to an error reaches its function as a value:
  `COUNT(A1:A3*2)` in row 5 is 0 and `COUNTA` of it 1, as in Excel. The
  model answered the error.
- DAY, MONTH and YEAR read a serial number before March 1900 on Excel's
  calendar, which has a 29 February 1900: `DAY(1)` is 1 and `YEAR(1)`
  1900. The model read such a date a day early.
- A defined name whose formula comes to a value rather than cells, such
  as `=IF($A$1:$A$3>1,1,0)`, is worked out where it is used. The model
  answered `#NAME?`.
- `MATCH(3,3,0)` is `#N/A`: MATCH looks only in a range or an array.
- ISREF, which reported itself unsupported.
- `Rows.Count` is 1048576 and `Columns.Count` 16384: Count counts the
  rows of whole rows and the columns of whole columns, as Excel does,
  `EntireRow` and `EntireColumn` included. The model counted their cells,
  so `Cells(Rows.Count, 1).End(xlUp)` failed. `Cells.Count` is error 6,
  as in Excel.
- `Range.Address` writes R1C1 style when asked, a relative part counted
  from RelativeTo, or from A1 without it. The model ignored
  ReferenceStyle and wrote A1 style.
- A constant is a Long, as VBA holds an enum's member however small:
  `TypeName(xlUp)` and `TypeName(vbOK)` are `Long`. Only the key codes
  are declared Integer. The model made every constant that fits an
  Integer one, so `vbYes * 10000` overflowed.
- Arithmetic that passes the largest double, as `=1E+300*1E+300` does,
  is `#NUM!`, as in Excel. The model answered infinity.
- A formula that ends on a `+` or `-` whose answer all but cancels is 0,
  as in Excel: `=0.5-0.4-0.1` is 0, not -2.8E-17. Excel sets the answer
  to 0 when its binary exponent is 50 or more below the left operand's,
  about seven steps of the last bit. Brackets around the whole formula,
  or a function around the difference, keep the bits. SUM and AVERAGE
  do the same to their last addition, and a defined name to its
  formula. 1,788 pairs of doubles measured in live Excel pin it.
- Comparisons round each number to fifteen significant digits, an exact
  tie going away from zero, as Excel does. The model found the digits
  with a logarithm, which misses by one beside a power of ten:
  999999999999999.375 equalled 1E+15.
- MATCH, VLOOKUP, HLOOKUP and XLOOKUP compare numbers bit for bit,
  exact and approximate, as Excel's lookups do. The model matched them
  to fifteen digits like `=`. COUNTIF, COUNTIFS, SUMIF and SWITCH still
  match to fifteen digits, as in Excel.
- A number turned to text in a formula rounds an exact tie at the
  sixteenth digit toward zero: 4503599627370495 reads
  `4503599627370490`. The model rounded half to even.
- A defined name that stands for a formula, such as
  `=Sheet1!$A$1-Sheet1!$B$1`, is worked out where the name is used. The
  model read that one as $B$1 on a sheet called `Sheet1!$A$1-Sheet1`,
  and answered `#NAME?` for other formulas. One with a relative
  reference reports itself unsupported.
- SUM, AVERAGE and AVERAGEIF add one number after another, each sum
  rounded to a double, as Excel adds. The model added exactly, and SUM
  came out a bit away from Excel's on 23 of 87 measured sets.
- STDEV, STDEVP, VAR and VARP are Excel's doubles. Excel takes the
  one-pass sum of squares, and the squares about the mean when that
  cancels; the model worked the exact variance. STDEV of 1, 6, 7 and 8
  is 3.1091263510296048, as in Excel. 340 of 348 answers over 87 sets
  match to the bit. The other 8 are on data under a thousandth and stay
  known misses.
- `CStr` and the text of a number round an exact tie at the last digit
  shown away from zero, as VBA does, for a Single as for a Double.
  `CStr(CSng(2 ^ -11))` is `4.882813E-04`; the model rounded it to even.
- `Range.Cut` moves a reference that loses a whole edge to the cells it
  moves, as Excel does. Across sheets the reference keeps the rest:
  `SUM(B2:B7)` reads `SUM(B5:B7)` once B2:B4 is cut away. On its own
  sheet the edge goes where the block goes when the block moves straight
  along the range: B7 cut to B9 leaves `SUM(B2:B9)`. The model left
  both as they were. 24 layouts in live Excel pin it.
- `Range.Delete` and `Range.Insert` with a `Shift` rewrite the references
  to the cells they move, as Excel does. A reference whose columns lie in
  the band that shifts moves and stretches as it would for whole rows,
  `#REF!` once everything it reads is deleted, and so do names, formulas
  on other sheets and the AutoFilter's range; one reaching outside the
  band stays, unless a delete takes a whole edge of it: `SUM(B2:C6)` reads
  `SUM(C2:C6)` once B2:B6 is deleted up. Inserted cells take the formats
  above them or to their left. With no `Shift`, a range taller than it is
  wide shifts across, as Excel decides; the model always shifted up or
  down, and never touched a reference. 47 layouts in live Excel pin it.
- A value assigned to a member that answers an object goes to that
  object's default member, as VBA sends it: `Range("A1") = 5`,
  `Cells(2, 2) = 7`, `ws.Range("A1:B2") = 0` and `r(2) = 3` all fill
  cells. The model raised error 438 for each.
- A formula cut to another sheet keeps pointing where it did. A reference
  into the cells that move with it stays as it was written, and one to a
  cell left behind names its old sheet: `=D1+B2` cut from CutAway1 reads
  `=CutAway1!D1+CutAway1!B2`. The model named the new sheet in the first
  and nothing in the second.
- `Range.Value` read a number under an elapsed-minutes format, `[m]` or
  `[mm]`, as a Date, taking the `m` for a month; Excel reads a Double.
- `NumberFormat` kept every escaped character as written. Excel drops the
  backslash wherever the bare character would mean nothing, so
  `\T\R\U\E` reads back `T\RU\E`: digits but 0, most marks and the
  letters that are no date, time or era code lose it, and a dot keeps it
  only in a section with digits, a comma only straight after one. Every
  printable character, measured four ways, pins it.
- `Range.Formula` read a stored number as VBA's `CStr` spells it; it now
  spells it as Excel does, written out up to 21 characters. `=A1&""`
  writes numbers out up to 20 characters as Excel does, where it switched
  to an exponent from 1E+11 and below 1E-4.
- `End` stopped at a cell with only a format; Excel walks over one as
  over an empty cell.
- A formula a macro writes is spelled as Excel spells it back, through
  `Formula`, `FormulaR1C1`, `Value`, an array or `Replace`: references in
  capitals and ranges from their top-left corner (`SUM(B2:A1)` is
  `SUM(A1:B2)`), sheets as they are named, Excel's functions, TRUE,
  FALSE and errors in capitals, defined names as they were defined, any
  other name as the workbook first saw it, and numbers written out again
  from their first fifteen digits (`=1.50` is `=1.5`, `=+1` is `=1`).
  Spaces stay, except before a comma, at the end and in an array
  constant. A formula Excel refuses to read raises error 1004 where the
  model kept it, and `=` alone is text. The model kept the text as
  written. 120 formulas in live Excel pin the rules.
- Formulas read ranges between two sheet-qualified references
  (`Data!A1:Data!B2`), intersections (`A1:B2 B1:C3`) and unions in
  brackets; working one out reports itself unsupported, where the text
  could not be read at all. A function's bracket has to follow its name,
  as in Excel, and an error typed in lower case is read.
- `FormulaR1C1` spelt a whole row or column in its own row or column
  twice, `C:C` for `A:A`; Excel writes `C`.
- Formula functions that differed from Excel's: `COUNT` of values given
  as arguments counts TRUE, FALSE, text that reads as a number and an
  argument left out; `COUNTA` counts one left out; `POWER(0,0)` is
  `#NUM!`; `DAYS` reads text dates; `CEILING.MATH` takes the magnitude of
  its significance and a mode; `INDEX` of one row with one index counts
  along it; `TRANSPOSE` of a single value is the value.
- `Find` among formulas read a constant as the cell shows it, so 1234 in
  `#,##0` was found by `1,234` and not by `1234`; it reads it as the
  formula bar shows it, as Excel does.
- The whole sheet's address was spelt `$A$1:$XFD$1048576`; Excel spells it
  `$1:$1048576`.
- `CLng`, `CInt`, `CByte`, `CSng`, `CDbl` and `CCur` of an Error value
  give its number, as VBA's do, and `Val` of one raises error 13.
- `CDbl("1E-25")` and any other number string with an upper-case
  exponent raised a Python error.
- `CStr` of a Date chose what to print from the serial rather than the
  moment rounded to the second: `CDate(0)` printed `12/30/1899` where VBA
  prints `12:00:00 AM`, and a time a hair short of midnight printed the
  time where VBA prints the next day's date.
- `CStr`, `Format`, `Hour`, `Minute` and `Second` of a Date that rounds
  past the last second of 9999 raised a Python error; VBA stays on that
  second.
- Built-in number formats 5 to 8, 12, 13, 37 to 44 and 48 were missing
  from the model's table, so a save gave them custom ids, and 47 was
  spelt `mmss.0` where Excel reads `mm:ss.0`.
- Saving a sheet took time proportional to the square of the cells in a
  row, since each cell was spliced into its row's XML one by one; a row
  of thousands of cells took minutes.
- A column a file gives no width reads as Excel reads it: hidden, and 0
  wide.
- `ClearContents` and `Clear` drop a cell they leave empty and in its
  inherited format, and no longer visit every position of a large range.
- `ColumnWidth` and `RowHeight` were kept only in memory: a file's widths
  and heights were never read, and a macro's were never saved.
- `UsedRange` and a saved sheet's dimension count rows that have a height
  or are hidden, as Excel's do.
- A new workbook's cells read Excel's defaults: `Font.Name` is the
  workbook's own font (Aptos Narrow in a current Excel, where Calibri was
  assumed), `HorizontalAlignment` is `xlGeneral` rather than `xlLeft`, and
  `Font.Color` answers a Double. `NumberFormat` over cells that differ
  answers Null.
- Excel classes whose members the type library keeps on an interface,
  Range on IRange and Font on IFont among them, are found there. A member
  Excel lacks raises error 438 and one it has reports itself unsupported;
  before, a typo on Range read as unimplemented, and Font was checked
  against Word's Font.
- Writing to or formatting more than 1,048,576 cells at once reports
  itself unsupported instead of raising error 1004, since Excel does it.
- Multi-selection controls now read Excel's lowercase `seltype` attribute
  and `multiSel` indexes; scalar `Value` reads raise error 1004 in VBA.

- Dropdown/list selections now read and persist the `sel` property and
  VML `Sel`, preserving the separate `val` field. Excel opens new and
  subsequently edited controls with the expected type and selection.

- Checkbox snapshots now report Excel's unchecked (`-4146`), checked (`1`)
  and mixed (`2`) values, replayed from live Excel measurements.
- Checkbox state persists in both control properties and VML; newly created
  checkboxes use the correct properties-part control type.
- Literal error cells read back as VBA error values, and assigned VBA error
  values save as error cells. Text TRUE/FALSE assignments become booleans,
  matching the measured linked-cell cases.

- Deleting a form control now removes its worksheet record, properties
  part, relationship and VML shape as well as its drawing. Shared parts
  and other VML shapes survive; the final control can be deleted and a
  new one added on a later save.
- Renaming, moving, resizing and changing a control's caption now reach
  both its worksheet/VML records and its drawing. Live Excel exposed
  that an existing control can have a second caption in DrawingML which
  takes precedence over the VML caption.
- Shape names containing a backslash are written literally, and a macro
  can be assigned to a shape whose drawing omitted the macro attribute.

### Internal

- pyOfficeEditor's formula engine, copied into `pyopenvba.formula._calc`
  from its commit 098e441: a parser, an evaluator and 493 of Excel's
  functions, with arithmetic as the x87 does it. Over the model's own
  workbook, through `_engine_book`, it works out all 10,958 formulas of
  pyOfficeEditor's corpus as Excel cached them, to the bit or within the
  units pyOfficeEditor allows; the engine the model calculates with
  matched 6,718. Cells keep the model's engine until its own live
  measurements agree with the new one: legacy formulas, lookups compared
  to the bit, the zero a last sum snaps to and variance bits among them.

## [6.0.0] - 2026-09-20

VBA runs. Not read, not analysed: executed, against an Office
application that lives in memory, with the file written back out
afterwards.

```python
from pyopenvba.apps.excel import ExcelApplication

app = ExcelApplication.open("report.xlsm")
app.run("BuildReport")
print(app.describe())
app.save("report_out.xlsm")
```

### Added

- **A VBA interpreter.** Every statement form, the runtime library, and
  three separate failures: `VBACompileError` for VBA that does not
  compile, `VBARuntimeError` carrying the number `Err` would hold, and
  `VBAUnsupportedError` for real VBA this does not implement, which
  `On Error` deliberately cannot trap. An unknown member is told apart
  by the type library rather than guessed: `Worksheet.PivotTables` is
  unsupported, `Worksheet.Pivottabel` is VBA's own error 438.

- **Excel, Word and PowerPoint object models**, each with real state
  behind it, each measured against its own application. Cells, ranges,
  names, queries, sheets; a document's text, paragraphs and ranges; a
  presentation's slides. `ExcelApplication`, `WordApplication` and
  `PowerPointApplication` open a file, run a macro and write it back.

- **A calculation engine.** A macro can write a formula and read the
  answer. Calculation is on demand, writing a cell spoils whatever
  reads it, and about a hundred worksheet functions are implemented;
  anything else Excel has says so by name rather than answering
  `#NAME?`.

- **A Power Query evaluator.** `WorkbookQuery.Refresh` works out the
  query's M and lands its rows on the sheet it loads to, with the table
  and its queryTable following. Queries can name each other and
  `Excel.CurrentWorkbook()` reads the workbook's own tables. Nothing
  reaches off the machine: the connectors report themselves.

- **Shapes on all three surfaces.** Adding, reading, moving, resizing,
  renaming, retyping and deleting, with the macro a click runs attached
  or removed -- `Shape.OnAction` in Excel, form-control buttons
  included, and `ActionSettings(ppMouseClick).Run` in PowerPoint. Word
  has no macro on a shape and says so. An Excel form control is written
  as its four parts, so a button a macro made opens as a button.

Every answer was measured in the application itself rather than
reasoned about: 245 VBA expressions, 223 formulas, 160 Excel object
model probes, 46 PowerPoint, 43 Word, 269 Power Query expressions, one
of every shape type in all three hosts, and a shape fixture per host,
all committed so the tests need no Office.
Seven live gates behind `RUN_LIVE_EXCEL`, `RUN_LIVE_WORD` and
`RUN_LIVE_POWERPOINT` ask what only Office can answer.

### Changed

- `WorkbookQuery.Refresh` evaluates the query instead of reporting
  itself unsupported.
- A date is written out rather than handed to `strftime`, whose way of
  asking for a number without its leading zero is Windows-only. The
  Access SQL writer had the same defect and the same fix.

## [5.2.4] - 2026-09-17

Two defects in the member list beside a form or report, the stream that
decides what `Me.` can reach. Both were reported with the measurements
and the fixtures that pin them.

### Fixed

- **An entry holds two names, and the reader took one** ([#22]). It stores
  the identifier VBA compiles against, a NUL, the name the designer shows,
  a NUL, leaving the second empty where the two are the same. A control is
  rarely named as VBA would name it: the wizard names one after its field,
  so `Order Date` is ordinary and code reaches it as `Me.Order_Date`.
  Reading to the first double NUL walked into the next entry's type id, so
  every later entry was misaligned. A form with an ActiveX control failed
  the edit outright; one without was rewritten with `Order Date` as the
  member, which no code can reach, and `Me.Order_Date` stopped compiling
  after any edit made here. Both names are read and written now, and an
  entry carried forward writes back the bytes it was read from. An
  identifier that is not one can only have come from the old writer, so
  the next edit rebuilds it from the design's name.

- **A name the code page cannot hold got a `???` member** ([#23]). Access
  writes no entry at all for one, and takes no ordinal for it. Every edit
  to a form with a control named in Cyrillic appended a `???`, and a
  second edit gave the form's class two members of that name. Such a name
  is left out now, on add and on rename alike, and no best fit is
  accepted: a name with a `?` in it names something else.

  Two members that are one identifier to VBA are refused, as Access
  refuses the second control name as already in use.

  The page is no longer hardcoded. PROJECTCODEPAGE is tried first, being
  the page of the machine that last saved the project, and the stream's
  own bytes overrule it. An earlier note claimed the stream is fixed at
  cp1252 because patching PROJECTCODEPAGE to 1251 changed nothing. That
  experiment was void: VBA went on reading the project as cp1252 and
  Access wrote 1252 back over the patch on its next save.

- **A refused edit left the design written and the member list stale.**
  `_rewrite_design` wrote the design blob before building the member
  list, so a refusal from the list left the database holding a design its
  class does not describe. Everything is computed before anything is
  written.

Every `TypeInfo` stream in the fixtures still rebuilds byte for byte, and
the live gate compiles a project whose code reaches each control by the
identifier Access gives it.

[#22]: https://github.com/WilliamSmithEdward/pyOpenVBA/issues/22
[#23]: https://github.com/WilliamSmithEdward/pyOpenVBA/issues/23

## [5.2.3] - 2026-09-08

Three defects found by hunting the same seam the last few reports came
from: the sheet writer reading its own output back, or Excel's, and
assuming everything else looks the same.

### Fixed

- **A sheet could not be picked by a name holding `&`, `<`, `>` or `"`.**
  Those are stored escaped, so a sheet named `A & B` sits in the file as
  `A &amp; B`. Attribute values were read as stored and compared against
  the name a caller passed, so the lookup failed and reported the escaped
  spelling back in the error. Values are decoded on the way in and
  escaped again on the way out.

- **An apostrophe in a sheet name was not doubled in the defined name.**
  A reference spells `It's` as `'It''s'`; a single apostrophe closes the
  quoting early, so the name pointed at something else. Names opening
  with a digit are quoted now as well, since unquoted they read as part
  of a cell reference. Excel accepts all of them and reports the
  references back exactly.

- **The sheet's declared extent was never widened on a foreign sheet.**
  The element was matched by its exact spelling, so a writer that puts a
  space before the closing slash, as openpyxl does, kept the extent it
  started with and the part said the sheet ended before the table began.

Nothing here changes what is written for a workbook Excel authored.

## [5.2.2] - 2026-09-08

### Fixed

- **Loading onto any sheet but the first made the workbook unopenable.**
  The hidden `ExternalData_N` name was written with a constant
  `localSheetId="0"`, and that attribute is the zero-based position of
  the sheet the name belongs to. Loading to the second sheet wrote a name
  claiming the first while its reference named the second, and Excel
  refused the file outright. It is looked up now, from the same list of
  sheets that settles which part a sheet is.

  A one-sheet workbook cannot show this, which is why every fixture here
  missed it. The live gate loads onto each of three sheets and has Excel
  open and refresh all three.

## [5.2.1] - 2026-09-08

### Fixed

- **Loading a query onto a sheet failed on a workbook another tool
  wrote.** Three assumptions here came from reading only Excel's output,
  and each broke on openpyxl's, which is as legal. Relationships were
  matched with the id written first, where openpyxl writes it last, so a
  sheet looked as though it had no part behind it and the load was
  refused. `<definedNames />` written closed got a second block appended
  beside it rather than being filled, leaving two in the workbook. And a
  `tablePart` was added using the `r:` prefix on worksheets that never
  declared it, because a worksheet with no table has no use for one,
  which left the part not well formed. Excel now opens and refreshes a
  query loaded into a workbook openpyxl wrote.

  Saving through openpyxl still discards the Power Query package, since
  it rebuilds the file from the parts it models and drops the rest. That
  is not ours to fix, and `docs/power_query.md` now says so.

- **`[trash]` parts are dropped instead of carried forward.** Excel's
  file recovery leaves the parts it threw out under `[trash]/NNNN.dat`.
  An OPC part name cannot open a segment with a bracket, and Excel will
  not open a package holding one: the same workbook opens before the
  entry is added and raises after. The container preserves entries as
  they arrive, which handed back a file that stayed broken, so `save()`
  now takes them out and warns. Nothing depends on them: no content type
  declares them and no relationship points at them.

### Internal

- `openpyxl` joins the dev extras, so the interop tests run against the
  real writer rather than only against a copy of its output. The runtime
  still has no dependencies.

## [5.2.0] - 2026-09-05

### Added

- **`rename_form()` and `rename_report()`** ([#21]). A design's name lives
  in four places: its container's listing, its catalog row, its
  navigation-pane row, and the module behind it, which Access binds by
  name as `Form_<name>` or `Report_<name>`. A design Access has never
  opened a code window for has no module, and then there are three. The
  live gate renames a form that has code, opens the design in Access, and
  reads the code back through the VBE.

### Fixed

- **A stale `DocClass=` line made Access call the whole project corrupt**
  ([#21]). `PROJECT` names the module behind a form or report only as
  `DocClass=<name>/<flags>`, never as `Module=` or `Class=`. The renamer
  reached the first two and the workspace line and left `DocClass` alone,
  so a rename produced a project naming a module it no longer had. Access
  opened the database and showed the form, then reported the project as
  corrupt on the first VBE reference. The remover had the same gap, so a
  `DocClass=` line survived a delete. Both handle it now, carrying the
  flag word after the slash through untouched.

- **`delete_form()` and `delete_report()` left the module behind.** The
  design went and its code stayed, listed in the VBE for a form that no
  longer existed, and `delete_module` would not take it either because a
  design's module has no storage folder. Deleting a design now removes
  its module first, while the design is still there to name it.

- **The first module in a project with none went unlisted.** Module
  creation only updated the container's listing when the container
  already had one, where design and macro creation create it. It also
  needed an existing `Module=` or `Class=` line to insert after, so
  adding a module back after deleting a project's last one raised
  `ValueError` out of the PROJECT editor. Both containers now go through
  one helper, and the first line of an empty module block opens below
  `ID=`, where Access puts it. Found by the xlide_vscode port, which hit
  it first on a blank template with no modules at all.

[#21]: https://github.com/WilliamSmithEdward/pyOpenVBA/issues/21

## [5.1.3] - 2026-09-05

### Fixed

- **`delete_module` took a module's storage folder by position** ([#19]).
  The position came from `modules()`, which is dir-stream order and
  includes the module behind a form or report. Those have no folder under
  `Modules`, so the two lists were different lengths and every module
  after a document one sat a place high. Deleting the last module of a
  project that has a coded form raised `IndexError`, and the same
  arithmetic quietly took another module's folder in the cases that did
  not run off the end. The folder now comes from the container's own
  `DirData`, which names it, the way `_delete_design` already read it.
  Asked to delete the code behind a form, `delete_module` now says so
  instead of taking the next module's folder.

- **Decimal keys and values rounded at 29 digits** ([#20]). Both codecs
  scaled through Python's default arithmetic context, which rounds at 28,
  while the column is a 16-byte magnitude. `2**96 - 1` came back five too
  high, so a key stopped naming the row it pointed at, and an index entry
  that does not match its row sorts wrong and never matches on lookup.
  The row and index codecs now share `scaled_decimal` and `scaled_int`,
  which run at a precision no value a column can hold will overflow. The
  index also rounds a fractional value the way the row encoder does,
  rather than truncating, so the two agree on the same input.

[#19]: https://github.com/WilliamSmithEdward/pyOpenVBA/issues/19
[#20]: https://github.com/WilliamSmithEdward/pyOpenVBA/issues/20

## [5.1.2] - 2026-09-05

### Fixed

- **The Access VBA writer read no PROJECTCODEPAGE** ([#18]). It encoded
  and decoded every ANSI string as latin-1: module source, the dir
  stream's name and stream-name records, the PROJECTwm entries, the
  PROJECT stream, and reference names and libids. Two consequences, and
  the first reaches ordinary English projects. Writing a module died
  outright on anything latin-1 cannot hold, so an em dash, a curly
  quote, an ellipsis or a euro sign in a comment raised
  `UnicodeEncodeError`; latin-1 and cp1252 differ over exactly the
  0x80-0x9F band those live in. And every non-1252 project was read
  through the wrong page, which is invisible on a round trip because
  latin-1 is a byte-identity codec and only shows once the text is
  displayed or re-encoded.

  The writer now resolves the project's declared code page the way the
  readers already did, through `encoding_for_codepage`, and encodes with
  `encode_mbcs`, so a character the page genuinely cannot hold folds to
  `?` in the ANSI record while the UTF-16 record beside it stays exact.
  That is what the VBE writes. A live gate has Access read an em dash, a
  curly quote, an ellipsis and a euro sign back out of a module written
  here.

### Added

- **The Access write path joins the language-matrix CI job.** The same
  per-code-page sweep the Excel writer has had since [#13] now runs
  against a database whose PROJECTCODEPAGE is each of the twenty-one
  pages, covering module source, native module names through add,
  rename and delete, and the cp1252 punctuation band on its own.

### Documentation

- **TypeInfo member names in `_designs.py` are cp1252, measured.** The
  issue raised the hardcode there as a possible second instance. It is
  not one: Access named a control with an em dash and wrote one byte,
  `0x97`, which cp1252 gives U+2014 and latin-1 cannot represent at all,
  and patching the project's PROJECTCODEPAGE to 1251 changed nothing.
  That stream does not follow the VBA code page and must not be threaded
  with it. Access also drops a member whose name the page cannot hold
  rather than substituting it, which the writer here does not; it
  differs only for names cp1252 cannot express.

[#18]: https://github.com/WilliamSmithEdward/pyOpenVBA/issues/18
[#13]: https://github.com/WilliamSmithEdward/pyOpenVBA/issues/13

## [5.1.1] - 2026-09-05

### Changed

- **`examples/power_query_refresh.py`** no longer demonstrates a state
  Excel's dialog cannot build. Its third query removed its data before
  saving without also refreshing on open, which Excel draws as a ticked
  box greyed out beneath an unticked one. That query now does both, and
  is called `Transient` rather than `Manual`, since its rows are fetched
  when the workbook opens and dropped again when it is saved.

### Added

- **What that pairing actually is**, in `docs/power_query.md`. The
  dialog's wiring is not a rule about the file. Excel's object model sets
  `QueryTable.SaveData = False` with refresh-on-open off and raises
  nothing; Excel writes the same shape itself when refresh-on-open is
  ticked, remove-data is ticked, and refresh-on-open is then unticked;
  and Excel resaves a workbook pyOpenVBA wrote that way with the
  attribute intact. So `keep_data = False` on its own stays legal, and
  the only cost is a query whose table opens empty until someone
  refreshes it.
- `test_excel_keeps_remove_data_without_refresh_on_open` in the live
  gate, which holds that measurement, and a test that the shipped example
  writes nothing the dialog cannot reach.
- Live coverage for `examples/power_query_refresh.py`. All three of its
  queries now have to refresh in Excel, and Excel's object model has to
  report the profiles the example's own printed table claims. The other
  two Power Query examples already had this.

### Fixed

- The 5.0.0 entry's link to `docs/power_query.md` was relative, which
  resolves from inside the repository but not from the GitHub release
  page the entry was published to. It is an absolute URL now, in the
  changelog and in the published release notes.

## [5.1.0] - 2026-09-05

### Added

- **Refresh control on a loaded query.** `query.refresh` gives the
  settings behind Excel's Connection Properties dialog: `background`,
  `interval_minutes`, `on_open`, `keep_data`, `in_refresh_all` and
  `enabled`. None of them live in the Power Query package; they sit on
  the workbook connection the query loads through, and two are mirrored
  onto the query table. Where each is written was measured by toggling it
  in Excel and diffing the file, and two are not where the wording
  suggests: removing data before saving reads from
  `queryTable/@removeDataOnSave`, and staying out of Refresh All is
  stored as an exclusion in a connection extension. The live gate sets all
  six from Python and has Excel's object model read them back.

  "Enable Fast Data Load" is deliberately absent. Excel's object model
  does not expose it, so there was no way to watch Excel write it.

- **`examples/power_query_refresh.py`** gives three loaded queries three
  refresh profiles, and prints what each one will do.

- **`examples/power_query_steps.py`**, a second Power Query example built
  around a query's Applied Steps. It writes a sales pipeline whose main
  query has eleven steps, and keeps those steps in a list with `#PREV`
  standing for the step before, so inserting or removing one is a list
  operation and the references rewire themselves. Both the built workbook
  and an edited copy refresh in Excel.
- `quote_name()` and `unquote_name()` are public, so code that generates M
  can spell a name the way M spells it without reaching into a private
  module.

## [5.0.0] - 2026-09-05

### Added

- **Power Query, read and written in pure Python.** A workbook keeps its
  Get and Transform queries in one custom XML part: an OPC package
  holding an M section document, a metadata document beside it and a
  permission list, all base64 inside a `DataMashup` element.
  `PowerQueryWorkbook` opens any Excel package -- `.xlsx` included, which
  has no VBA project at all -- and reads, edits, adds, renames, groups
  and removes the queries in it.

  - `queries()`, `query(name)`, `query_names()`, and per query the
    formula, description, steps, load target, group and every metadata
    entry Excel recorded.
  - `add_query()`, `remove_query()`, `rename_query()` (which rewrites the
    queries that name the old one, through the M tokenizer, so a match
    inside a text literal or a record's field name is left alone),
    `set_section_text()` for the whole document at once.
  - `add_group()` / `remove_group()` and `move_to_group()` for the
    folders in the Queries pane.
  - `load_to_sheet()` and `unload()`, which write and remove the
    connection, query table, table and sheet reference that actually put
    a query's result on a worksheet.
  - `PowerQueryWorkbook.create_new()` builds a workbook from nothing.
  - `pull_power_query()` / `push_power_query()` and `pq-ls`, `pq-pull`,
    `pq-push` on the command line: one `.m` file per query beside a
    manifest carrying the names a file name cannot hold.

  The writer aims at Excel's own bytes rather than at something Excel
  will accept. A workbook read and written unchanged is identical to the
  byte; a package rebuilt from its parts reproduces Excel's ZIP exactly,
  down to the raw deflate at level 6 and the growth-hint extra field; and
  the metadata section, the entry encodings and the query-group blob are
  byte-identical to what Microsoft's own packaging assemblies produce for
  the same input. Twelve live gates have Excel open what was written and
  evaluate it, and
  [docs/power_query.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/power_query.md)
  records how every rule was measured, including the three that a
  specification would not tell you: a query is a metadata item with at
  least one entry, the permission list and the metadata content package
  are both load-bearing, and the permission bindings are not.

- **`examples/power_query_demo.py`** builds a workbook of twelve queries
  in four groups, six of them loaded onto a sheet, reading JSON from four
  public APIs in five shapes: a list of records, one request per row, a
  list of bare numbers, GeoJSON features, a record whose field names are
  data, and a record of records. Every query in it was refreshed in Excel
  before it was committed.

### Changed

- The deflate port moved from `pyopenvba.access._deflate` to
  `pyopenvba._deflate`, where both the Access attachment codec and the
  Power Query package writer use it, and gained `raw_compress()` for the
  headerless stream a ZIP entry carries.

## [4.0.0] - 2026-09-05

### Added

- **Access has the same API as the other hosts.** `AccessDatabase` gains
  `module_names()`, `get_module()`, `set_module()`, `add_module()` (taking
  `VBAModuleKind`), `vba_project()` with `add_module` / `rename_module` /
  `delete_module`, and `pull_modules()` / `push_modules()`; the top level
  gains `push_access()` beside `pull_access()`, and the command line
  `access-push` beside `access-pull`. Forms and reports come back from
  `forms()`, `reports()`, `form()` and `add_form()` / `add_report()` as
  `AccessForm` objects with `walk()`, `control()`, `properties()`,
  `set_property()`, `add_control()`, `remove_control()` and `set_code()`,
  the calls a UserForm takes; a control is an `AccessControl` with the
  same `properties()` / `set_property()`. Removing a control re-marks the
  ones left in its section and closes the tab order up behind it; the
  live gate has Access read back what remains.

- **Compact and Repair.** `db.compact_and_repair()` builds a new database
  the way DAO's `CompactDatabase` does: a bare engine skeleton with every
  object copied in -- containers, forms and modules, then tables, queries
  and links in the order of the catalog's Name index, rows in primary-key
  order, indexes built over them, AutoNumber counters reset to the largest
  value present, permission rows and property blobs written in a last
  pass, relationships re-created. Jet 4 `.mdb` files take their own
  skeleton. Under a frozen engine clock five compactions -- tables with
  deletes, keys and memos; counters and queries; attachments; forms,
  reports, macros, modules, a link and seven query shapes; a Jet 4 file
  -- match the engine's output on every page but page 0. The
  result keeps the source's creation date, because the engine keys its
  encoding of owner and permission SIDs to that date (a 30-bit fold of
  five of its bytes feeds a generator that stayed unknown), so those
  bytes copy across unchanged. A table whose definition runs past one
  page is built column by column as the engine builds it, two writes
  per column onto fresh continuation pages, and the destination grows
  and lets rewritten pages back into use by the engine's rule (64 pages
  ahead, then 8 or 64 at a time, the waiting pages returning once 24
  pages have been taken since they last did), so the stale pages behind
  a wide table and every page after them land where the engine puts
  them: a sixth compaction, two 161-column tables with a memo, a key and
  rows, matches the engine page for page.

- **Fourteen more form property codes are named** -- `Glow`, `Shadow`,
  `QuickStyle`, `HoverForeColor`, `PressedForeColor` and the hover and
  pressed theme index, tint and shade properties -- by differencing a
  command button and a toggle button in Access, taking the table to 180.
  The four codes still keyed by number, 700-703, turned out not to be
  properties: they are the insets a button keeps for an effect that is
  off, dropped in favour of the paddings when Glow or Shadow is set.

- **`SELECT ... INTO` makes a table.** The make-table query creates the
  table the engine creates: a copied column keeps its definition, an
  expression gets the type the engine's own inference gives it (157
  expressions measured, every definition page byte-identical), the rows
  follow, and the three quirks a made table's headers carry -- computed
  columns variable-length, ordinals from one, a Decimal's precision and
  scale leaking into the columns after it -- are written as the engine
  writes them.
- **`UPDATE` and `DELETE` run over a join**, writing to either table of
  the join, applying every join row in the engine's order, and refusing
  what the engine refuses: a subquery in a SET clause, a join DELETE that
  does not name its table, a row a DELETE reaches twice.
- **A FROM clause with joins in parentheses reads** -- `(A INNER JOIN B
  ON ...) INNER JOIN C ON ...`, which is how Access writes every query
  over three tables, and the group on the right side too.
- **A numeric literal with an exponent** (`1E3`, `1.5E-1`) reads.

- **A table's rows can be packed onto fewer pages.**
  `db.rebuild_table(name)` reads a table's rows out, drops the table,
  makes it again from the same definition and writes the rows back, so
  they land on as few pages as they need; `db.compact(rebuild=True)` does
  that for every table it can and then reclaims the file's tail. A
  2000-row table cut to 200 went from 270 pages to 92.

  Deleting rows does not shrink a table -- not here and not in Access,
  which is what Compact and Repair is for -- so until now a large delete
  left the space stranded: free pages, but in the middle, where the
  trailing-run reclaim could not reach them.

  Both halves are writers the engine was measured against, so this adds
  no new way to lay a table out. It refuses a table it cannot carry
  across whole -- one with a complex column, a link, or a relationship
  naming it -- and it compares the rows before and after, putting the
  whole database back if they differ. A rebuild either round-trips or
  changes nothing. The AutoNumber values and their counter are kept,
  where Access's own compact resets the counter.

  `db.table_specs(name)` is the piece that made it possible: the table
  described as the `ColumnSpec` and `IndexSpec` list that would create it
  again, with sizes in the units those take rather than the header's
  bytes.

- **Charts and Edge browsers can be written**, and the reader knows two
  more types (navigation buttons and the navigation control's own kind).
  That takes the reader to 28 control types and the writer to 23. Naming
  a record is not needed to write one: a slot whose code has no
  established meaning is keyed by its code, which also expresses the one
  case a name could not -- an Edge browser carries code 450 twice, at two
  ids.

  Three records Access writes on these are deliberately left out. 596,
  597 and 600 put the control in a layout, which is where the designer
  drops a new one; with them written, Access stacks a chart under
  whatever else claims the same layout instead of leaving it where it was
  put. Isolating that took building the same three-control form five ways
  and comparing where Access reported each control.

  A navigation control stays read-only for a reason that is not about
  codes: it is not one control. One of its records names a sibling
  subform by name, and Access builds navigation buttons beside it, so one
  written alone would point at a subform that is not there.

- **A design's properties can be changed.**
  `db.set_control_property(form, control, "FontSize", 18)` and
  `db.set_design_property(form, "Caption", "My window")` write one
  property of one control, section, or of the form or report itself.
  Access reads every one back exactly as written -- captions, fonts,
  colours, sizes, control sources, tips and tags among them.

  `PROPERTY_CODES` grew from 40 names to 107 along the way. Codes whose
  values are small integers every other property also uses cannot be
  named by matching values, so those were named by differencing: build
  the same form twice, identical but for one property, and see which
  record moved.

  A record's id is its slot in that object type's own schema, so changing
  a property an object does not already carry means knowing the id Access
  would have given it. `PROPERTY_SLOTS` holds 959 of those across 26
  object types, every one read off an object Access itself wrote; across
  seven databases every id, code and value type agreed, and so did every
  length except the strings', whose length is their text's. A property
  the object already carries keeps its record where it stands; a name
  that is not in the table is refused rather than written somewhere it
  does not belong.

- **Every control type pyOpenVBA can read, it can now write.**
  `add_control` wrote a Label or a TextBox; it now writes all eighteen --
  CommandButton, ToggleButton, OptionButton, CheckBox, OptionGroup,
  ListBox, ComboBox, Rectangle, Line, Image, PageBreak,
  BoundObjectFrame, ObjectFrame, Subform, Tab and Page as well. Every
  slot's id, code, value type and width was read back from a control
  Access itself made, and each type gets only the slots it has: a page
  break carries a top and nothing else, a tab control no left or top at
  all, an image no overlap flags, a combo box its GUID ahead of its
  name. A button is written with the padding Access gives it, a list or
  combo box with `Table/Query` as its row source type, and anything that
  takes the focus with the next tab index.

- **Twenty-nine more design property codes named**, taking the table from
  40 to 69. Access's `SaveAsText` writes a design's properties with their
  names, so pairing that against the blob names the codes -- but pair on
  the *value*, not the position: the blob carries records the text does
  not write, so a straight walk drifts and starts naming codes wrongly a
  few records in. Both runs re-derived every code already named and
  contradicted none, which is what makes the new names evidence rather
  than a guess. `docs/research/access_designs/` carries the scripts.

- **Six control types the reader did not recognise at all**: custom
  (ActiveX) controls, attachment controls, the web browser, the
  navigation control, charts and the Edge browser. All six now parse and
  report their type; the first three are written as well, and Access
  accepts each. The other three are read but not written: each
  carries records this project cannot name -- 450, 456, 458, 600 among
  them, and a dozen more on a chart -- and on the navigation control and
  the Edge browser one of those, 450, appears twice at two ids, which a
  table keyed by property name cannot express. Writing one is refused
  with the reason rather than attempted.

  That takes the reader to 27 control types and the writer to 21. Two
  more codes Access accepts, 131 and 132, turn out to be alternate ways
  to ask for a subform and a text box: the file records them as 112 and
  109.

- **A control can hold controls.** `add_control(..., parent="Tabs")` puts
  a page on a tab control, which is written as a group of its own right
  after it. Reading a design is now a tree walk rather than a flat scan,
  since a section's count is of its own controls and not of everything
  beneath them. A page must have a parent tab and nothing else may have
  one, both of which are refused rather than written.

  The live gate puts one of every writable type on a single form and has
  Access name them all back, and builds a tab control with two pages
  between a text box and a button -- Access reports the tab as each
  page's parent and the form as the parent of the two beside it.

- **Jet 3 (Access 97) databases read.** `AccessDatabase` opens a 2 KiB
  page `.mdb` and reads its catalog, tables, rows and long values the
  same way it reads an `.accdb`. Everything that moved between the two
  versions lives in one `Layout` record, so there is a single parser, a
  single row splitter and a single set of value decoders rather than a
  second implementation: the page halves, a row counts its columns and
  its variable-column offsets in bytes rather than words, text is stored
  in the code page page 0 names rather than UTF-16, and the definition,
  column, index and name headers all shrink and move.

  Every offset was measured against files the Jet engine wrote, and the
  parser checks that what it consumed equals the length the page
  declares. The live gate has DAO 3.6 -- which still creates Access 97
  files though Access dropped the format in 2013 -- build a database with
  every Jet 3 column type, a 4000-character memo, code page text, a
  deleted row and four hundred rows over many pages; pyOpenVBA reads the
  same file with no COM involved and the two agree cell for cell.

  Writing a Jet 3 file is refused rather than attempted, at the page
  store itself, so nothing can put a Jet 4 shape into one.

- **Compaction.** `db.compact()` gives back the free pages at the end of
  the file and says how many went. That is what a dropped table or a
  large delete leaves behind, and it is the part of compaction that can
  be done without moving a page: nothing moves and no object is
  rewritten, so nothing can be lost. Free pages in the middle keep their
  place, and the first two pages are never dropped.

  A 954 KB database with a dropped 3000-row table came back at 300 KB
  with its other table intact. The live gate has the engine itself read
  a compacted file row for row and then run its own Compact and Repair
  over it, which rebuilds every page it kept.

  This is not Access's Compact and Repair, which also renumbers pages,
  resets AutoNumber counters and drops deleted rows from the middle of a
  table.

- **The domain functions.** `DLookup`, `DCount`, `DSum`, `DAvg`, `DMin`,
  `DMax`, `DFirst`, `DLast`, `DStDev`, `DStDevP`, `DVar` and `DVarP` run
  in `db.execute(...)`. Each is a query over another table, so each runs
  as one: `DLookup("Total", "Orders", "Id = 1")` is
  `SELECT Total FROM Orders WHERE Id = 1`. A criteria can name a column
  of the row it is evaluated in, which is what makes them worth having.

  These are Access's own rather than the database engine's, so DAO cannot
  answer for them: the gate compares every one against Access's `Eval` on
  the same database, and all seventeen agree.

- **The password guard reaches Access.** `db.vba_is_protected()` reads
  the `PROJECT` stream's `DPB` record, and `db.save()` refuses to write a
  VBA change into a protected project unless it is told
  `allow_protected=True` -- which is what the other hosts already do on
  their own `save`. A change that is not to the VBA project saves as
  before.

- **The project's references.** `db.references()` reads the libraries a
  VBA project points at, `db.add_reference(name, guid, major, minor,
  path=, description=)` adds one and `db.drop_reference(name)` removes
  it. A live gate adds the Scripting Runtime, writes a module that uses
  `Scripting.Dictionary`, and has Access compile and run it.

  Access keeps them only in the dir stream, three records each --
  `REFERENCEORIGINAL`, its Unicode twin, and `REFERENCEREGISTERED`
  holding `*\G{GUID}#major.minor#lcid#path#description`. `PROJECT`
  carries no `Reference=` line, and the two libraries every project has,
  VBA itself and Access, are not in the file at all. The version is
  written in **hex**: DAO 12.0 is stored as `c.0`.

- **Forms and reports.** `db.forms()`, `db.reports()`, `db.form(name)`,
  `db.report(name)`, `db.create_form(name)`, `db.create_report(name)`,
  `db.delete_form(name)` and `db.delete_report(name)`. A live gate opens
  what this writes in Access's own designer, and a created form also
  opens in form view.

  A design is a stream of property records, `<u32 id><u16 code><u32
  type><u32 width><u32 length><value>`, with the ids ascending inside one
  object. Three ids are not properties but markers that open the next
  object: `0xFE` a section, `0xFD` the next object at the same level, and
  `0xFF` a control, which carries a second `u16` naming its type. Every
  design measured -- an empty form, a form with a label and a text box,
  and a report with its three sections -- rebuilds byte for byte.

  `db.add_control(design, type, name, ...)` puts a Label or a TextBox on
  one, and Access reads back every measurement it was given. Thirty-three
  property codes are named, worked out by exporting a design with
  `SaveAsText` -- which writes the same records with their names, in the
  same order -- and walking the two together.

  A control belongs to a section and is written immediately after it, and
  its marker depends on **how many controls that section holds**: one is
  a single child, `0xFE`; two or more open a group, `0xFF` then `0xFD`.
  Access writes both in one report -- a page header holding one control
  and a detail band holding two -- and refuses each in the other's place,
  with "saved in an invalid format" for one of them.

  `db.set_design_code(name, code)` puts code behind one, creating the
  module when the design has none, and Access runs it. A document module
  belongs to its design rather than to `Modules`: no storage folder, no
  catalog row, and a `DocClass=` line in `PROJECT` where a class module
  gets `Class=`. Without that line Access loads the module and the form
  still does not answer to it. The design's `TypeInfo` and the module's
  `VB_Base` share a CLSID, and a byte in the design folder's `PropData`
  records that it has a module at all.

  Creating cuts from a captured empty design with a GUID of its own
  patched in, since the catalog row repeats the one the design carries.

- **Macros.** `db.macros()`, `db.macro(name)`, `db.create_macro(name,
  actions)` and `db.delete_macro(name)`, with `MacroAction(name,
  arguments)` for each step. A live gate creates a macro, has Access run
  it with `DoCmd.RunMacro`, and reads back the value it set.

  Access stores a macro as a binary blob, not as the XML its designer
  shows: a 32-byte header, a length-prefixed `"33"`, then one record per
  action carrying the action id, the row number, fourteen `u16` slots
  holding byte offsets into a string area, and the strings themselves.
  Arguments occupy slots from 4 upward and an empty one takes no slot,
  so a gap in the middle reads back as an empty string. Every blob in
  the fixture rebuilds byte for byte.

  Twenty-four action ids, measured by loading one macro each through
  `LoadFromText` and pairing storage folders with `MSysObjects` rows in
  id order. A macro's object id steps by **one** where a module's steps
  by four, and a macro gets no navigation-pane group row where a module
  does -- so the step is what an object reserves for itself rather than
  a global stride.

- **Attachments and multi-valued columns.** `db.complex_columns()` finds
  them, `table.attachments(column, key)` and
  `table.multi_values(column, key)` read them, `set_attachments` /
  `set_multi_values` write them, and
  `table.add_complex_column(name, kind)` creates one. An inserted row is
  given its complex id automatically. A live gate builds a database, a
  table and both kinds of column from nothing and has the ACE engine read
  the bytes back.

  Creating one costs four things beyond the column itself, and three were
  invisible until the engine refused the result. The flat table keeps its
  two Long bookkeeping columns **among the variable columns**, where a
  Long would normally sit in the fixed block, with no collation and no
  fixed bit -- `ColumnSpec(..., variable=True)` now says so. Its catalog
  row carries `Flags` `0x800A0000`. And the table that *has* the column
  carries `0x40000`, which no other table does; without it DAO opens the
  child recordset and finds no fields in it.

  A `Complex` column keeps only a Long in the row -- an id shared by
  every complex column in that row, handed out from a counter at 0x1C of
  the table definition and never reused. The values live one per row in
  `f_<GUID>_<Column>`, joined on that Long, and `MSysComplexColumns`
  names the pairing. `FileData` is a container of its own: a flag and an
  inflated size, then either a zlib stream or the bytes as they are, and
  inside that a header carrying the file's extension.

  Access decides whether to compress **by file type**, measured across 45
  extensions: it leaves `docx`, `gif`, `jpeg`, `jpg`, `png`, `pptx`,
  `xlsx` and `zip` alone and compresses everything else, including `7z`
  and `mp4`. Those eight are written byte-identically; a compressed one
  was not at the time -- see the fix above: the deflate is classic zlib's
  at level 5 and memLevel 7, which the zlib-ng behind Python's `zlib`
  cannot reproduce.

  Two corrections to the table definition came out of this: the field at
  0x18 is a constant (1 in ACE, -1 in Jet 4) and not a counter, and the
  complex-id counter is the u32 at 0x1C. A complex column is flagged
  AutoNumber like any other, so the row writer had been handing it a
  value from the ordinary AutoNumber counter, which gave two columns in
  one row two different ids.

- **Access VBA is writable.** `AccessDatabase` gained `modules()`,
  `module(name)`, `create_module(name, code, kind="module"|"class")`,
  `set_module_source(name, code)`, `rename_module(old, new)` and
  `delete_module(name)`. Standard and class modules both, with whatever
  source you give them.

  The route is not a p-code writer. `_VBA_PROJECT` is [MS-OVBA]'s
  PerformanceCache, and its `Version` field names the build of VBA that
  compiled it; writing a version the host does not recognise makes VBA
  discard the cache and compile the project from the module streams, the
  same thing Access's `/decompile` does. A module's stream is therefore
  the compressed source alone with MODULEOFFSET at zero, and none of the
  compiled tables have to be generated -- one of which could not have
  been, since the 32-slot table ahead of the module table is runtime
  state and adding the same module to the same database twice gives two
  different tables. The cost is a recompile on the next open: the code
  has to compile, and the cache stops matching what Access wrote until
  Access rewrites it.

  Three of the rules are invisible from the file and were caught only by
  asking Access: a module's storage folder is named from the rows its
  container already holds and Access will not look under any other name,
  `MSysObjects` ids step by four rather than one, and a delete has to
  free the folder or the next create picks a name Access rejects. A live
  gate (`RUN_LIVE_ACCESS_VBA=1`) runs the result in Access and compares
  the value the code returns, class instantiation included.

- **More of the statement.** `TOP n PERCENT`, `ORDER BY <position>`, and
  a comparison against `ALL`, `ANY` or `SOME` of a subquery. A column
  name two sources share is now qualified in the output the way the
  engine qualifies it, and a crosstab pivoting on a comparison names its
  columns -1 and 0.

- **The rest of the Jet expression functions.** Replace, Space, String,
  StrComp, StrReverse, Asc, Chr, Sgn, Sqr, Exp, Log, Fix, Val, Str, Hex,
  Oct, the CBool/CByte/CCur/CSng/CDate family, DateAdd, DateDiff,
  DatePart, DateSerial, TimeSerial, Weekday, WeekdayName, MonthName,
  DateValue, TimeValue, Time, IsNull, IsNumeric, IsDate, Switch, Choose,
  Format and Partition, with the First, Last, StDev, StDevP, Var and
  VarP aggregates. A gate runs 131 expressions through DAO and through
  the executor over the same four rows and compares them cell by cell.

- **Truth values as Jet writes them.** A computed comparison, logical
  operator or yes-or-no function now answers -1 or 0, and so does an
  aggregate over a Boolean column, which is what the engine answers. A
  Boolean column selected on its own still reads as a Boolean.

- **Linked tables.** `db.links()` and `db.link(name)` read the tables a
  database only points at; `db.link_table(name, database, source,
  connect=...)` writes one and `db.drop_link(name)` forgets it, byte for
  byte as DAO's `TableDefs.Append` and `Delete` do, for a link to another
  Access file and to a folder of text files. Following a link is left to
  the caller: the path comes out of the database, so opening it is not
  something the library does on its own.

- **A column's own rules.** Required, DefaultValue, ValidationRule and
  ValidationText are properties on the column, not bits in its header,
  and the writers now put them there the way the engine does: one blob
  write per column, the engine's DAO type and flags, and the catalog
  stamp that goes with it. `CREATE TABLE ... NOT NULL` and
  `DEFAULT <expr>` set them, and `ColumnSpec` carries them.

- **The engine's rules applied to every row.** A column an INSERT does
  not name takes its DefaultValue, evaluated as Jet evaluates it; a null
  in a Required column and a value against a ValidationRule, the
  column's or the table's, are refused with the engine's own message, on
  insert and on update. A live gate has the engine reject the same four
  statements.

- **Which LVAL page a long value lands on**, measured: a value of 256
  bytes or fewer goes on the first page the free-space map lists, a
  larger one on the last, and when that page cannot take it, on a new
  page rather than an earlier one.

- **A gate for a database built without the engine at all**: every
  column type, keyed and unique indexes, a foreign key, long values,
  four saved queries and table and column properties, all written by
  pyOpenVBA, then read back field for field by DAO and compacted by it,
  which fails on any structure the engine cannot follow.

- **Subqueries in saved queries.** A subquery in a WHERE, one used as a
  value, and a bracketed SELECT standing where a table does all save as
  the engine saves them; a derived table's text goes in its row's
  expression with only the alias naming it.

- **Pass-through queries.** `db.create_query(name, sql, connect=...)`
  saves one, byte for byte as DAO does, dead row and all; the SQL is
  kept exactly as given because the server parses it, not Jet.
  `query.connect` reads the connect string back.

- **Usage maps past their row.** A map whose bitmap can no longer grow
  inside its page becomes the engine's reference form: a row of chunk
  pointers, each naming a page holding one 32 736-page bitmap. The
  global free map converts the same way and marks each new chunk's
  unreached pages free. A 130 MB database of long values is now
  byte-identical to the engine's, both maps converted.

- **`with db.transaction():`** groups writes so they all land or none
  do; an exception puts the pages and the session's state back exactly
  as they were. The engine writes the same bytes either way, and a live
  gate checks that against DAO running the same statements inside
  BeginTrans/CommitTrans.

- **Crosstabs run, not just save.** `db.execute("TRANSFORM ... PIVOT
  ...")` returns the pivoted rows: one column per pivot value, `<>` for
  a Null heading, an `IN` list fixing the columns and their order, an
  aggregate allowed among the row headings, and the rows sorted by those
  headings as the engine sorts them. Two of these answer exactly as DAO
  does (live gate).
- **Jet's `Mod`, `\` and `^` operators**, in VBA's order of
  precedence, with both sides rounded half to even and the division
  truncating toward zero.

- **Subqueries, unions and saved queries in `db.execute`.** `IN`,
  `NOT IN`, `EXISTS`, `NOT EXISTS` and a scalar subquery work in any
  expression, correlated when they name the outer query; a bracketed
  SELECT or a saved query's name can stand where a table does in FROM;
  and `UNION` / `UNION ALL` fold left to right with a trailing ORDER BY
  over the result. Eight of these shapes answer exactly as DAO does on
  the same database (live gate).

- **DDL through `db.execute`.** CREATE TABLE (every Jet type word, named
  primary keys, inline and table constraints, foreign keys), CREATE
  [UNIQUE] INDEX with ASC/DESC columns and WITH IGNORE NULL, DROP TABLE,
  DROP INDEX, and ALTER TABLE ADD / ALTER / DROP COLUMN and ADD / DROP
  CONSTRAINT. Fourteen statements leave the same bytes DAO's Execute
  leaves for the same SQL (live gate). What the Jet parser refuses --
  `CHAR`, `DECIMAL`, `NUMERIC`, `WITH COMPRESSION` -- is refused here
  with the reason.
- **A table with a BigInt column carries the engine's version
  properties** (`FCMinReadVer`, `FCMinWriteVer`, `FCMinDesignVer`),
  written one at a time as the engine writes them.

- **Crosstab saved queries.** `db.create_query(name, "TRANSFORM ... PIVOT
  ...")` writes the rows DAO writes, byte for byte, with an `IN` list, a
  TOP, a join or a parameter; `db.query(name).sql` gives the statement
  back. HAVING is refused there, as the engine refuses it.

- **`db.drop_index(table, name)`**, byte-identical to the engine's DROP
  INDEX: the index's pages released with their bytes untouched, its
  usage-map row deleted, its records taken out of the definition with
  the indexes after it moved up, and the catalog row stamped. The
  primary key and an index a relationship rests on are refused, as the
  engine refuses them.

- **SQL executor.** `AccessDatabase.execute(sql, parameters)` runs Jet
  SQL against the engine in pure Python: SELECT with a column list or
  `*`, INNER / LEFT / RIGHT JOINs, WHERE, GROUP BY with Count, Sum, Avg,
  Min and Max, HAVING, ORDER BY, DISTINCT and TOP; the comparison,
  logical, arithmetic and `&` operators, LIKE with the engine's
  wildcards, IN, BETWEEN, IS NULL, `[parameters]` and the common string,
  numeric and date functions; INSERT ... VALUES, INSERT ... SELECT,
  UPDATE and DELETE through the row writers, with values coerced to the
  column type as the engine coerces them. Three-valued logic follows the
  engine. Eleven SELECT shapes answer exactly as DAO does on the same
  database, name for name and value for value, and an UPDATE plus a
  DELETE write the same bytes DAO's Execute writes (live gate).
- **Index row counters.** An index built over existing rows records how
  many rows it holds next to its distinct-key count. Every row that
  leaves the index takes one off, whether deleted or written by an
  UPDATE that names one of the index's columns, and the distinct count
  is capped at what is left; a null key in an ignore-nulls index costs
  nothing, the count stops at zero, and an unfiltered DELETE zeroes
  both. Inserts leave the row counter alone, as the engine does.

### Fixed

- **Code behind a written form sees its controls, and their events fire.**
  Access lists a design's sections and controls in a `TypeInfo` stream
  beside it, and VBA takes that list as the form class's members. A form
  written here carried the empty template's list, so `Me.<control>` failed
  to compile and a button's click never bound. The stream is now carried
  forward on every change the way Access carries it: a new member is
  appended with the next ordinal, a removed one drops out and the rest
  keep their ordinals, a renamed one moves to the end with its ordinal.
  Each member's type id was read off a form and a report Access built
  with one of every control and every section
  (`tests/live_access_test/designs_every.accdb`): a report's controls
  share one class index, a label attached to a control or a button inside
  an option group has a class of its own, an ActiveX control's entry
  carries 36 more bytes, and a form's header, footer, page and group
  sections are read as sections. Three copies of that form edited in
  Access sit in the fixture, and the same edits made here give the same
  streams.
- **`RowSourceType` takes effect.** The text is only what the property
  sheet shows; Access acts on a one-byte companion record, written now
  with the text (a value list, a field list, or none for a table or query).
- **A written control brings its type's control-defaults object.** Access
  keeps one nameless object per control type ahead of a design's
  sections, the type's theme-derived defaults, and reads a control's
  themed properties against it: a button written without it ignored its
  `UseTheme`, colours and gradient and came out as a default themed
  button. The first control of a type now writes the object Access
  writes (captured per type for forms and for reports), the run of
  top-level objects is marked as the group it is, and a button's six
  hover and pressed slots are in its schema.
- **A colour or font set on a control takes effect.** With the defaults
  object in place, Access reads the theme index ahead of the colour, so
  `set_property("ForeColor", ...)` now writes the -1 index Access writes
  beside it (and `Gradient` 0 with a button's fill, `ThemeFontIndex` -1
  and the pitch-and-family byte with a font name), measured one property
  at a time on nine control types.
- **`set_database_properties()`** sets the database's own options --
  `StartUpForm` to open a form with the file, `AppTitle` and the rest --
  writing the MSysDb property blob byte for byte as DAO's
  `Properties.Append` does.

- **A second text box gets its tab index.** The text box Access made for
  the slot table was the first on its form and so carried no `TabIndex`;
  every later focusable control does, and the writer now gives one to a
  text box as it already did to every other type (measured: three text
  boxes written without it came back from Access tabbed 1, 2, 0).

- **A large long value lands on the highest-numbered listed page.** The
  engine takes the highest page the column's free-space map lists when
  it has room, and otherwise a fresh page; it does not go back to the
  page the last value went to. The two coincide until a delete lists an
  older page again, which is what Access's own module save does, and the
  library kept a per-column cursor for the wrong rule. Measured with DAO
  and held by the long-value placement gate; the page store's
  `lval_cursor` is gone.

- A long value of 256 bytes or fewer now takes the first listed page of
  its column rather than the page the last value went to, which is where
  the engine puts it (measured on a compaction's 200 memos).
- A relationship on a column that already has an index shares that index,
  as the engine's does; a table may now be related to itself; the
  relationship rows go before the index root.
- Rewriting a definition no longer resets its complex-id counter.
- A new object's owner is taken from MSysDb rather than the commonest
  owner among the tables, which in a fresh `.mdb` was the engine's own.
- Access's own MSysNameMap and MSysAccessXML, whose OLE column has no
  free-space map, can be written to and copied.
- **Compressed attachments are byte for byte what Access writes.** The
  engine's deflate turned out to be classic zlib's at level 5, memLevel 7
  and a 32 KB window -- one of eight engine-written streams admits exactly
  that parameter set, and classic zlib reproduces all eight -- where the
  earlier note that it "was not zlib's" came from comparing against the
  zlib-ng that Python bundles. Since a Python's zlib may not be classic
  zlib and exposes no memLevel, `pyopenvba.access._deflate` carries zlib's
  own algorithm; a live gate attaches five files through DAO and finds
  each stored container identical to ours.

- **A number keeps the engine's type through every operator.** `db.execute`
  now answers an `int`, a `float` or a `Decimal` exactly where the engine
  answers a Long, a Double or a Currency/Decimal: `5.5` and its arithmetic
  are Decimals, `5.5 / 3` runs to 28 places, `Sum(Long)` and `Abs` are
  Doubles, `Currency * Double` is a Double while `Currency + Double`
  stays Currency, `Date + 1` is a Date (it was a float), a Large Number
  takes everything into itself, and `IIf` widens its branches. Measured
  through DAO's reported type of 135 expressions and the full
  operator-by-type matrix; a live gate holds 185 of them and
  `docs/access_engine.md` has the rules, including the two that look at
  the shape of an operand rather than its value.

- **What a query stores in a column now matches the engine**, measured
  statement by statement against DAO: over-long Text is cut to size
  rather than refused (a Memo is not), a Byte takes the low byte of its
  number, a number in a Date column is its serial, a value list with no
  column list covers the AutoNumber too, an explicit AutoNumber moves the
  counter and a Null there is an error, and no query updates an
  AutoNumber.
- **An action query is all or nothing.** An INSERT, UPDATE or DELETE that
  fails part way leaves the rows as they were, through a page journal in
  the store rather than a copy of the file. The AutoNumber counter and the
  header's row count keep what the attempted rows took, which is what the
  engine leaves too.
- **Joins, GROUP BY and DISTINCT answer in the engine's order.** An inner
  join scans the smaller side and probes the other newest-first, groups
  and distinct rows come out sorted by their keys; an UPDATE over a join
  keeps the last join row's write, and which row that is now follows.
- **One- and two-character Text values are stored uncompressed**, as the
  engine stores them; compression is applied only where it shortens the
  value. Rows holding such values were one byte off the engine's.

### Removed

- **Jet 3 (Access 97) databases are refused again.** Reading them worked
  and was gated against the engine, but the scope is now what the Access
  application opens today, and Access has not opened an Access 97 file
  since 2013. A 2 KiB page format is a second set of offsets through the
  reader for a format nobody authors; `git log` has the implementation if
  it is ever wanted back.

  The plumbing that carried it went with it: the `Layout` record, the
  page size and text encoding read from the file rather than fixed, and
  the layout argument threaded through the page, row and definition
  readers. One page format means one set of constants, in the modules
  that own them, so there is no second source for an offset and no
  parameter that can only take one value.

## [3.5.1] - 2026-08-31

### Changed

- The README now introduces the form designer where a reader starts.
  3.5.0 documented it in full, but the "Why use this?" pitch and the
  "good fit for" list still described module operations only, and the
  architecture map predated `forms.py`, `_oforms_records.py`,
  `_oforms_pages.py` and `_ppt_container.py`, the `forms` CLI command,
  and the `.xlam` and `.accdb` templates.  PyPI renders a project page
  from the README in the released sdist, so correcting it there takes a
  release.

The library itself is unchanged: 3.5.0 and 3.5.1 are the same code.

## [3.5.0] - 2026-08-31

### Added

- **UserForm designs are now read and written, not just preserved**
  (issue #15).  A form's *code* was always a module like any other; its
  *design* -- which controls exist, how they nest, and what their
  properties are -- lived in streams the library carried verbatim.  It is
  now a first-class surface, with no Office installed:

  ```python
  with pyopenvba.ExcelFile("book.xlsm") as wb:
      form = wb.add_form("Wizard", caption="Setup", width=300, height=200)
      form.add_control("Frame", "Shipping", left=12, top=40, width=200, height=80)
      form.add_control("OptionButton", "Ground", container="Shipping")
      form.add_control("MultiPage", "Tabs", left=12, top=140, width=280)
      form.add_page("Tabs", name="Review")
      form.control("Ground").set_property("Caption", "Ground shipping")
      wb.save()
  ```

  `host.forms()` reads the tree; `host.add_form()` composes one from
  nothing; `form.add_control()` / `remove_control()` / `add_page()` /
  `remove_page()` and `control.set_property()` edit it.  Containers
  recurse -- a `Frame`'s children and a `MultiPage`'s pages live in
  storages of their own -- and each is created and deleted with its
  storage.  Geometry is in points.  `python -m pyopenvba forms <file>`
  prints the tree; `--mask` gives the raw property bits instead.

- **Only what the developer set.**  MSForms stores a property just when it
  differs from that control's default, so `control.properties()` is the
  set the author chose -- which a live COM read cannot distinguish from
  inherited and default values.  That is the reason this belongs in a
  file-level library.

- **Writing is lossless.**  An unedited form saves back byte for byte:
  alignment padding, raw string bytes, pictures and any tail the property
  tables do not model are all replayed as read.  Bytes inside a record
  that the tables cannot explain are refused rather than dropped, and a
  form whose streams do not reconcile raises `FormParseError` rather than
  returning a partly guessed control list.

- **Verified against live Excel and live PowerPoint**, which is where four
  defects surfaced that no structural check could catch: an added control
  colliding with the last one's id (`NextAvailableID` is the highest
  handed out, not the next free), a MorphData record omitting reserved
  mask bit 31 ([MS-OFORMS] 2.2.5.2), a container written with a leaf's
  site, and a designer edit leaving the `_VBA_PROJECT` cache stale.

- **Path-addressed CFB navigation and editing**: `CFB.list_storages_at`,
  `list_streams_at`, `get_stream_at`, `write_stream_at`, `add_stream_at`,
  `add_substorage_at` (which can set a storage's CLSID), and
  `remove_storage_at` (recursive).  Nested designer storages repeat
  names -- every container owns an `f` -- so a name-based lookup finds
  whichever comes first in directory order.

- `VBAForm`, `FormControl`, `Size` and `FormParseError` are exported from
  the package root.

### Fixed

- **A UserForm edit left the VBA performance cache stale.**  Only module
  changes counted as mutating, so a designer-only save kept a
  `_VBA_PROJECT` cache describing the form's old members and Office
  refused to load the form.  A designer edit now invalidates it too.
- **`.ppt` was advertised but could not be read** (issue #17).
  `PowerPointFile` listed `.ppt` and failed on every real one with
  "No 'dir' stream found", which reads like file corruption and is not.
  Unlike `.doc` and `.xls`, a binary presentation's CFB root carries no
  VBA storage: the project is a whole CFB, zlib-deflated, inside an
  `ExOleObjStg` record of the `PowerPoint Document` stream, reached
  through the persist chain.  Both directions now work; the write path
  splices the record back in and shifts every absolute offset past it.
  Verified against live PowerPoint, each check run first against an
  untouched control: a rewritten presentation opens with its slides,
  titles and body text intact, and an edited macro returns the new value.

## [3.4.0] - 2026-08-03

### Fixed

- **Non-Latin module names were corrupted in the PROJECT stream**
  (issue #11).  The PROJECT stream is code-page ANSI per [MS-OVBA]
  2.3.1, but four sites hardcoded cp1252, so any rewrite of it -- add,
  rename, or delete -- re-encoded module names with `errors="replace"`.
  A cp1251 project containing `МодульТест` came out as
  `Module=??????????` while the dir stream kept the real name; Excel
  cross-checks those declarations, so the project was left internally
  inconsistent.  `serialize_project_stream`, `parse_project_stream`, and
  `parse_projectwm` now take the project's `code_page` (defaulting to
  1252 for standalone callers) and the save path passes it.  Verified in
  live Excel: a cp1251 workbook whose module is *named* `МодульТест`
  now compiles and returns `Привет, мир` from a Cyrillic-named function.
- **Vietnamese text was destroyed on encode** (issue #13).  Python's
  charmap codecs do no composition, so `'Tiếng Việt'.encode('cp1258')`
  lost every stacked-diacritic character -- and NFD does not help, since
  cp1258 stores `ệ` as precomposed `ê` plus a combining dot-below rather
  than its canonical decomposition.  The new
  `pyopenvba.vba.encode_mbcs` decomposes unmappable characters and folds
  each combining mark back into the base until the codec accepts the
  result, emitting the remaining marks as combining bytes.  Text the
  codec already encodes directly is returned byte-for-byte unchanged.
- **Code pages resolved differently on Windows than on Linux/macOS.**
  CPython falls through to the operating system's code-page registry on
  Windows, so `cp10000`, `cp20866`, `cp21866`, `cp28592`, and `cp28595`
  resolved there while raising `LookupError` elsewhere -- text in those
  pages decoded correctly on one platform and became latin-1 mojibake
  on another.  `_CODEPAGE_ALIASES` now maps 30 Windows code-page
  identifiers (Macintosh, KOI8, the ISO-8859 family, ISO-2022, EUC, GB,
  UTF-7, GB18030) to portable Python codec names and is consulted
  first, so every platform resolves identically.  Found by the new
  cross-OS CI job on its first run; two tests now assert portability
  against the pure-Python codec registry so a regression fails on every
  platform rather than only the affected one.
- **Unresolvable code pages failed silently** (issue #12).  Falling back
  to latin-1 now emits a `UserWarning` instead of quietly producing
  mojibake that survives round-trip checks.
- **ANSI and Unicode dir records are reconciled** (issue #12).  When a
  module's name, stream name, or doc string disagrees between its ANSI
  record and its UTF-16 partner, the Unicode record -- lossless by
  construction -- is now authoritative.

### Added

- **20-language code-page test matrix** (issue #13, ported from
  `xlide_vscode`): one native-language sample per supported code page,
  each asserting zero substitution bytes on encode, an NFC-normalized
  round trip, and a full write -> read -> list -> validate cycle on a
  workbook whose PROJECTCODEPAGE is that page, plus native-language
  module names for cp1251 / cp932 / cp936.  The zero-substitution
  assertion is the load-bearing one: with `errors="replace"` a wall of
  `?` round-trips happily.  Fixtures are generated by patching one
  template's dir record, so no per-language binaries are committed.
- **Dedicated cross-OS `languages` CI job** running that matrix on
  ubuntu and windows, mirroring the equivalent job in the port, so a
  code-page regression names its own OS.
- Live Excel gate case for a Cyrillic-named module (opt-in via
  `RUN_LIVE_EXCEL=1`).

## [3.3.0] - 2026-08-01

### Added

- **Excel fixture CI on real Office** (#4, contributed by
  @DecimalTurn): a Windows workflow that builds fixture workbooks with
  the checked-out pyOpenVBA (no Office needed for the build), installs
  Excel on the runner via the SHA-pinned `DecimalTurn/setup-vba`
  action, runs each fixture's macro over COM, verifies its sentinel
  output, and uploads a desktop screenshot on failure.  Path-filtered
  to fixture and harness changes.  Complements the local
  `RUN_LIVE_EXCEL` gate with per-PR live-Office coverage -- the
  `with_class` fixture is a genuine VBE-export-form class module, so
  the issue #1 bug class is now regression-tested on real Excel in CI.
- **`ExcelFile.create_new` supports `.xlam`** (Excel add-in), joining
  `.xlsm` and `.xlsb`.  The baked-in template is captured from a
  freshly Excel-authored add-in (`ThisWorkbook`, `Sheet1`, bare
  `Module1`) via the new `scripts/bake_xlam_template.py`, following
  the existing bake pattern.

## [3.2.0] - 2026-08-01

### Changed

- **Decompression is 1.76x faster, byte-for-byte identical** (issue #5).
  `decompress` now emits output with slice operations wherever the spec
  allows -- non-overlapping copy tokens move as one slice, runs of
  literal tokens within a flag byte extend once -- and recomputes the
  copy-token masks only when the chunk-local output size crosses a
  power of two.  Overlapping copies keep the spec's byte-at-a-time
  semantics.  Measured 12.4 -> 21.8 MB/s across the 31 module and dir
  streams in the live fixtures; new oracle-equivalence tests pin the
  optimized decoder against the original per-byte implementation,
  including identical error messages and offsets on malformed input.
- **Module source loads lazily** (issue #5).  Decompressing module
  source is 88-96% of the cost of opening a project, so
  `parse_vba_project` now decompresses only the first chunk of each
  module stream (enough for the `Attribute VB_*` header; for
  single-chunk modules it already is the whole source) and defers the
  rest until the first `VBAModule.source` access.  Stream lookup and
  MODULEOFFSET bounds checks stay eager.  Opening the large-module
  fixture for `module_names()` drops from 1.47 ms to 0.79 ms.  Two
  visible consequences: a corrupt chunk past the first one raises
  `VBAProjectError` at first access instead of at parse time, and
  `VBAModule` is now a regular class rather than a dataclass -- the
  constructor signature is unchanged, a new `source_loaded` property
  reports materialization, but dataclass-generated field equality and
  repr are gone (equality is identity).

### Added

- `decompress(..., max_bytes=N)` stops at the first chunk boundary at
  or beyond N output bytes and returns the chunk-aligned prefix.  Copy
  tokens never cross chunk boundaries (the decoder enforces it), so
  the prefix is byte-identical to the same range of a full
  decompression.

## [3.1.0] - 2026-07-22

### Fixed

- **Class modules built from VBE-exported `.cls` sources now compile in
  the host** (GitHub issue #1). `add_module(kind=VBAModuleKind.other)`,
  `set_module`, and `push_modules` normalize class sources from
  file-export form to stream form via the new
  `pyopenvba.vba.normalize_class_source()`: a leading
  `VERSION 1.0 CLASS` / `BEGIN` / `END` preamble is stripped, and
  `Attribute VB_Base` is inserted after `VB_Name` when missing.  On
  replacement of an existing module the prior header's `VB_Base` line is
  preserved, so document-module host CLSIDs are never overwritten.
  Previously a supplied header was written into the stream verbatim: a
  missing `VB_Base` made Excel raise "Invalid procedure call or
  argument" at the first `New` site, and a VERSION preamble in the
  stream raised "Compile error: Expected: end of statement" (both
  verified against live Excel, as is the fix).  Supersedes the 2.0.1
  guidance that callers must supply the `VB_Base` line themselves.
- `pyopenvba.__version__` reported 2.0.0 while PyPI shipped 3.0.x.  A
  new test pins it to the installed package metadata so the two sources
  cannot drift again.
- CFB `get_stream_in_storage` / `write_stream_in_storage` /
  `list_streams_in_storage` now operate on the named storage's own
  child subtree instead of linear-scanning the whole directory.  The
  old scan could read or overwrite a same-named stream in a different
  storage (two UserForms both carry `o` / `f` streams) and reported
  root-level streams as members of every storage.  The host facades now
  address `PROJECTwm` at the project root, where [MS-OVBA] 2.2.1 puts
  it.  Byte output for well-formed files is unchanged (verified by
  hashing a 25-case save matrix across all live fixtures).
- `python -m pyopenvba pull / push / ls` now route Word and PowerPoint
  files by extension instead of assuming Excel; legacy `.xls` / `.doc`
  / `.ppt` are accepted everywhere the modern extensions are.  `disasm`
  no longer advertises `.xltm` / `.ppam`, which no facade accepts.
- `python -m pyopenvba access-pull` delegates to
  `AccessReader.pull_modules`, so Access class modules export as
  `.cls` (previously everything was written as `.bas`).
- README support section named the wrong project; roadmap.md's link to
  the feature-gate matrix pointed outside `docs/`.

### Changed

- The MS-OVBA compressor's LZ encoder uses a 3-gram position index
  instead of re-scanning the whole window at every position: about 60x
  faster on the 17 KB large-module fixture (0.44 s to 0.007 s) and
  0.4 s on a 1 MB input.  Output is byte-for-byte unchanged -- Access
  validates OVBA cache blobs against exact compressor output -- pinned
  by new naive-oracle equivalence tests across random, repetitive, and
  boundary inputs.
- `AccessReader.pull_modules` walks the database's LVAL rows once
  instead of four times per call.
- `save()` emits pending module additions and deletions in sorted
  order, making multi-add saves byte-deterministic across processes
  (Python randomizes set iteration per process via string hashing).
- `ExcelFile`, `WordFile`, and `PowerPointFile` are now thin subclasses
  of a single shared implementation
  (`pyopenvba._host.VBAHostFile`), removing three hand-synchronized
  copies of the read/edit/pull/push/save pipeline (~900 duplicated
  lines).  The public API is unchanged and the refactor was verified
  byte-identical against the previous implementation on every live
  fixture and save operation.

### Added

- **Live Excel compile-and-run gate** (`tests/test_live_excel_gate.py`
  plus `tools/live_excel/`): builds a workbook with an export-form
  class module, runs its macro in desktop Excel under a popup-aware
  bounded harness (VBE modals are dismissed, captured, and reported
  instead of deadlocking the run), and requires a clean run plus the
  macro's sentinel output.  Opt-in via `RUN_LIVE_EXCEL=1` on Windows;
  skipped in CI.  Issue #1 shipped because "opens without a repair
  prompt" was the strongest live verification; this gate closes that
  gap.
- CI matrix now tests Python 3.14 (the classifiers already claimed it).

## [3.0.0] - 2026-05-24

### Added

- **`AccessReader`** (EXPERIMENTAL) -- pure-Python **read-only** support for
  Microsoft Access `.accdb` / `.mdb` (ACE / Jet 4) databases:
  - `AccessReader(path)` parses the 4 KiB page-layout file header and
    validates the ACE / Jet signature.
  - `iter_vba_modules()` yields every embedded VBA module (`VBAModule`
    dataclass with `name`, `start_offset`, `attributes_text`, `source`).
    Modules are discovered by scanning for MS-OVBA stream signatures and
    walking the LVAL page chains they live on -- no Access COM, no
    MSysObjects parser required.
  - `vba_module_names()` deduplicates shadow / undo copies and returns
    the live module name list.
  - `read_vba_module(name)` returns the user-visible source string with
    `\r\n` line endings preserved; matches Access COM
    `CodeModule.Lines()` output byte-for-byte (verified on a 1000-line
    Module + 1000-line Class + 500-line Module live fixture against an
    Access COM oracle).
  - Re-exported from `pyopenvba` as `AccessReader`.
  - Write path (re-compress + re-allocate LVAL pages) is not implemented;
    Access support is read-only by design.

### Changed

- **BREAKING**: Renamed `pyopenvba.access` module to `pyopenvba.access_read`
  and renamed the `AccessFile` class to `AccessReader` to make the
  read-only access posture explicit.
- Adopted strict static analysis: pyright `typeCheckingMode = "strict"`
  and a curated ruff lint configuration (`E, F, W, B, UP, SIM, I, RUF,
  PIE, C4, PERF, N, TC, RET, TRY`) now run clean across `src/` and
  `tests/` with 0 errors.

### Removed

- Pruned ~1800 lines of dead Access write-path / probe code and the
  associated tests that exercised never-public APIs.

## [2.0.1] - 2026-05-24

### Added

- **`synthesize_class_header(name)`** -- new public helper (importable from
  `pyopenvba`) that returns the standard eight-line attribute header for a
  plain VBA class module, including the universal `VB_Base` CLSID. It is
  now also emitted automatically by `add_module(kind=VBAModuleKind.other)`
  when a bare body is supplied, matching the existing behaviour for standard
  modules. Callers no longer need to construct or hard-code the CLSID
  constant themselves.

### Fixed

- **README relative links were broken on PyPI.** The links to `LICENSE.md`,
  `docs/roadmap.md`, `docs/architecture.md`, and
  `docs/ms-ovba-implementation-guide_v2.md` were relative paths that
  resolved correctly on GitHub but 404'd on the PyPI project page. All
  five occurrences are now absolute `github.com/blob/main/...` URLs.

### Changed

- Demo scripts (`create_new_excel_with_class_demo.py`,
  `create_new_with_class_demo.py`, `create_new_word_with_class_demo.py`,
  `inject_xlsb_with_class_demo.py`) updated to use the body-only
  `add_module` call, removing the manual `_CLASS_VB_BASE` constant and
  `DATAMODEL_HEADER` block.
- README Architecture section updated to include `synthesize_class_header`
  in the `__init__.py` public API listing.

## [2.0.0] - 2026-05-24

### Added

- **`WordFile`** -- full read/write support for Word macro-enabled files:
  `.docm`, `.dotm` (OOXML/ZIP), and legacy `.doc` (raw CFB/BIFF8).
  Exposes the same API as `ExcelFile`: `module_names()`, `get_module()`,
  `set_module()`, `vba_project()`, `save()`, `pull_modules()`,
  `push_modules()`.
- **`PowerPointFile`** -- full read/write support for PowerPoint
  macro-enabled files: `.pptm`, `.potm` (OOXML/ZIP), and legacy `.ppt`
  (raw CFB). Same API surface as `ExcelFile` and `WordFile`.
- **`WordFile.create_new(path)`** -- create a brand-new `.docm` from
  scratch without launching Word. Ships with `ThisDocument` and an empty
  `Module1`; opens cleanly with no repair prompt.
- **`PowerPointFile.create_new(path)`** -- create a brand-new `.pptm`
  from scratch without launching PowerPoint. Ships with an empty
  `Module1`; opens cleanly with no repair prompt.
- **`ExcelFile.create_new()` now supports `.xlsb`** in addition to
  `.xlsm`. The extension in the path controls which baked-in template is
  used.
- **`pull_word(document, dest_dir)`** / **`push_word(src_dir, document)`**
  -- disk-based pull/push helpers for Word, mirroring the Excel `pull()`
  / `push()` API.
- **`pull_ppt(presentation, dest_dir)`** / **`push_ppt(src_dir, presentation)`**
  -- disk-based pull/push helpers for PowerPoint.
- **`scripts/bake_xlsb_template.py`** -- bakes the empty `.xlsb` template
  blob into `_templates/__init__.py` using the same splice pattern as the
  docm/pptm bake scripts.
- Class module creation is now fully supported across all three hosts.
  When adding a class module via `add_module(kind=other)`, callers must
  supply the full attribute header including
  `Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"` (the
  universal VBA class CLSID); without it Office raises "Invalid procedure
  call or argument" on instantiation.

### Changed

- `pyproject.toml` description updated to reflect all three supported
  Office hosts; `word`, `powerpoint`, `docm`, and `pptm` added to
  keywords.
- README fully updated: tagline, supported formats tables, 30-second tour,
  `create_new` section, and pull/push workflow section now cover Excel,
  Word, and PowerPoint.

## [1.1.1] - 2026-05-22

### Fixed

- **Editing a document module's source via `set_module()` (e.g. `ThisWorkbook`,
  `Sheet1`) silently broke the workbook in Excel.** The leading
  `Attribute VB_Name = "ThisWorkbook"` / `Attribute VB_Base = "..."` /
  `Attribute VB_PredeclaredId = True` header lines that bind a document
  module to its host object were being stripped on a source replacement.
  Excel then re-compiled the module without those bindings and either
  silently dropped the code or showed an empty module in the VBE.

### Added

- **VBE-style body-only source edits.** `ExcelFile.set_module(name, text)`
  now accepts either a full source replacement (text beginning with
  `Attribute VB_*` or `VERSION ... CLASS`) or a bare body. When a bare
  body is supplied, the module's existing attribute header is
  automatically re-prepended, matching the VBE UX where the user only
  types the executable code.
- **`VBAModule.body`** property: read or write a module's executable body
  without touching its attribute header.
- **`VBAModule.attribute_header`** field: the contiguous leading
  `VERSION ... CLASS` block + `Attribute VB_*` lines + separator,
  captured at parse time.
- **`split_attribute_header(source) -> (header, body)`** public helper.
- **`add_module(name, body, kind=standard)` now synthesizes a minimal
  `Attribute VB_Name = "<name>"` header** when the caller doesn't supply
  one. Caller-supplied headers are passed through unchanged.
- **`add_module(kind=other)` requires an explicit attribute header.**
  pyOpenVBA refuses to invent class or document module headers since
  their host-binding metadata can't be safely guessed.
- **`rename_module()` re-keys the in-source `Attribute VB_Name = "..."`
  line** to the new logical name so the source matches the dir-stream
  binding.
- New `TestAttributeHeaderPreservation` test class covering:
  header splitting (standard, document, class, headerless),
  `set_module` body-only preservation on a document module,
  `set_module` full-source replacement,
  `add_module` header synthesis vs. caller-supplied,
  `add_module(kind=other)` rejection without a header,
  and the `VBAModule.body` property round-trip.

## [1.1.0] - 2026-05-22

### Added

- **`ExcelFile.create_new(path)`** -- create a brand-new macro-enabled
  workbook from scratch in pure Python, without ever launching Excel.
  The new file ships with a fresh VBA project containing `ThisWorkbook`,
  `Sheet1`, and an empty `Module1`, opens cleanly in Excel with no
  "found a problem with some content" repair prompt, and is ready for
  immediate edits via the normal `vba_project()` / `save()` flow.
- New `TestExcelFileCreateNew` test class covering write-out, expected
  modules, empty `Module1`, round-trip with user code, overwrite of an
  existing file, and creation of missing parent directories.

### Internal

- New `src/pyopenvba/_templates/__init__.py` module embedding a
  byte-for-byte clone of a freshly Excel-authored empty `.xlsm` as a
  zlib-compressed base85 constant. No binary fixtures are shipped in the
  wheel; the template is regenerated by `scripts/bake_empty_template.py`
  from `tests/live_excel_testing/freshly_touched.xlsm`.

## [1.0.1] - 2026-05-22

### Fixed

- **Excel rejected modules whose source spanned more than one 4 KB chunk**
  with *"An error occurred while loading <Module>"*. The MS-OVBA compressor
  was emitting raw (CompressedChunkFlag = 0) chunks for full 4096-byte
  blocks. Although spec-legal, Office itself never writes raw chunks for
  module source streams -- empirically confirmed against an Excel-authored
  workbook containing a 16,881-byte module (all five of its chunks were
  token-compressed). The compressor now always emits token-compressed
  (flag = 1) chunks for module source; raw chunks remain only as a fallback
  for adversarial 4096-byte high-entropy input that overflows LZ encoding.
- **Re-running an add-module workflow after a delete produced duplicate
  `PROJECT` entries**, which Excel treats as corruption. Calling
  `add_module(name, ...)` after `delete_module(name)` in the same save now
  cancels the pending delete and treats the operation as a source rewrite,
  matching Excel's own behaviour. `serialize_project_stream` additionally
  scrubs duplicate `Module=` and workspace declarations on every structural
  save, healing files that were corrupted by earlier versions.

### Added

- `demo/` folder containing a runnable end-to-end demo
  (`push_demo_module.py` + `test_macro_workbook.xlsm` + `demo.md`).
- New regression tests:
  - `TestCompress.test_full_chunk_emitted_as_token_compressed_not_raw` and
    `TestCompress.test_long_module_round_trip_through_excel_save` verify
    that no raw chunks are produced for realistic VBA source.
  - `TestLargeModuleFixture` uses an Excel-authored 16 KB module as an
    empirical anchor and round-trips it through pyOpenVBA's saver.
  - `test_delete_then_readd_same_name_does_not_duplicate_project_decl` and
    `test_save_heals_preexisting_duplicate_project_declarations` cover the
    PROJECT-stream fix.
- `tests/live_excel_testing/large_vba_module.xlsm` fixture (Excel-authored
  reference for multi-chunk module compression).

## [1.0.0] - 2026

Initial public release. Pure-Python read/write support for VBA projects
inside `.xlsm`, `.xlsb`, and `.xls` containers, covering CFB parsing,
MS-OVBA compression, module add/edit/rename/delete, `PROJECT`/`PROJECTwm`
serialization, `_VBA_PROJECT` cache invalidation, and round-trip
preservation including password-protected projects.

[2.0.0]: https://github.com/WilliamSmithEdward/pyOpenVBA/compare/v1.1.1...v2.0.0
[1.1.1]: https://github.com/WilliamSmithEdward/pyOpenVBA/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/WilliamSmithEdward/pyOpenVBA/compare/v1.0.1...v1.1.0
[1.0.1]: https://github.com/WilliamSmithEdward/pyOpenVBA/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/WilliamSmithEdward/pyOpenVBA/releases/tag/v1.0.0
