# Host feature-completeness plan

The target is the whole headless library: file editing, VBA execution,
the host object model, calculation, events, shapes and Power Query where
the host supports them. The order is **Excel, then Word, then
PowerPoint**. Access runtime expansion follows those three. Shapes
continue in both pyOpenVBA and pyOfficeEditor, as recorded in the
[roadmap](roadmap.md#shape-support-in-both-libraries).

Excel is not feature complete. The existing
[MS-OVBA gates](xlsm_feature_completeness_gates.md) cover VBA project
file editing. Neither passing those gates nor passing the current test
suite establishes completeness of Excel's object model or runtime.

For a compact status table grouped by functionality, see
[Excel API coverage](in_memory_engine/excel_api_coverage.md).

The owner renewed this as an active implementation goal on 2026-10-04,
without a time or effort budget. Work continues through measured, reviewable
increments against this checklist. pyOfficeEditor is a reference for Excel
file surfaces; the initial reference revision is
`4723383e5d3039763e64ecd2a3ee092786c91bb5`. Adaptations must share the
runtime's state and pass the same Office conformance and persistence gates.

The first increment implements outer-body `GoSub`/`Return` and computed
subroutine calls, with 42 measured Excel probes in
`tests/fixtures/vba_semantics/gosub.json`. It also fixes computed `GoTo`
index coercion and bounds. Labels inside structured blocks now support direct
and computed jumps,
GoSub returns, error handlers and Resume through procedure-wide instructions.
A further 253 raw native traces cover skipped headers, loop reentry, lexical
With context, array enumeration locks, header coercion and Select Case value
capture. Module-level DefType defaults now apply to declarations, implicit
variables, arrays, parameters and function returns, with explicit type
precedence, ReDim declaration/retyping rules and compiler restrictions over
105 additional native probes. Project edits refresh cached bindings. This
remains partial language coverage: full compiler type checking, references
and object lifetime need further conformance work. Runtime parsing now retains
malformed-body diagnostics without blocking unrelated valid calls, while the
explicit parser stays strict. Fifty-three additional native controls cover body
recovery and ampersand syntax; malformed module boundaries, broader reference
typing and VBE import normalization remain further work.
Numeric For bounds now respect typed array elements, UDT fields and qualified
module variables; Next reevaluates indices and follows resized array storage.
Thirty-one native controls cover those targets and skipped initialization.
Fifty-nine further compiler controls cover numeric counter storage and
implicit bindings. Validation rejects invalid counters before execution while
preserving nested UDT fields, With temporaries and unused invalid routines;
global source edits invalidate cached checks. General declaration and
function/member reference typing still need work.
Declaration suffix/As conflicts and typed references now replay 73 native
controls, including early/late-bound members, ByRef calls, setters and literal
CallByName names. Intrinsic typing is measured incrementally; full call-argument,
constant-expression and host reference validation remain open.
User-call argument compilation and returned-default-member calls now replay
268 native controls, including constants, primitive/array ByRef matrices, early/late
binding, named setters, ParamArray, recursion and deferred unused errors.
Memoized expression typing covers deep member chains. Host/library signature
coverage and complete object/interface compatibility remain separate gaps.
The VBA financial library now provides Rate/IRR/MIRR/NPV/IPmt/PPmt with 263
native controls, typed array and named-argument compilation, VBA convergence
and extended-precision NPV. Native Excel opens and reruns a saved project
using all six and a worksheet UDF. Pmt/FV/PV/NPer timing coercion is corrected.
Universal bit-level financial agreement, broader numerical domains and full
library/host signature coverage remain open work.

The next increment connects class `WithEvents` subscriptions and custom
`RaiseEvent`, with 15 live Excel cases, and worksheet/workbook/application
sinks for existing events, with 14 isolated live cases. Event signatures,
broader handler error/debugger behavior and missing host sources still
need conformance work. Calculation events report newly calculated sheets and pending notifications
from earlier disabled or limited calculation cycles.
Another 12 native traces verify that attaching sheet, workbook or application
listeners retains dirty manual-mode caches and UDF call counts until the next
calculation request, including the listener's subsequent Calculate event.
Eighteen native traces distinguish EnableEvents changes within an emitted
event from suppression of the following workbook/application notification.
Another 42 cover bounded Calculate-handler reentry, input edits, volatile UDF
call counts and the following calculation cycle.
Another 78 uninstrumented traces distinguish local recovery from unhandled
host/custom callbacks, caller and event-raiser error policies, prior Err state
and ByRef arguments passed to later subscribers. Ordinary calls preserve Err;
23 lifetime cases, class controls and worksheet UDFs were remeasured without
the harness's injected On Error handlers. Native debugger/reset choices remain
outside the headless API, which raises VBARuntimeError for an escaped callback.
Ten raw Erl controls distinguish language error-policy transitions from
Err.Clear and assigning zero to Err.Number, which retain the line context.

Save/close lifecycle work adds 22 native cases for BeforeSave, AfterSave
and BeforeClose, and a live persistence gate for XLSX/XLSM conversion and
VBA-project presence. Save prompts, omitted filenames, other SaveAs formats and parameters,
additional event sources and complete signature validation remain gaps.

Runtime source persistence now writes added standard, class and document
modules and edits to existing modules into the owning workbook's VBA
project. A native Excel gate opens the result and executes a macro that
calls all three module kinds. Unchanged projects keep their original bytes;
mutations use the file writer's protection, signature and compiled-cache
handling. Macro-free saves omit runtime source. Another open workbook does
not receive the owning project's modules. Module deletion,
runtime-created form designers, and complete format
coverage remain separate gaps.

Each open workbook now has an independent VBA project. Twenty-one native
probes cover qualified and caller-local macro lookup, nested ThisWorkbook
contexts, class/document invocation and Err propagation. Shared WithEvents
subscriptions call each subscriber's own interpreter; native Excel executes
two persisted projects subscribing to the same sheet. Workbook_Open and
application WorkbookOpen precede workbook activation, with four measured
traces and a native gate executing saved workbook/worksheet handlers.
Project references and qualified types, full compile-time checks, project
unload/lifetime behavior, event errors and events whose handlers change
windows or event settings still need conformance work.

Worksheet document work uses 22 native measurements for lazy code-name
allocation, Count truncation, sheet caption counters and copied/deleted
modules. Copy and cross-book Move transfer document source into the target
project with fresh fields; Delete removes the runtime and saved module.
Unloaded source copies without execution, and a native gate executes saved
copied code and checks deleted modules. Complete project component order/
line-count fidelity, grouped-sheet Add defaults, wider Count/Type/anchor
conformance and complete copied sheet
content remain gaps; these measurements do not establish all 22 operations
as implemented runtime interfaces.

VBIDE project/component and source-buffer operations have 128 native Excel
measurements in `tests/fixtures/vbide_projects.json`, exercised through VBA
in `tests/test_excel_vbide.py`. Pending text is independent of compilation;
invalid edits remain inspectable and can be saved. Native Excel opens a
saved edited workbook, recognizes renamed project/document identities and
executes both document and standard code. Component/project naming, basic
line operations and procedure queries are implemented for the measured
cases. References, designers, complete import/export, Find/CreateEventProc, complete
procedure boundaries (continuations/conditional code), ByRef procedure-kind
queries, project reset/lifetime and protected-project conformance still
need work. The generated `docs/vbide_checklist.csv` records all 228 VBIDE
members, including missing interfaces. Its source is the pyVBAReference
checkout at `a6c39d0617085297e88bec7de0ee8afb7aa0746b`. This is a partial
VBIDE model, not proof of the full interface.

Source-file work adds 29 native probes in `vbide_source_files.json` for
standard/class imports, standard/class/document exports, AddFromFile,
reference-name collisions and physical/editor line boundaries. Native Excel
executes saved imported standard and class code. Exported class headers are
converted to their stream form for persistence. Incomplete source stays
editable; blank/leading/trailing lines survive save/reopen. UserForm files,
wider extensions and attribute/collision/encoding behavior remain gaps.

Eleven native lifecycle probes in `vbide_compile_lifecycle.json` verify that
editing or renaming a component resets its module variables on subsequent
execution while preserving other modules. Adding/removing components,
renaming the project and reading source preserve existing module state.
Opening an invalid-source workbook retains its complete editable project;
it can be inspected, saved and repaired before execution retries compilation.
Explicit Python `load_vba` imports remain strict and atomic. Full compilation
checks and wider project lifetime remain gaps.

Twelve native probes in `vbide_event_edits.json` and
`vbide_compiled_cache.json` cover adding, replacing, removing and renaming
worksheet handlers and per-module compiled source caching. Structural edits
refresh handler bindings before dispatch. `ReplaceLine` updates editable and
saved text while preserving an already executed module until an insertion,
deletion or rename invalidates it, matching native Excel. Modules not yet
executed use the replacement. Unrelated invalid source during events,
class-instance lifetime and broader compiler invalidation still need work.

Eighteen native probes in `module_namespaces.json` establish a measured slice
of standard-module qualification and ByRef argument binding. Public module
variables and array elements alias storage; class public fields pass a
temporary property value. Parenthesized arguments likewise use temporary
values. References, qualified types/enums, full name resolution and compile
checks and broader default-member binding remain incomplete.

Twelve native probes in `instance_array_byref.json` and `record_byref.json`
verify ByRef aliases to fixed/dynamic/Variant arrays inside class and worksheet
instances and to user-defined record fields, nested fields and array elements.
Local arrays shadow instance fields, named arguments retain aliases and
parenthesized record fields use temporary values. Full static argument-type
checks and broader member binding remain gaps.

Twenty-four native probes in `record_copy.json` and `array_copy.json` cover
record assignment/returns, nested records/fixed arrays, retained object
references, Variant/typed array assignment, ByVal and parenthesized copies,
ParamArray storage references and borrowed array-element locks. Record copies
write existing field storage so active scalar/nested/array-element aliases
remain valid. Array replacement, ReDim and Erase report error 10 while an
element is borrowed, and locks release when a handler unwinds a failed call.
Complete static type checks and fixed-array assignment rules,
record layout/Len/LenB and wider coercion/locking behavior remain gaps.

Thirteen native record-array probes in `tests/fixtures/record_arrays.json`
cover fixed/dynamic/multidimensional allocation, independent record elements,
assignment, ReDim Preserve, Erase and fields borrowed during cell copying,
resizing and error unwinding. Record fields and nested fixed-array elements
retain their enclosing array storage ownership; array copies retain the
record element factory for future allocation. Wider record layouts, type and
fixed-array checks, coercion, scope/reference resolution and compiler
validation remain incomplete.

Nine native probes in `tests/fixtures/compile_demand.json` establish demand
compilation for whole-record call arguments. Before executing a procedure,
the engine checks statically known record arguments forced to temporary
values by parentheses, including dead branches. Unused invalid procedures
and modules stay deferred; ordinary function-result records remain accepted
as temporaries. Successful checks track weak module dependencies and are
repeated after module replacement or renaming. Full name/type/visibility
validation, references, overload/default-member resolution, compiler
signatures and broader expression inference remain incomplete.

Eight native probes in `tests/fixtures/record_qualification.json` cover
module-qualified record declarations, including scalar, fixed/dynamic array
and nested record fields, and whole-record expression compile errors.
Allocation and demand validation share record-type lookup and construct
fields in their defining module. Broader reference and visibility checking
remains incomplete.

Twelve native probes in `record_scope.json` and `record_returns.json` cover
private record visibility, qualified duplicate public types, local type
precedence and initialized function/Property Get return fields. Known
inaccessible or ambiguous record declarations are checked before executing
the procedure body, including declarations in dead branches. Unknown type
names, project references and full static type validation remain gaps.

Eight native probes in `tests/fixtures/record_type_checks.json` cover nominal
record argument and assignment types, matching function-result temporaries
and prohibited ByVal record signatures. Statically known mismatches fail
before executing the called procedure body. Expression inference retains
the defining module when resolving fields and return types. Application.Run
reports error 1004 for a target module containing a ByVal record signature;
direct execution reports the native compile error. Broader macro eligibility,
array parameter and general static argument checks remain incomplete.

Fifteen native probes in `record_array_arguments.json` and
`record_variant_calls.json` cover typed record-array matching and standard
module record Variant barriers. Known array/scalar mismatches and forced
array temporaries fail before procedure execution, as do Variant assignments,
Variant parameters and TypeName applied to these records. Broader intrinsic,
host late-bound call and public object-module marshaling rules remain gaps.

Thirty native 64-bit Excel probes in `record_lengths.json` cover record
Len/LenB widths, numeric alignment, pointer/object/Variant/string storage,
fixed arrays, nested records, enums and dynamic arrays. Fixed-length strings
preserve their width, pad/truncate storage and resolve constant sizes in the
defining module. LongPtr uses 64-bit LongLong storage. Enum-typed variables,
parameters, results, arrays and fields are Longs, with members read through
the enum's name (11 probes in `enum_storage.json`), and ReDim resizes a
record's dynamic array field (6 probes in `record_field_redim.json`). 32-bit
layouts and wider fixed-string coercion remain incomplete.

## What counts as complete

Every headless feature must have a recorded implementation status and
conformance evidence. Preserving a feature's existing bytes is useful,
but does not establish that it can be created, edited or executed.
Registering a member name also does not establish its behavior.

For each feature, the checklist must record:

- The supported operation, formats, arguments and return types, including
  error cases, defaults, named arguments and read/write behavior.
- The Python API and the VBA API, where applicable, and whether they
  operate on the same state.
- Office-authored fixtures and measured results for both ordinary and
  boundary cases. Record the host version and relevant locale settings.
- Offline regression tests replaying those measurements.
- A save/reopen gate for persisted changes, including live Office opening
  the result without repair and preservation of unrelated content.
- Dependencies on other features, remaining gaps and any explicit scope
  exception. An unsupported error is a reported gap, not a passing gate.

Use **VERIFIED**, **PARTIAL**, **MISSING**, **UNASSESSED** and
**EXCLUDED** per operation:

- **VERIFIED**: every argument, error case and persisted effect is
  measured against live Office and replayed offline.
- **PARTIAL**: implemented, with live measurements behind some of it;
  the note says what is not covered.
- **UNASSESSED**: implemented, but not yet checked against Office.
- **MISSING**: not implemented; using it reports itself unsupported.
- **EXCLUDED**: out of headless scope, with the reason recorded.

An area becomes VERIFIED only when its operations and evidence are
complete. Do not advance the completion milestone to Word while Excel
has unresolved in-scope gaps; apply the same rule before PowerPoint.
Maintenance of existing support in the other hosts continues.

Headless scope does not require a desktop UI. Features involving files,
network sources, binary workbook formats or external data must still be
classified explicitly; being unsupported today is not itself a reason
to exclude them. Existing protection and signing exclusions remain
visible in the roadmap. A completion claim must state any retained
exceptions rather than imply unrestricted parity with desktop Office.

## Reproducible discovery inventory

Run this from the repository root, with Python 3.10 or newer:

```console
python scripts/audit_excel_coverage.py --output excel-coverage.json
```

The report compares registered object members, cell formula functions
and M library names with the committed reference inventories. It lists
missing names, direct unsupported paths and unmapped reference types.
It needs neither Office nor third-party packages.

This report seeds the behavioral checklist; it does not compute a
feature-completeness percentage. The reference contains aliases,
interfaces, event names and UI members. Some model classes have no
matching reference entry, and some registered methods are stubs.
Dynamic dispatch and context-provided functions require separate review.

The initial audit of the v6.0.0 implementation at `48392b5` found:

| Surface | Registered names | Reference names | Interpretation |
| --- | ---: | ---: | --- |
| Application | 37 | 315 | Includes configuration values; events are not connected |
| Workbook | 17 | 226 | Counts only registered names also present in the reference |
| Worksheet | 17 | 104 | Many collections and operations remain absent |
| Range | 38 | 201 | Includes explicit unsupported paths such as Find and Merge |
| Shape | 19 | 79 | Reading and common edits do not cover the full surface |
| Shapes | 8 | 26 | Some operations remain unsupported; count is not a pass |
| VBA WorksheetFunction | 72 | 406 | Every member the cell formula engine has, through the engine |

That initial cell formula engine audit registered 122 names; its inventory listed
362 additional names without registrations. The M library registers 234
names, of which 225 match its normalized 859-name reference inventory;
634 reference names have no library registration. These figures include
names such as constants and types and are not counts of proven functions.
`Excel.CurrentWorkbook` is provided by the workbook evaluation context
and must be checked separately from the static M library registry.

## The operation checklist

[`excel_checklist.csv`](excel_checklist.csv) has one row for every member
of every Excel class the model implements, with its kind and signature
from the type library, a status, the fixtures that are its evidence, and
a note. Its `interface_gaps` column lists what a registration visibly
lacks next to the type library: a read-write property with no setter,
or a method missing some of Excel's parameters. A member with an
interface gap cannot be VERIFIED.

Regenerate it after adding or removing a member:

```console
python scripts/build_excel_checklist.py
```

The script reads the type-library dumps in the sibling pyVBAReference
checkout, the same source as the object inventory, and keeps the status,
evidence and note columns already in the file.
`tests/test_excel_checklist.py` fails when a registered member is listed
as missing, a missing one is listed as implemented, an implemented class
has a member the file leaves out, or a pass names evidence that does not
exist.

The first pass marked PARTIAL only the members whose behavior a measured
fixture exercises by name, and left every other registered member
UNASSESSED. Individual operations now have VERIFIED evidence; the
[Excel API coverage summary](in_memory_engine/excel_api_coverage.md) reports the
current counts. No complete functional area has finished its audit.

## Excel work areas

All rows below require an operation-level audit before they can become
VERIFIED. This is the initial work breakdown, not an exhaustive list of
Excel features. Reference members and file-format features must be
assigned to these areas or added as new areas, with every unresolved
entry tracked explicitly.

| Area | Current evidence | Remaining audit and implementation |
| --- | --- | --- |
| Packages and formats | VBA editing covers xlsm, xlsb, xlam and xls; runtime loads xlsx, xlsm and xlam; a workbook whose parts spell SpreadsheetML's namespace under a prefix reads and saves as the same workbook spelled as Excel spells it, a pivot cache's records kept as the file had them, over seven copies of Excel's workbooks and a live gate | Inventory creation, open, save, conversion and persistence by format; xlsb and xls cell models are absent; a part that mixes that prefix with a default namespace or puts an attribute under it is not read |
| VBA projects and forms | In-scope MS-OVBA gates, designer round trips and live file gates exist; shared library-reference add/remove operations pass native Excel/Word/PowerPoint checks and Access persistence tests | Reconcile every operation with its format and protection/signing limits; distinguish form design from form runtime; linked VBA-project creation and arbitrary library runtime resolution remain separate gaps |
| VBA language and runtime | Measured semantics, classes, errors and control flow exist; every constant of VBA's own type library, the colour, key-code and system-colour modules included, typed as declared; Scripting.Dictionary in memory, its key matching measured; VBScript.RegExp in memory, Test, Replace and Execute in VBScript's dialect over 56 live cases; FileSystemObject's path functions over 37 live paths, its disk access reporting itself unsupported | Audit complete syntax, type system, lifetime, references and intrinsics; track file I/O, external calls and known typing gaps explicitly |
| Application, workbooks and sheets | Core collections, cells, names and queries exist; groups of sheets -- `Sheets(Array(...))` and the whole collection, selected, copied, moved, deleted and hidden, `ActiveWindow.SelectedSheets`, and a group's tabs saved and read -- and a sheet deleted or hidden, its readers turned #REF! and the next visible sheet active, follow 74 live cases, 608 group moves and 48 group copies, and a live gate; a sheet's tab colour -- `Worksheet.Tab` with Color, ColorIndex, ThemeColor and TintAndShade, Excel's errors, its 16-bit tints, the ColorIndex a tab finds for 1,397 colours off the palette, and the sheet's tabColor saved and read -- follows 47 live cases and a live gate; sheet protection -- Protect, Unprotect, the Protection object and every edit a protected sheet refuses a macro, with Excel's own errors -- follows 168 live cases, and is saved and read as Excel saves and reads it, SHA-512 and legacy passwords included, over 13 files and a live gate; workbook protection -- Protect, Unprotect and the sheet changes a protected structure refuses -- follows 41 cases and 5 files, with a live gate; windows -- ActiveWindow, Windows, Goto, each sheet's zoom, view, display settings, scrolling, frozen panes and selection, the tabs and the window order -- follow 113 live workbooks, saved and read as Excel saves and reads them; each workbook keeps its own 1900 or 1904 calendar, `Workbook.Date1904`, which cells, typing, Text cells, Find and Replace, AutoFill, TextToColumns, filter and validation criteria, Evaluate and WorksheetFunction follow as Excel's do over 8 native controls and 14 probes, native Excel reopening the model's files; `Application.ReferenceStyle`, `International` and the separators answer as en-US Excel does, the R1C1 style reaching Evaluate and the file's calcPr; `Worksheet.Range`, `Union` and `Intersect` refuse another sheet's cells as Excel does over 16 cases; a sweep of 77 everyday macros agrees with Excel throughout | Map all reference members; identify stubs, missing collections and incomplete arguments; allowed edit ranges and copying or moving sheets between workbooks while one is protected; a split that is not frozen, which Excel keeps in twips of the window; NewWindow, hidden windows and a window's size and state; a PivotCache holding dates in a 1904 workbook, Power Query dates on the 1904 calendar (followed, not measured), and separators or a locale other than en-US's |
| Ranges and editing | Values, formulas, areas and basic movement exist; written values are typed as Excel types them, with the formats typing brings, prefix characters, Text cells and Value/Value2 types, pinned by 4,902 live writes and a saved workbook, and every format with an `@` and no digits or dates makes a Text cell as Excel's does, over 50 formats under 11 writes; formulas a macro writes are spelled as Excel spells them over 120 live formulas; NumberFormat codes round-trip Excel's spelling on both sides, and a cell's or a rule's is refused where Excel refuses it -- a section mixing a number's, a date's and text's characters, a section out of place, a bracket Excel does not read, a fill or an escape with nothing to act on, an `E` or an `N` where it is no code, an accented letter read as the letter under it -- over 664 printable characters in four kinds of section and 330 whole codes; Range.Text, TEXT and WorksheetFunction.Text follow Excel's number-format engine over 156 live codes and 28 values, and round times as Excel does over 152 times on the half second; in a column too narrow, Range.Text cuts a General number to fewer decimals, scientific notation or 0 and shows anything else as #, and repeats a `*` fill as often as it fits, over every width from 2 to 130 pixels for 31 values and 22 fills and pads in Aptos Narrow and Calibri 11, regular, bold and italic, and at ten sizes from 8 to 36pt; CurrentRegion and End pass 34 live layouts, FillDown/Up/Right/Left 16 and SpecialCells 18; single-area Find/FindNext/FindPrevious pass 48 native cases; Replace passes 80 live layouts and the formula-bar text of 20 numbers in 26 formats; the clipboard's Copy, Cut, Paste and PasteSpecial pass 49; Range.Sort and the Sort object pass 58, in Excel's order for text over 583 values and 400 strings of Latin, Greek and Cyrillic with their accents, marks, punctuation and symbols, case ignored and matched, each string's key Windows' own word-sort key over 1,500 more, and the Sort object and a table's sort on a cell's colour or its font's as Excel does over 26 more, saved and read as Excel saves them, and the Sort object's on a conditional format's icons over 8 more; AutoFill passes 208 layouts and 520 columns, its series of numbers, dates, times, text and names to the last digit; RemoveDuplicates passes 98, text beyond ASCII among them by its word-sort key; TextToColumns passes 54 of 57 live cases of delimiters, qualifiers, typing, FieldInfo and its six date orders, separators and fixed widths, the other 3, FieldInfo pairs out of order, reporting themselves unsupported; Range.Range passes 29 live probes and a space-separated intersection 5; AutoFilter passes 113 layouts, reads and writes 41 filters as Excel saves them, filterMode included, and passes a live gate, and filters text beyond ASCII as Excel does over 36 spellings, 51 criteria, 6 lists and 13 pairs of letters under a wildcard, with a second live gate; filters on a cell's colour, its font's, no fill and the automatic font show the rows Excel shows, read back as Excel's do and save with their dxfs in Excel's order, on a sheet and a table, over 15 live layouts saved and opened again and a live gate; and filters on dates by year, month, day, hour, minute or second, as the macro recorder writes them, filter, read back and save as Excel's do, over 26 live layouts, three files written the classic way and a live gate; and the 32 dynamic filters on dates, today to next year and every quarter and month, over dates around the day Excel made them and a live gate; and filters on the icon a conditional format's icon set shows, or on none, each cell showing the icon Excel shows it, over 13 icon sets filtered by every icon and a live gate; edits on a filtered sheet -- writes, formats, clears, fills, copies, sorts, deletes and inserts, and the filter following rows and columns -- match 190 of 195 live layouts, the other 5 reporting themselves unsupported (two-dimensional arrays over several areas); SUBTOTAL passes 31 live layouts over hidden, filtered and nested rows and references that functions and operators give; every edit above meets sheet protection as Excel's does; Cut moves the references that overlap it as Excel does over 24 live layouts; cell shifts -- Insert and Delete with a Shift or none -- move references, names, the filter's range and formats as Excel does over 47 live layouts, and Insert's CopyOrigin takes the new lines' formats, heights and widths from before or after them over 5 more; a second sweep of 77 everyday macros -- copies, pastes, cuts, sheet copies, names, finds, filters, lookups, VBA's conversions, regular expressions and paths -- agrees with Excel, a value past 32,767 characters cut as Excel cuts it; a third sweep of 65 -- newer worksheet functions, tables, the Sort object, filters, conditional formats, data bars and icon sets, outlines, workbooks and VBA's Format and Split -- agrees with Excel, and Hidden on the rows or columns EntireRow or EntireColumn makes of cells apart hides each and reads the first's, as Excel does; merged-range geometry/value operations pass 21 native cases and four persistence gates; FormulaR1C1 scalar and array writes/reads pass 27 native cases and five persistence gates; A1 Formula array writes share the implementation and pass 15 native cases plus four persistence gates; Range.Rows and Range.Columns answer the range taken a row or a column at a time, Count, items and For Each by its rows or columns and every other member on its cells, over 79 live probes | Complete multi-area/comment/format/MatchByte searches, complex across/multi-area merges and merge formatting; Range.Text in a hidden column, which shows what Excel drew there before, in fonts other than Aptos Narrow and Calibri or sizes past 8 to 72pt where the text may not fit, and of a negative number General shows beside a fill at 12 to 16 pixels, as Excel shows it; the NumberFormat code `0;0;@`, which Excel rewrites as `0;0;0;@` though it keeps `0;@` and `General;General;@`; typing a space before a month's name and a time, or digits after AM or PM and a point, which Excel misreads into numbers; a table's sort on a conditional format's icons, which reports itself unsupported; a dynamic filter on dates in a file opened on a later day, which the model applies to its own today and Excel has not been measured on; audit insertion, deletion, copying, sorting and filtering |
| Named ranges | Workbook/worksheet CRUD, validation, scope precedence, rename/retarget/delete, visibility/comments and Python snapshots share one implementation; 19 native cases and three persistence gates; Names.Add over a table's structured references, the sheet a name is read on deciding what Excel refuses, an unqualified Add reaching the active sheet's own name, and Names(...) finding the first sheet's own name before the workbook's, follow 39 live cases with Excel's messages; names and sheet names beyond ASCII are kept apart and found as Excel does over 25 sheets, 27 names and 17 pairs of spellings that case, accents, ligatures, a sharp s, Greek and Cyrillic letters and combining accents set apart, with Names.Add redefining a name word sort cannot tell from the one given and Worksheets(...) finding a sheet so | Active-cell-relative name semantics, locale translation, Excel 4 macro metadata and arbitrary named-formula evaluation remain incomplete |
| Multiple workbooks and sheet copying | Add/open/activate/close; VBA sheet Copy/Move before/after or into a new book, matching native held-object invalidation on cross-book moves; inter-book range copies import used names and aliases with reuse/rename/error policies, backed by native probes and persistence gates | Making or updating external workbook links (a closed link's cached cells are read and saved unchanged), project references and unload/lifetime behavior, sheet-copy/move name conflicts, relative/other-sheet name imports, transferred drawings/controls, sheet modules and advanced worksheet metadata remain incomplete |
| Structural edits | Whole-row/column insertion/deletion and formula/name reference updates pass 12 native cases and four persistence gates; occupied-edge insertion is atomic; cell shifts and CopyOrigin as under Ranges and editing; conditional formats move with inserted and deleted rows, columns and cells as Excel moves them over 303 live layouts -- their ranges, formulas and points, and the rules Excel makes anew at its new priorities -- with Excel's file edited and saved again and a live gate; `Range.Subtotal` and `RemoveSubtotal` put subtotal rows into a list and take them out as Excel's Data > Subtotal does -- the list a cell's region or a range cut to the used block, groups by the text a cell shows with case and outer spaces aside, a bold label and `=SUBTOTAL` per TotalList column, a Grand row, every function by constant or the dialog's number, labels moved left of a totalled GroupBy column or into a column put in, rows taking the formats of the row they follow or above the data the row after, the outline and `Outline.SummaryRow`, Replace and Replace:=False over subtotals already there, page breaks between groups with the setting Excel keeps on the sheet for a later call, Excel's errors, the first label's leftover number format, Calculate and no Change -- over 166 live cases, 21 more of its page breaks, each cell's xf in the saved file, Excel's file read back and a live gate | Relative name context, merges/shapes and table/chart metadata updates remain incomplete; a filtered list, which Excel subtotals over its visible rows and which reports itself unsupported; the first label's leftover number format after a method other than ClearContents or Sort, a row put in or a WorksheetFunction reading a cell, which the model leaves as it was; conditional formats that Excel splits (a relative reference left reading at different distances), joins (rules alike beside one it makes anew) or numbers in ways not measured (several ranges made anew at once) report themselves unsupported |
| Range copying | Explicit-destination copies pass 17 native cases and five persistence gates: overlap, blanks, repeated blocks, formulas, cross-sheet copying and measured formatting | Merged-cell and multi-area copies, destinations larger than 1,048,576 cells, clipboard/paste behavior and full formatting fidelity remain incomplete |
| Formatting and layout | Cell formats are complete for Font, Interior, Borders, alignment, protection, NumberFormat and ClearFormats: 161 live probes, an Excel-authored workbook read back value for value, the border storage rules, a live gate for model-written formats, and the stylesheet Excel writes for the same edits, entry for entry. Row heights, column widths and hidden rows and columns follow Excel on a 96-DPI display: 185 live probes, Excel-authored and foreign files read back, the saved XML Excel's, and a live gate. Rows grow with their fonts from a 25,662-row table measured for eleven font families in four styles, beside the Normal font and alone on a row. Row, column and sheet-wide formats follow Excel: 149 live probes, Excel's saved workbook read back, 16 foreign files read and written back as Excel does, the model's saved XML Excel's, and a live gate. Medium, thick and double borders make rows taller as Excel's do, 39 more probes in the same fixture. Data validation reads, adds, changes, deletes and checks rules as Excel does, each rule moving with inserts and deletes, splitting round a delete or a Clear and carried by a copy, over 39 live workbooks and the dataValidations element Excel saves. Named cell styles -- `Range.Style`, `Styles` with `Add` and `Merge`, and the `Style` object's properties, read, set and deleted -- answer as Excel's do over 11,147 live reads and 6,459 more of styles a macro changes, deletes and merges: the 47 built-in styles given to cells in both orders, styles over formats and formats over styles, which cells follow a changed style, Normal's font, protected sheets, names beyond ASCII refused, found, listed and saved as Excel tells them apart by Windows' word sort over 38 names and 19 lookups; the stylesheet Excel writes for them, in the order its tables hold what a session made, under two themes and after a reopen, over 45 saved workbooks, files holding entries nothing uses, entries twice and custom formats out of order among them, and a live gate opens ten of the model's in Excel, which reads them as its own; an edge two cells' styles both draw reads as Excel reads it over 415 pairs; column AutoFit makes a column as wide as Excel does for the text its cells show, each character of ASCII and Latin-1 as wide as GDI makes it in Aptos Narrow or Calibri in any style, numbers in General a digit wide, over nearly 14,000 live widths of characters alone and beside others, 1,560 runs to the widest column, 1,302 random texts, 138 numbers in their formats and a saved workbook at 11pt, and 31 texts at each of 35 sizes from 8 to 72pt in both fonts and four styles, a big font's pad and italic's reach included, and Range.Columns.AutoFit fits a range's cells alone; setting NumberFormat, NumberFormatLocal, Style or Font.Size widens a column wherever AutoFit would make it wider for a number the cells show, as Excel does over 34 live cases, seven sizes, a font's size set on numbers and dates at six sizes and a saved workbook; the Normal font changes between Aptos Narrow 11 and Calibri 11, which make the same digit and row; colour scales of two or three colours, made by AddColorScale, criteria and colours read and set, show each cell the colour Excel blends for it, over 29 live layouts and all 246 colours measured, and are saved and read as Excel saves them, with a live gate; data bars made by AddDatabar, with their points, colours, fills, borders, directions, lengths, axes and negative bars, and icon sets made by AddIconSetCondition, with Excel's 20 sets, custom icons and every criterion, read and set as Excel's are over 248 live layouts, hidden values hiding a number's text, and saved and read as Excel saves them, Excel 2010's worksheet extension included, with the ids Excel keeps and makes for rules, Excel's own file changed and saved again, and a live gate; conditional formats read from a workbook without a VBA project join where Excel joins them -- rules alike, of a kind a formula decides, whose areas make one block, unless a rule between them covers the area -- and are saved again in Excel's order, over 84 live layouts in three file formats, Excel's files saved again, and a live gate; a rule's number format reads back, saves and is read again as Excel's does -- a built-in format under its place in Excel's own list of them, a custom one under the id the cells' formats share and numbered again after theirs as a file opens, whatever id the file gave it, and every literal escaped -- over 88 live layouts, every printable character, and 11 files spelling and numbering rules' formats as other programs might; a file's own number formats take the ids Excel gives them as it opens the file -- from 164 in the order the file lists them, one code one id, one nothing uses included, and a built-in id over another code or an id below 164 made custom -- and a save writes them under those ids, spelled as Excel spells them and without those nothing names, pivot tables' ids following, though no macro changed the stylesheet, over 11 files Excel opened and saved unchanged and a live gate; a save of a stylesheet nothing changed writes what is in use and nothing else, as Excel's save does -- an xf nothing uses left out with the parts only it names, an xf alike to an earlier one joined to it whatever its flags, a pivot table's button kept apart, a built-in style nothing uses left out, and every cell, row and column pointing where its xf went -- over 12 files Excel saved unchanged and a live gate; and a save writes the dxfs as Excel does -- a new one ahead of the file's, each name of one its own, in the order Excel reads them, those nothing names left out, and every rule, sort, filter, table, table style and pivot table naming its dxf where the save wrote it, a rule in the worksheet's Excel 2010 extension numbering its format with the rest -- over 13 files Excel saved unchanged, three it edited and a live gate; a rule a macro takes away by a Delete, ClearFormats, Clear or a paste leaves its dxf in every save until the workbook is opened again, and one that goes with its rows or its sheet takes its dxf with it, over 16 files Excel saved and a live gate; and a FormatCondition's Delete through a range holding part of its rule takes that part out of the rule, which keeps its priority and its dxf, over 22 live layouts and a live gate; manual page breaks -- `HPageBreaks` and `VPageBreaks` with `Add`, `Count`, `Item`, For Each and each break's `Location`, `Type`, `Extent` and `Delete`, `Range.PageBreak` read on whole rows and columns and set on them or a cell's row and column, `ResetAllPageBreaks`, a collection counting the breaks before the lines the sheet uses and the one after, breaks moving with rows and columns put in and going with the line after them, kept by a copied sheet, saved as rowBreaks and colBreaks and read back -- as Excel does over 80 live cases, Excel's file read back and a live gate | Automatic page breaks, which Excel works out only as it lays pages out, and breaks on a sheet with a print area, which Excel cuts to it and edges with automatic ones and which report themselves unsupported; `DragOff`; Normal's font size, face, bold, italic and underline beyond Aptos Narrow 11 and Calibri 11, which the sheets' geometry rests on; a style of a workbook that is not the active one, which Excel changes in the active workbook; conditional formatting and page setup; a validation list read from another sheet, which Excel keeps in the worksheet's x14 extension; conditional formats in that extension other than data bars, icon sets and colour scales, such as a cell rule naming another sheet, which leave the sheet's extension as the file had it and data bars and icon sets there unsupported; a rule formula's relative references, which Excel wraps round the sheet's edges and the model makes #REF!; a font nothing uses at all, which Excel leaves out as it saves, where the model keeps every font no xf names, as Excel keeps those a form control's label or a note's text is in; the colours of a colour scale whose points are out of order; cell shifts (`Insert`/`Delete` with `Shift`) on sheets with row or column formats; rows mixing fonts from different tables, unmeasured fonts, super/subscript and wrapped or turned text; a thick border along a hidden row; column AutoFit over text that wraps, is indented or turned, in fonts other than Aptos Narrow and Calibri, below 8pt, where Excel draws both in another font, or past 72pt, or holding Windows-1252's characters past Latin-1, whose GDI widths depend on their neighbours, and the widening a number format gives a column for such text, which leaves the column as it is; the width a number General shows widens its column to when General or a font size is set on it, which is not AutoFit's for its whole text (85 measured cases); Normal fonts other than Aptos Narrow 11 and Calibri 11 |
| Tables and structured references | Power Query reads workbook tables and writes load targets; ListObjects reads a file's tables and makes new ones as Excel does -- headers, generated column names, the header row put in, the name and style -- with ListObject, ListColumns and ListRows answering, and saves them in the parts Excel writes, over 8 live workbooks and 15 live questions; a table read from a file keeps its part until it changes, and whole rows inserted or deleted through it carry it. Structured references are read, spelled, saved and worked out as Excel does over 80 live formulas and a 50-name character sweep: this row, special items, spans, escapes, the table's name left out inside it, a column read as this row's cell where one value is wanted, Range and Evaluate, AutoFill along the columns, and renaming a table or its columns through every formula. Renaming a sheet in a workbook that holds a table reaches the table's calculated columns and totals as it reaches cells and names, in the table part a save writes too, over 7 live cases and a workbook Excel saved and opened again. The totals row shows and hides as Excel does over 20 live workbooks, with its default label and function, every TotalsCalculation, and labels, SUBTOTALs and formulas written into it. ListRows.Add and Delete, ListColumns.Add and Delete, Resize and Delete, calculated columns, and a table growing over what a macro writes beside it, with the references it stretches, match 28 live workbooks and 21 live writes. A table filters by a filter of its own, which a range inside it reaches, its field counted from the table's first column: ListObject.AutoFilter, ShowAutoFilter, the criteria saved in the table's part and read back, a table made over the sheet's filtered range taking that filter over, and the sheet's AutoFilter, ShowAllData, AutoFilterMode and FilterMode reaching the filter at the active cell as Excel's do, over 76 live layouts, Excel's file read back, 11 steps on it opened again and a live gate | Unlist, which turns the table style into cell formats; a copy or fill beside a table, which puts a table row in first; whole columns inserted or deleted through a table; the dxf a calculated column made from one cell brings; and a header holding a carriage return, which VBA reads as a line feed; a filter on a table with no header row, rows put in or taken out inside a filtered table, a column it is filtered by deleted and a filtered table resized wider or longer, which report themselves unsupported, and the longer AutoFilter.Range Excel reads, once its file is opened again, for a table resized smaller over data left below it |
| Formulas and calculation | Cells are worked out by pyOfficeEditor's engine, 501 of Excel's 525 functions, which comes to what Excel cached for all 10,958 formulas of its corpus over the model's workbook; formula parser, dependency invalidation and manual mode exist; a formula's last sum set to zero when it cancels, comparisons to fifteen digits and lookups to the bit follow Excel over 1,788 live pairs of doubles; defined names that stand for a formula work out where they are used, as array formulas; a range or array where one value is wanted is cut to the formula's row or column or run item by item, SUMPRODUCT, INDEX, ROWS and COLUMNS work their arguments out as arrays, and a value where a function reads cells is refused, as 232 live formulas show and as each of the engine's 493 functions measured there does, 15 of them taken; a call with too few or too many arguments, a LET's or a LAMBDA's name that is none, and a formula past Excel's limits on text, brackets, calls, signs and length are refused, over 14,377 live formulas, every function with 25 counts of arguments; a running total 3000 rows down and the deepest formulas Excel takes work out as Excel does; a formula written in the other notation, `Formula = "=R[-1]C+1"` or `FormulaR1C1 = "=A1"`, reads as Excel reads it, over 79 live formulas; LET's and LAMBDA's names, those looking like cells among them, and LAMBDA's optional parameters read, work out and save as Excel's do, over 24 live formulas; a formula written to a General cell takes the number format Excel gives it, over 111 live cases and every one of the engine's 501 functions written alone, the financial functions' currency and percent formats among them; `FormulaArray`, `HasArray` and `CurrentArray` follow Excel over 78 live cases, array formulas saved and read as Excel does; the range operator, intersections, unions and INDEX, OFFSET, INDIRECT, CHOOSE, IF, IFS, SWITCH and XLOOKUP as references follow Excel over 45 live formulas, and INDEX with one index into rows and columns and a `+` before a reference over 26 more; `Formula2` and `Formula2R1C1` read a legacy formula with Excel's `@` over 2100 live formulas, every function in every way it takes an argument; an `@` or a call of SINGLE written through `Formula`, `Formula2`, `FormulaR1C1` or `Value` is kept, left out, read back and saved as Excel does, over 274 live writes and a saved workbook; dynamic arrays spill, grow, shrink and are #SPILL! where something is in the way, with `HasSpill`, `SpillParent`, `SpillingToRange`, `A1#` and their file format, over 42 live cases, 18 expressions and two saved workbooks; a formula calls the project's VBA Functions, user-defined functions, as Excel does over 51 live cases: their arguments, results, errors, Application.Caller, recalculation and the Err they leave; Application.Volatile registration and opt-out, cached reads, dependents, manual/automatic calculation and workbook ownership follow 24 more native cases; 35 native spill-edit cases cover row/column and partial-cell shifts, refused sorts, cuts and copies, with save/reopen regressions; 110 more native probes cover manual caches and save settings, with nine native files and native readback of model-written caches; 189 CSE copy/cut/shift/sort/PasteSpecial/conversion cases and fourteen native file gates cover array fragments and cached states; another 155 CSE fill/replacement/duplicate-removal probes, including eight event traces, and thirteen native files cover repeated/clipped blocks, anchor replacement and array data-row guards; 88 UDF source-lifetime probes and six macro-enabled files cover edited bindings, project-global resets, invalidation, cached text and persisted dirty-formula metadata; 404 full/rebuild, formula-entry and information-intrinsic probes, 390 repeated scheduling observations and six more native files cover explicit full calculation, volatile caches/dependents, learned order, provisional UDF retries and initial manual values; 192 scope/deferred-event probes, 1,080 additional UDF scheduling observations and four XLSM files cover limited dependencies, cached foreign precedents, per-cell dirty flags and notification delivery; thirteen native post-reopen CSE traces distinguish clipped runtime caches from their saved state; the engine is pyOfficeEditor's to its commit 4723383, the tests pyOfficeEditor wrote for the commits merged running here (`tests/test_formula_engine_sync.py`): ODDFPRICE, ODDFYIELD, ODDLPRICE and ODDLYIELD match 194 cells measured on their own to the bit but for nine long-first-period yields within 1.6e-11, with GROUPBY, PIVOTBY, BAHTTEXT and functions passed by name; an error where a function takes only a range is handed on, as 88 live readings show; a formula reading a closed linked workbook works out from the copy of its cells and names the file keeps, shows the link as the workbook's path and takes a macro's path or file name for it, over 19 live formulas through a calculation, a full one and a row put in, and pyOfficeEditor's 25; a 3D reference, `Jan:Mar!A1`, is #REF! alone, read on every sheet by the 62 functions that do so, #REF! to the 41 that take an array and #VALUE! to every other function that takes a range, and is written back first sheet to last, over 277 live formulas, every function that takes a range measured; a sheet's name in a formula is quoted as Excel quotes it, over 100 live names and pyOfficeEditor's tables of the characters past ASCII it quotes for; a circular reference is found as Excel finds it, through the cells a formula reads as it is worked out, so ROWS, an IF's branch not taken or a MATCH stopping at its match closes none: with Application.Iteration off the formula closing a loop is 0 and the rest keep what they hold, and with it on every calculation works loops out over and over to MaxIterations or MaxChange, a sheet at a time in the order the sheets were made and row by row, with Worksheet.CircularReference, the settings saved on calcPr, over 93 live probes, 24 repeated runs and three saved workbooks; what-if data tables made by Range.Table, with one input cell or two, work each cell out as Excel does, the value it tries put in the input cell and every formula reading the input, on any sheet, worked out again and then put back, in automatic, semi-automatic and manual calculation, their cells saying =TABLE(), guarded, moved and copied as Excel guards, moves and copies them, every refusal and its message, saved as Excel saves them, over 44 live probes, three workbooks Excel saved and three the model saved that Excel opened | Complete function and argument coverage, broader UDF source-edit/class-instance lifetime and general calculation-chain scheduling, broader spill/structural-edit interactions with tables, merges, filters and dependencies, broader CSE array editing combinations, relative R1C1 in INDIRECT and unions for functions beyond the aggregates, errors, cached-value persistence and names with relative references; bond functions and basis-1 YEARFRAC at the start of 1900 (16 named cells, `tests/test_formula_early_1900_dates.py`); a LET-bound array in a range-only parameter of a Formula2 formula, which Excel makes dynamic; a formula naming a workbook the file has no link to, which would make a new link, and updating a link from its workbook; Excel's reading of a quoted 3D pair in brackets of their own as another workbook's; with Application.Iteration off, a loop's cell an earlier edit left to work out again and nothing read before a later formula closed the loop keeps its value from before that edit, as the model works formulas out when they are read; fills, sorts, Replace, cuts, merges, TextToColumns and pastes onto a data table, and one on a protected sheet or over merged cells, a spill or a table, which report themselves unsupported |
| VBA WorksheetFunction and Evaluate | WorksheetFunction.X and the late-bound Application.X call the cell formula engine for the 393 WorksheetFunction members it has, with arguments, answers and errors converted as 160 live calls made both ways show, and a range of another open workbook read as that workbook's cells, whichever is active, over 11 more; Evaluate, `[...]` and Worksheet.Evaluate work out expressions through the same engine as Excel does over 91 live cases: a Range for cells, blocks whole, error values, ROW() as an array of one; in the R1C1 style Evaluate reads R1C1 references, and in the A1 style it refuses them (`tests/fixtures/reference_style.json`) | The WorksheetFunction members the engine lacks: the cube functions, WEBSERVICE, STOCKHISTORY and the like; the members registered with the engine checked through VBA one by one |
| Events | Document modules hear Change, Calculate, SelectionChange, Activate, Deactivate and NewSheet, over 35 live actions and 7 live reach cases. WithEvents custom events have 15 live cases; worksheet, workbook and application sinks for existing sources have 14 isolated live cases. BeforeSave, AfterSave and BeforeClose have 22 live cases for cancellation, suppression and saving order; Workbook_Open and workbook activation have four native traces and persisted execution gates | complete event signature validation; remaining sheet and application events, such as BeforeDoubleClick and FollowHyperlink; errors inside handlers |
| Shapes and controls | Public SheetView snapshots and add/edit/delete APIs cover measured AutoShapes, text boxes and nine Forms control types; control edits reach supporting parts; checkbox/list/numeric/radio control values, numeric bounds/increments, A1 and workbook/local named sources/links, any formula that lands on cells worked out by the cell engine (CHOOSE, IF, IFS, SWITCH, INDEX, OFFSET, INDIRECT in A1 and R1C1, LET, XLOOKUP and the reference operators, over 73 live cases), inline items, whole-list arrays, range-source conversion and multi/extended selections and rebinding replay live measurements and pass persistence gates | External control links, names with relative references inside a control's formula, broader named-reference cases, grouping, copying, pictures, geometry, formatting, anchors and macro links |
| Charts | Embedded ChartObjects, Chart, SeriesCollection, Series and ChartTitle share drawing/cell state; 32 native API probes cover creation, lookup, geometry, titles, series and row/column source inference, with two window-placement gaps; eight model-written chart kinds open in native Excel; parts, relationships, reference caches and no-op chart XML have roundtrip tests; writer adapted from pyOfficeEditor at 4723383e5d3039763e64ecd2a3ee092786c91bb5 | Imported chart-content mutation, richer source inference and reference expressions, axes, formatting, error/lifetime behavior, chart sheets, copying/structural edits and window-dependent automatic placement remain incomplete; a workbook holding a chart sheet reports itself unsupported when it is opened |
| Pivots, caches and data model | Imported OOXML PivotTables and PivotCaches expose collections, source/cache metadata, field/axis collections, cached items and report ranges. Worksheet-source PivotCaches.Create and CreatePivotTable support pending caches, formatted/duplicate headers, shared snapshots and empty reports, with new cache/record/report parts persisted across saves. AddDataField supports all eleven aggregates over the cached snapshot, with compact row/column/page layouts, nested subtotals and multiple measures. Orientation supports field placement/removal and Values down rows or across columns. Explicit AddDataField removes its source axis while Orientation=xlDataField preserves it. Retained Values fields track moves; removed data-field references return native VBA errors. Native controls cover error/blank/date/numerical behavior, sorting, typed axis labels, saved reports and native recalculation over generated caches. GETPIVOTDATA reads displayed results across compact/tabular/outline layouts, multiple axes/data fields, subtotals, page filters, renamed items and manual/numeric/date groups; native controls cover error precedence, reference arguments, defined names, Evaluate, CSE/dynamic arrays and formula text. VBA PivotTable.GetPivotData returns the selected cell as a live Range, with native optional/named arguments and error 1004. Report metadata is shared across fresh object proxies without caching result values. Native fixtures cover shared caches; unchanged pivot parts survive unrelated cell edits and saving. Rewritten imported report/cache XML retains scoped namespace declarations and compatibility-attribute bindings so native Excel accepts supported edits. Field data/label ranges support compact/tabular/outline rows, multiple axes and measures, Values on either axis, grouped imported fields and multi-area selections. PageRange exposes vertical and wrapped filter blocks in both fill orders; TableRange2 includes their full width. Supported edits retain imported wrapping and clear obsolete filter cells. Native controls cover 32 generated layouts and 23 imported workbooks, including save/reopen and range-cell enumeration. Reports filter by their items as Excel's do: `PivotItem.Visible` hides an item of a row, column or page field and its records count nowhere, `CurrentPage` chooses a page field's item, `EnableMultiplePageItems` lets it show several, `ClearAllFilters` shows every item again and `VisibleItems` and `HiddenItems` list them, a refresh keeping hidden items by their values, marking those the source lost and hiding new ones where a field hides any, over 65 live layouts, Excel's saved workbook read back and a live gate; Excel cannot open its own file of a page field that chooses an item and allows several, which the model saves as choosing the item. Reports sort as Excel's do: `AutoSort` by a field's names or a data field's values, inside each outer item, ties in the field's order, `AutoSortOrder` and `AutoSortField` reading the sort back, `PivotItem.Position` moving an item among those showing, a refresh sorting new items in, and the table's row and column lines written as Excel writes them, over 45 live layouts and a live gate. Label, value and top 10 filters follow Excel too: `PivotFilters.Add` by an item's caption, its * and ? wild, or by a data field's totals inside each outer item, top 10 by count, percent or sum, one label and one value filter a field beside its hidden items, `PivotFilter` read back and deleted, `ClearLabelFilters`, `ClearValueFilters` and `ClearManualFilter`, over 73 live layouts and a live gate. Date fields and date filters follow Excel: items named and labelled in their field's format, the format every field and data field takes from its source's cells, a date filter by dates, by a period counted from today or by a quarter or month of any year, `Add2`'s WholeDayFilter, and a refresh dropping a lost date, over 79 live layouts and a live gate that counts the periods from the day it runs. Layouts follow Excel too: compact, outline and tabular forms by `RowAxisLayout` or a field's `LayoutForm` and `LayoutCompactRow`, `Subtotals` with every function, the subtotal location, blank lines, repeated labels, `SubtotalName`, the table's captions, grand totals and indent, a data field's `Function`, and every pivot cell in the format Excel gives it, over 122 live layouts and a live gate. The table's options follow Excel as well: `TableStyle2` and its style flags, the null and error strings in the value cells, the page fields' order and wrap, `ShowValuesRow`, the tag, alternative text and flags, `AddFields`, `ClearTable`, `Location`, `Update`, `VisibleFields`, `ManualUpdate` and `SaveData`, over 61 live layouts and a live gate. Show Values As follows Excel: `Calculation` with all fifteen calculations, `BaseField` and `BaseItem` by name, `(previous)` and `(next)` stepping over the items a group lacks, every empty, zero and error cell, subtotals by another function, a data field's `NumberFormat`, Excel's formats for a calculation, and a value sort giving way along a field a calculation steps along, over 117 live layouts and a live gate. Calculated fields follow Excel: `CalculatedFields.Add`, each cell's formula over the sums of the fields it names with any worksheet function, `Formula` read and set and spelled back as Excel spells it, `Delete`, `Caption`, `IsCalculated` and `DataType`, a refresh keeping them, and value sorts, filters and Show Values As over their values, over 59 live layouts and a live gate. Names follow Excel: a data field's, the Values field's, a field's and an item's `Caption` and `Name`, with Excel's errors and its own quirks, an item showing, sorting and found by its caption, over 42 live layouts and a live gate. Items collapse and expand as Excel's do: `PivotItem.ShowDetail` and `PivotField.ShowDetail` on the rows and columns in every form and subtotal place, with the Values field inside a collapsed item, over 24 live layouts and a live gate. Grouping follows Excel: `Range.Group` by ranges of numbers with Excel's default step and labels, by a date's periods into fields of their own, and chosen items into named groups, grouped again; `Range.Ungroup`; a refresh working automatic ranges out again; another table of the cache showing the grouping; `PivotItem.LabelRange` and `DataRange`; and `ParentField`, `ChildField`, `GroupLevel`, `TotalLevels`, `ParentItems`, `ParentItem` and `ChildItems`, over 134 live layouts and a live gate; workbooks Excel grouped read, lay out and save as Excel's do. Calculated items follow Excel: `CalculatedItems.Add` with formulas over a field's items by name, in apostrophes, as Field[Item] or by place, each cell worked out from the items it names there, hidden ones too, totals over the field taking each item's value by the data field's function, sorting, hiding, renaming, `Formula`, `Delete`, `IsCalculated` and a refresh, with Excel's errors, over 70 live layouts and a live gate. `ShowPages` follows Excel: a sheet for each item a page field shows, hidden and chosen items left out, made from the last item to the first, each before the sheet active then, a copy of the table at the top left of each choosing its item alone, the sheet named by its item cut to 26 characters with ` (2)` and on where another sheet has the name, or left with its `Sheet` name where no sheet can take the item's, the events once every sheet is made, a table made without a name taking the next number of one count through the session, page fields with no room above a table moving it down, and a page field's number or date aligned left, over 35 live scenarios and a live gate. A table's columns fit its cells each time it is laid out, as Excel's Autofit column widths on update does: narrower too, over a width a macro set, the table's cells alone, in the bold each built-in pivot or table style draws there less what its flags turn off, 16 pixels for a dropdown, the +/- buttons' 88/7 pixels and each depth's step in a compact label column, a subtotal's label at its plain indent, an outer row field's column in outline or tabular form kept, each run of item columns a total column closes as wide as its widest, a new style fitting again and the style's flags and the buttons waiting for the next layout, over 1,280 measured labels, 44 edited tables and every pivot sheet of the earlier fixtures. A cache saves as Excel saves it: shared items listed for the fields a table lists items for and the other fields' values in the records, each sharedItems with the attributes that differ from the schema's defaults in its order, numbers spelled as Excel spells them, every one of 2,477 fields Excel saved in the fixtures following that rule and the caches of all 715 tables the fixtures replay, fields and records, Excel's own. Tables sharing a cache refresh together: `RefreshTable` or `PivotCache.Refresh` lays out every table of the cache on every sheet, sheet by sheet with each sheet's newest first, `Workbook.RefreshAll` refreshes every cache a table uses from the last to the first, passing over one whose refresh is off, and `ClearTable` lays the cache's tables out again, the refresh ending with Excel's calculation and its Calculate events in the order the sheets were made, over 18 live scenarios, 19 event cases and a live gate. `Range.PivotCell`, `PivotTable`, `PivotField` and `PivotItem` say what each cell of a table stands for, its type, field, item, data field, row and column items and custom subtotal, every cell of 595 tables Excel saved agreeing. | `PivotCell.PivotRowLine` and `PivotColumnLine`, unmeasured because reading a row label's column line stops Excel; PivotCell on a table of a cube; `ShowPages` with a name no field has, which stops Excel itself, with another table's field, or on a field listing items its source lost; a sheet's pivot tables as Excel lists them after opening a file, which turns round from one open to the next, where the model lists them by their relationships' ids; a table's columns fitted in a style the workbook defines, whose bold the model does not read, in a font other than the Normal font AutoFit measures, or over text past Latin-1, which keep their width; a collapsed item's blank line after it, unmeasured, which the model keeps out as it keeps out the item's subtotal; `Range.ShowDetail` on a pivot table's cell, which drills down to a new sheet; a calculated field's formula Excel refuses while alerts are on, where Excel shows a message; grouping the cells of one item, numbers from a Start to the same End, ranges over a field whose items are grouped, the items of a field grouped by ranges or periods, ungrouping a group that is grouped again, and ungrouping a field a data field summarizes, which the model reports unsupported; a refresh losing an item a group holds, unmeasured, where the model keeps the item's place; calculated items in a grouped field, in more than one field of a table, or beside Show Values As or a label or value filter, and grouping a field with calculated items into groups of its items or by dates, which Excel does and the model reports unsupported; Show Values As beside the innermost field's own subtotals by another function, `NumberFormat` set on a field that is not a data field, and a `PivotField` object assigned to `BaseField`, error 1004 in Excel and 438 in the model, which reads the object's default member first; `PivotFilters.Add`'s Order, Name and Description, a period filter's days counted again on a later day, and filtered fields moved to the page area; `MergeLabels` set True, which merges and centres the label cells; a table put in no style, which Excel draws with borders of its cells; `SaveData` off on a cache a file holds; `IncludeNewItemsInFilter` and page items that are numbers or dates; sorting by a `PivotLine` or a custom subtotal and moving a page field's items; refresh, arbitrary imported-layout edits, external/consolidated cache creation, OLAP/data-model operations, page-layout mutation and exhaustive mixed-layout field-range fidelity, formatted/date runtime PivotItem captions, pivot cell edit/caption-write rules, and pivot-aware structural edits/copying remain incomplete. Creation validation and nonstandard source/destination shapes are not exhaustively measured. Complex user-defined/name expression shapes and nonstandard report layouts are not exhaustively measured. Cached reports do not refresh when source cells change. |
| Power Query | Packaging, M evaluation, local refresh and sheet load targets exist | Complete language/library conformance, data types, errors, query dependency behavior, connectors and refresh settings |
| Other workbook features | Coverage varies and needs classification; notes, which VBA calls comments, are added, read, edited, moved with their cells, copied and saved as Excel does, over 89 live probes and a saved workbook of every kind of note, its parts written byte for byte. Hyperlinks are added, read, set, taken away, moved, copied, pasted, cut, sorted and saved as Excel does, with the Hyperlink style and the HYPERLINK function's, over 26 sections of live Excel and 21 workbooks it saved, their links, relationships, cell styles and stylesheets matched, and a live gate opens the model's in Excel, which reads them as its own; a part's relationships come in Excel's order, from the hash Office places their ids by, over every count to 160 and 17 larger ones up to 65,530. Document properties -- `BuiltinDocumentProperties` and `CustomDocumentProperties`, every built-in property's type, default and coercion, custom properties added, changed, renamed, deleted and linked to named ranges, Excel's errors, dates kept to the second in UTC by the machine's time zone year by year -- follow 38 live cases, and a save writes docProps core.xml, app.xml and custom.xml byte for byte, its sheets and named ranges listed as Excel lists them, a live gate reading the model's files as Excel's own | Inventory names, links, threaded comments, protection, embedded objects, connections and remaining reference/file features; a note's box a macro moves or resizes, and notes under Sort and RemoveDuplicates; a hyperlink on a shape, EmailSubject on a link with no address, Follow, a paste of links turned or from several areas, and the relationship order of a part with more than 65,536; the [Content_Types] order Excel writes for parts other than docProps |
| Fidelity and release verification | Offline tests, live gates, language matrix and cross-platform CI exist | Add real-world mixed-feature cases, malformed/foreign-author packages and scale tests; require live evidence for new persisted behavior |

## Execution order within Excel

1. Extend the public shapes and form-control API required by the resolved
   ownership issue. The first slice exposes snapshots and common
   creation/edit/delete operations; control settings and remaining shape
   operations still need implementation and conformance gates.
2. Expand the operation-level inventory for the core workbook, worksheet
   and range model, including persistence and formatting. Implement
   missing foundations before the features that depend on them.
3. Complete calculation and its VBA bridges, then event dispatch, with
   measured error and mutation behavior.
4. Complete the remaining tables, shapes, charts, pivots, Power Query,
   formats and other inventoried headless features in dependency order.
5. Run the complete offline suite and live conformance gates, resolve
   every remaining in-scope entry, and publish the supported scope with
   the evidence before starting Word's completion milestone.

## Later hosts

Word and PowerPoint get separate inventories and measurements once
Excel meets its completion gates. They keep the same evidence standard
but use their own object models and file behavior. Word's text, tables,
sections, headers, fields and shapes and PowerPoint's slides, masters,
layouts, text, media, tables, charts and shapes all need explicit
coverage. Their existing support continues to receive regression fixes.
