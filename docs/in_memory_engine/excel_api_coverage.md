# Excel API coverage

Snapshot: **20 October 2026**, covering the current source tree, including
local AdvancedFilter, bounded AutoFilter wildcard matching, staged local Pivot
refresh, worksheet page configuration, table-owned sorting, and ordinary conditional formatting.
This is a functionality summary for the in-memory Excel engine.
**No functional area is complete yet.** Common worksheet operations have
substantially more support than advanced Excel automation.

The table consolidates **389 Excel object and event types into 22 functional
areas**. The detailed operation statuses and evidence remain in the
[Excel member checklist](../excel_checklist.csv).

## Status and counts

**Partial** means useful behavior exists, with remaining implementation or
conformance gaps. **Missing** means the area's VBA object API is not implemented.
Preserving an existing workbook part does not establish support for manipulating
that feature through VBA.

**Registered** counts reference members with a registered entry point, including
partial implementations, unaudited behavior and explicit unsupported paths.
**Verified** counts members marked `VERIFIED` in the checklist: their arguments,
errors and persisted effects have completed the operation audit. Neither count
is a percentage of working macros or complete functionality.

| Functionality | Status | Registered / reference members | Verified members | Supported behavior and main gaps |
| --- | --- | ---: | ---: | --- |
| Application, workbooks and worksheets | Partial | 127 / 752 | 6 | Core containers, opening/saving, sheet operations and settings. Adding worksheets or a first VBA project preserves prefixed OPC relationship/content-type metadata. Many application services, workbook features and host-specific behaviors remain. |
| Cells and ranges | Partial | 118 / 215 | 29 | Reads/writes, formulas, references, copying, moving and many editing operations. Native probes cover Currency precision and SpreadsheetML text/formula persistence for literal escapes, controls and Unicode. Full argument, error, structural-edit and mixed-feature fidelity remains incomplete. |
| Formulas and calculation | Partial | 388 / 463 | 0 | Broad worksheet-function exposure, dependency calculation, UDFs, arrays and calculation modes. Registered functions still need behavioral audits; remaining functions and Excel numerical edge cases are incomplete. |
| Defined names | Partial | 15 / 30 | 0 | Workbook/sheet names and reference evaluation. Full expression, relative-reference, external-reference and lifecycle behavior remains incomplete. |
| Cell formatting and styles | Partial | 103 / 206 | 20 | Fonts, fills, borders, number formats, alignment and named styles. Ordinary conditional DisplayFormat views are supported. Rich text, gradients, advanced display formats and remaining formatting APIs need work. |
| Conditional formatting | Partial | 77 / 212 | 0 | Cell-value/expression, top/bottom, above/below average/deviation and duplicate/unique rules, with specialized objects, lifecycle, differential styles, displayed formatting and conditional number formats. Text, blank, error and time-period predicates also support generated formulas, styles and worksheet copying. Measured disjoint range copies and all/formats pastes support fresh rules and loaded rules with distinct validated group identities within one workbook. Loaded group ordering, plain-load mode, shared-child deletion/range changes and ordinary formula rebasing have native controls. Measured worksheet copies initialize independent dense group identities and child order. Shared-group range transfers, broader worksheet-copy/context behavior, overlapping, tiled, transposed and cross-workbook transfers, structural edits, color scales, bars, icon sets and pivot scopes remain incomplete. |
| Excel tables | Partial | 46 / 155 | 25 | ListObjects, rows/columns, structured references, resizing, totals and ordinary local ListObject.Sort. Filtered/rich table sorting, configured structural edits, table-style objects and broader table-operation fidelity remain incomplete. |
| Sorting, filtering and outlines | Partial | 38 / 66 | 5 | Common worksheet/table sorts, direct sort-owned range parents, persisted ordinary worksheet/table settings, AutoFilter wildcard matching with bounded state transitions and a measured local AdvancedFilter subset. Ordinary row/column outlines support grouped levels, summary sides, detail visibility and persisted metadata. Formula criteria, advanced-filter state migration, sort variants, automatic outline styles, heterogeneous outline reads and broader grouping/editing interactions remain incomplete. |
| Data validation | Partial | 19 / 21 | 14 | Validation rules, formulas, input/error messages and persistence. Remaining members and full rule/error behavior need auditing. |
| Protection and permissions | Partial | 11 / 32 | 0 | Worksheet/workbook protection and enforcement for supported edits. Allow-edit ranges, user permissions and full protection semantics remain. |
| Comments and hyperlinks | Partial | 35 / 74 | 25 | Classic notes and cell hyperlinks, including many edit/move/save paths. Threaded comments, richer note geometry and remaining hyperlink behavior are incomplete. |
| Shapes and worksheet controls | Partial | 54 / 2,025 | 0 | Basic shapes/text boxes and nine Forms control types, with measured values, links and persistence. Rich drawing APIs, grouping/copying, pictures and newer controls remain incomplete. |
| Charts and sparklines | Partial | 32 / 951 | 0 | Embedded charts, basic series/titles, ordinary series deletion/reordering with held-series identity, and eight persisted chart kinds. Axes, richer formatting, imported-content mutation, chart sheets and sparklines remain incomplete. |
| PivotTables and slicers | Partial | 66 / 637 | 0 | Imported report queries, GETPIVOTDATA, cache metadata, empty-report creation and data-only AddDataField totals using all eleven aggregate functions, with measured events, field ranges and saved-report edits. Native probes cover date/time serials, Currency, typed values, errors, numerical edge cases and immutable cache snapshots across saves. Compact row/column/page layouts, nested subtotals, Values-axis placement and field removal now aggregate cached records. Field/page range queries support native measured compact, tabular, outline, grouped and multi-area reports; sparse layouts avoid scanning all records for each blank intersection. Fixed-schema local cache refresh stages all linked supported same-sheet reports before publishing callbacks. Broader refresh, arbitrary layout edits, filtering, grouping and slicers remain incomplete; general linked rich-data creation is unsupported. |
| Queries and external data | Partial | 8 / 278 | 0 | Workbook queries plus local Power Query evaluation, refresh and sheet loading. Connection objects, remote connectors, QueryTables and full query semantics remain incomplete. |
| Data model and OLAP | Missing | 0 / 261 | 0 | No in-memory Power Pivot model, relationships/measures or cube automation. Reading some external-pivot metadata does not implement these operations. |
| XML mapping | Missing | 0 / 75 | 0 | No XmlMaps, schema/namespace binding or mapped-data import/export API. |
| Excel events | Partial | 0 / 123 | 0 | Change, calculation, selection, activation, lifecycle and custom WithEvents paths have native measurements. Remaining events and full signature/dispatch behavior need work. |
| Windows, views and dialogs | Partial | 38 / 304 | 23 | Selected window, view, freeze/split and visibility state. Full interactive UI, dialogs, protected views and additional view APIs remain incomplete. |
| Printing and publishing | Partial | 27 / 152 | 0 | Persisted worksheet PageSetup margins, scaling, initial paper/layout enums, print options and scoped print areas/titles. Printer services, header/footer editing, pagination, rendering and publishing remain incomplete. |
| Add-ins and embedded automation | Missing | 0 / 124 | 0 | No full Excel AddIns, embedded OLE/ActiveX or external automation-host implementation. |
| Legacy features and utilities | Missing | 0 / 386 | 0 | Legacy menus/toolbars, dialog/module-sheet utilities, SmartTags, scenarios and assorted web/research services remain. |
| **Total** | **Incomplete** | **1,202 / 7,542** | **147** | **15.9% of reference members have registrations; 2.0% are checklist-verified. These are API inventory ratios, not compatibility scores.** |

â€  Excel events use the interpreter's event dispatcher and sinks rather than
ordinary registered object members. The registration counter therefore does not
measure the event behavior already implemented. Event status comes from the
native traces described in the [host coverage plan](../host_completeness.md).

## Integrated validation

The filtering, Pivot refresh and page-configuration increments pass **582 focused
integration tests**, with 28 optional native tests skipped in the normal run.
A separate final gate opened **28 engine-generated output workbooks in one owned
native Excel session**: ten AdvancedFilter, seven Pivot refresh and eleven page
configuration/name-lifecycle cases. Every result matched its existing independent
native oracle; no recorded expectations were normalized or changed. This validates
the admitted cases together, rather than claiming full API conformance.

## Worksheet outlines

`Worksheet.Outline` and ordinary whole-row/column `Group`, `Ungroup`,
`ClearOutline`, `OutlineLevel` and `ShowDetail` support measured level changes,
summary placement and `ShowLevels`. Depths range from one through eight; collapsing
or expanding updates visibility and dependent calculations. Ungrouping retains
hidden detail, while clearing outlines removes their visibility state. Plain
worksheet copies preserve these modeled settings.

The evidence retains **26 exact native controls** and **12 macro-free native
inputs**. Twelve runtime controls use explicitly documented observation subsets;
unsupported mixed-range getter observations remain intact in the raw evidence.
Two additional exact Variant/CallByName matrices establish fractional level
truncation, Boolean level assignment, separate summary-side coercions and
string/error behavior. Their model checks explicitly exclude only the native
version/build header; every operation/error/readback remains exact. Unmeasured
integer summary-enum values still fail before mutation. The original raw
`summary-sides` control separately establishes integer `SummaryColumn=3`
selecting the right-hand summary; an explicit regression retains that evidence.
Six input snapshots also have exact model import/save/readback checks. All twelve
generated-output controls passed native Excel readback in one closed session
(12/12, 16.05 seconds). Additional
regressions exercise metadata ownership, lossless row bodies and column attributes,
combined-axis atomicity, calculation, protection and scope admission. See the
[outline tests](../../tests/test_excel_outlines.py) and
[native evidence](../../tests/fixtures/worksheet_outlines/initial.json).

Outline-only changes patch owned metadata without rebuilding existing cell XML.
Unknown row bodies, column attributes and sibling configuration remain preserved.
The initial axis writer supports ordinary unprefixed SpreadsheetML row/column
metadata; configuration-only changes also preserve prefixed worksheet content.
Worksheet copying rejects unparsed qualified outline rows/columns before creating
a destination, including sheets without explicit `outlinePr` flags. Axis admission
also checks that loaded and actual direct row/column elements agree on
nondefault outline metadata in both directions.
Operations bound metadata work to 50,000 rows, 16,384 columns, 16 MiB XML and
250,000 existing cell records for row-budget checks. These are implementation
limits rather than Excel limits; unsupported cases fail before mutation.

Automatic style creation, cell/Pivot grouping and heterogeneous `OutlineLevel`
reads remain unsupported. Native mixed-range observations are not yet consistent
enough to define one getter contract. Affected structural edits and whole-axis
range copies or formatting pastes are declined atomically rather than losing
outline metadata. Non-whole-axis `ShowDetail` and broader protection/UDF, outline
style, structural and cross-feature behavior need further measurements.

## Advanced filtering

`Range.AdvancedFilter` supports a single local rectangular list, in-place
filtering or extraction, criteria joined by column AND and row OR, repeated
criteria headers and projected or reordered output columns. **61 independently
measured native controls** cover the admitted subset; **70 raw controls** retain
nine explicitly deferred diagnostics without treating them as passing behavior.
Ten native macro-free workbooks cover imports and forced rewrites. Ten
model-generated workbook snapshots were also reopened and checked in native
Excel; model save/reopen covers the same independently measured snapshots. Two additional raw native calculation controls are retained
byte-for-byte with JSON newline conversion disabled in Git, with hash-checked,
explicitly projected model replays that omit
only the unsupported `Application.CalculationState` member. They verify
SUBTOTAL/AGGREGATE and dependent values after filtering, calculation and
`ShowAllData` in automatic and manual calculation modes; a filtered automatic
calculation workbook also survives generated save/reopen.

Text criteria use measured case-insensitive prefixes, exact comparisons and
wildcards with tilde escapes. Numeric equality also matches numeric strings;
ordered numeric comparisons exclude them, and numeric inequality retains them.
Boolean/error criteria and actual blank versus formula-empty values have typed
controls. Unique compares the projected output population, preserving numeric,
text, Boolean and error distinctions. Extracted formulas become evaluated
constants while cell styles follow the source; strings beginning with `=` remain
literal values. Old extract contents are cleared through gaps while unused
cell formats survive.

In-place filtering has `FilterMode=True` without AutoFilter dropdowns. It
persists hidden rows, the list/criteria names and the worksheet flag;
The persisted flag is read from hardened namespace-aware direct-child metadata
without constructing a worksheet element tree. `ShowAllData` restores that list
while leaving unrelated hidden rows alone.
Deleting the owning worksheet removes its local filter and ordinary names,
including after `ShowAllData`; workbook names and other sheets' local names
remain scoped independently of which cells their formulas reference.
Renaming the owner stages its supported local filter and print names together
before changing any entry. Four native controls cover in-place and extracted
filters, retained names after `ShowAllData`, criteria-free Unique filtering,
quoted apostrophe names, and macrofree save/reopen. Owned `Criteria`, `Extract`,
`_FilterDatabase`, `Print_Area` and `Print_Titles` keep their reserved metadata
and local scope; unrelated ordinary names remain outside this migration.
See the [combined name lifecycle tests](../../tests/test_excel_owned_reserved_names.py)
and [native name evidence](../../tests/fixtures/owned_name_rename/native.json).
Worksheet UDF attempts return `False` without changing output or names.
Protected in-place controls succeed; protected extraction remains explicitly
unsupported because native Excel can partly write before raising an error.

Formula criteria/unmatched headers, unsafe overlaps, cross-sheet or cross-book
arguments, merges, tables, arrays/spills, conditional/validation rules and
annotations are outside this initial scope. Mixing AutoFilter with an active
AdvancedFilter, structural edits and cuts with active state fail before mutation.
Worksheet copying also refuses retained owned AdvancedFilter metadata after
extraction or `ShowAllData`, including criteria-free Unique-only filters. Non-ASCII comparisons, text ordering and broader argument,
event and protected-sheet interactions remain incomplete. Implementation work
bounds currently admit at most one million list, output or projected input
cells (also when Unique collapses the output), 4,096 criteria cells, four million
row/criterion comparisons, four million aggregate text characters scanned for
headers, criteria, comparisons and Unique keys, and 254-character criteria;
these are explicit support limits, not claims about Excel's limits. See the
[AdvancedFilter tests](../../tests/test_excel_advanced_filter.py) and
[native evidence](../../tests/fixtures/advanced_filter/native.json).

## Persisted worksheet page configuration

`Worksheet.PageSetup` exposes **24 configurable properties** plus its three
standard object getters. This headless configuration API stores margins, scaling,
initial paper/layout enums, print options, print areas and repeating titles.
It does not print, paginate, preview, export, select printers or emulate printer
communication batching. `Chart.PageSetup` remains unsupported.

**18 completed native controls** and **11 macro-free native workbooks** retain
independent runtime/import/save observations; **12 configuration/lifecycle
controls** are replayed in their admitted scope. Initial observations have an
explicit `ConfigurationState` projection removing only `Application.ActivePrinter`;
source/result hashes identify the unchanged private environment-bearing originals.
Follow-ups use that environment-free observation schema directly. Additional
coercion diagnostics are preserved without calling them passing full controls.
Model-generated save/reopen covers the same configuration snapshots. The opt-in
native generated-file gate passed all **11 configuration/lifecycle workbooks**,
including their complete environment-free configuration snapshots. Raw printer parts and
relationships survive edits.

Margins use points at the VBA boundary and inches in XML. Common finite,
nonnegative numeric margins and Empty-as-zero are supported. Negative margins
are explicitly unsupported before mutation: native Excel accepts and saves them,
but a measured negative-margin file could not be reopened. The invalid file is
not a claimed passing input. Null/Boolean/string margin dispatch differences
between typed setters and CallByName remain unsupported rather than normalized.

Zoom supports False or integer percentages 10 to 400; fit counts support False and
bounded integer counts. Orientation/Order support their two ordinary values;
PaperSize setters initially support Letter/A4. FirstPageNumber supports
xlAutomatic or nonnegative integers 0 to 32767, with measured Variant readback types.
Broader paper/number/coercion and printer-dependent defaults remain incomplete.
Boolean options cover centering, headings, gridlines, black-and-white and draft;
comment/error printing enum values map to their persisted XML attributes.

Local bounded A1 print-area unions and single whole-axis titles use owned,
visible scoped `Print_Area`/`Print_Titles` names. Each title setter preserves the
other axis; stored titles put columns before rows. Native controls cover quoted
renaming, moving before/after another sheet, deleting another sheet and removing
the owner. Namespace/quote-aware remapping also preserves unmodeled reserved
names without interpreting fake attributes inside opaque quoted values; malformed
scope admission occurs before package parts or relationships are added. General
local-name renaming and broad formula rewriting are not established by this work;
the supported filter/print reserved names use one atomic rename plan.

The measured LeftMargin and Orientation worksheet-UDF assignments return without
changing configuration; PrintArea assignments do update their owned name,
idempotently and without recursively calculating. Other worksheet-UDF setters
are explicitly unsupported. Protected-sheet controls cover those three members;
broader protection/event interactions remain incomplete. Removed-sheet PageSetup
proxies are unsupported before they can recreate orphan names.

Only main-namespace configuration elements are owned; no-namespace lookalikes
retain their bytes and do not alter configuration reads. Writers patch only owned
attributes and preserve QName scopes, nested sheetPr
metadata, printer relationships/opaque parts, unknown attributes and header/footer
bytes. Creating a pageSetup node for scaling or page numbering explicitly retains
the existing portrait default, while subsequent edits retain explicit orientation.
Imported pageSetup without an orientation reports the measured Double zero; the
unmeasured explicit `default` orientation token is unsupported.
The shared hardened OPC scanner supplies nested byte spans in the complete
original XML. Reading configuration caches only bounded metadata; repeated reads
avoid building a second worksheet tree. Worksheet.Copy of nondefault configured
sheets or sheets with print names is explicitly unsupported before destination
publication; default-equivalent copies retain their existing scope. Same-workbook
moves retain the worksheet and its configuration. Cross-workbook print-setting
transfer, header/footer editing and printing/publishing services need further work.
See the [configuration tests](../../tests/test_excel_page_setup.py),
[name lifecycle tests](../../tests/test_excel_page_setup_names.py) and
[native evidence](../../tests/fixtures/page_setup).

## Excel engine behavior beyond member counts

AutoFilter wildcard matching preserves the existing literal, tilde, Unicode case,
newline and cell-type behavior without introducing a criterion-length cap.
[Wildcard compatibility tests](../../tests/test_excel_autofilter_wildcards.py)
cover short-pattern equivalence, repeated-state work and isolated completion;
existing native runtime, file and edit fixtures remain the behavioral oracle.
This changes no registered members, checklist statuses or completion totals.

The shared VBA engine implements procedures, classes, arrays, error handling,
events and much of the runtime library. Compiler/type checking, reference and
object compatibility, lifetime behavior and remaining language edge cases are
still partial. See [VBA runtime support](../vba_runtime.md).

VBA string lengths, `Left`/`Right`/`Mid` boundaries and `AscW` now count UTF-16
units when worksheet text contains supplementary Unicode characters. Binary
comparisons and binary `StrComp` use the same unit order. Native
[cell-string controls](../../tests/fixtures/cell_strings/controls.json) cover
text and formula results before and after native saves and recalculation.
Broader Unicode search, replacement, assignment and collation conformance still
requires auditing; these controls do not establish complete string support.

Cell-formula evaluation and Power Query's M evaluator also have execution paths
beyond VBA object-member registrations. Formula registration does not prove full
function conformance, and local query refresh does not establish remote-connector
support. See the [host coverage plan](../host_completeness.md) and
[Power Query documentation](../power_query.md).

OOXML persistence preserves many untouched parts and writes supported edits back
to workbooks. Feature completeness also requires native Excel to accept the
generated workbook and reproduce the intended behavior. Tree and offset readers
share the same XML parser factory, which rejects DTDs and declared entities
directly, including UTF-16 input. Passing the current
tests does not prove that every combination of Excel features works.

Native fixtures cover imported `PivotField.Function` getters for all eleven
aggregate functions, including Count Numbers returning `11` rather than the
`xlCountNums` input constant. Native PivotTables enumeration can change across
repeated opens of the same workbook; name-based comparisons establish content
fidelity without claiming identical collection order. The model currently uses
deterministic relationship-ID order on import. See
[aggregate getter tests](../../tests/test_excel_pivot_functions.py) and their
[native measurement](../../tests/fixtures/pivots/functions.json).

Native compact PivotTable controls cover 23 layouts, nine mixed-item/sorting
cases and 23 field-orientation argument/move/removal cases. All 55 recorded
results match the model; the 45 successful reports also reopen in native Excel
and accept an additional native aggregate over their saved caches. This includes
nested row/column subtotals, multiple measures, Values down rows, page-inclusive
report ranges, month sorting, case-insensitive item merging and typed labels.
Cache snapshots remain independent of later source edits, including after saving.
These controls do not establish full filtering, grouping operations, refresh,
formatting or event-interaction fidelity. See the
[axis regressions](../../tests/test_excel_pivot_axes.py) and
[native replay script](../../scripts/measure_pivot_axes.py). Member registration
and verified totals are unchanged: the added setter belongs to an existing member.

Local `PivotCache.Refresh` and `PivotTable.RefreshTable` now stage a new source
snapshot and every linked supported same-sheet report before changing live cells,
cache records or field nodes. Successful refresh retains held cache/field
identities and package cache IDs, moves the cache's collection ordinal to the
end, and leaves unrelated cache snapshots unchanged. `RefreshTable` returns
Boolean True. Disabled populated caches raise the measured error 1004 before
mutation; pending caches retain their standard error 1004.

Evidence includes **four initial controls**, **nine follow-up controls** and
**five successful calculation/event controls**, plus **two imported-refresh
controls**, **three growth/collision controls**, and seven native macro-enabled save/reopen samples. Seven engine-generated
and imported-refreshed output workbooks also reopen successfully in native Excel. The existing
macro-free simple Pivot fixture also has a model source-edit, refresh and
save/reopen regression. Eighteen admitted runtime
cases compare complete operation output;
five explicitly project only host-specific `Application.Build` metadata while
preserving the original raw results. Automatic source formulas refresh normally;
manual-mode sources retain cached formula values until explicitly calculated.
Every measured callback observes all shared reports already committed. Each
report notifies Change, Calculate and PivotTableUpdate in reverse creation order,
including Calculate in manual mode; automatic mode adds a final Calculate.
Disabled events suppress callbacks.

Three native header-change controls successfully remove affected measures and
invalidate removed-field proxies; this first implementation instead reports
schema changes unsupported before mutation. Cross-sheet report groups, named or
external sources, OLAP, grouped/filtered/rich reports, unknown metadata and
unsupported cell features also remain outside refresh admission. Source and
output planning have explicit size limits: at most 64 reports on the affected
sheet, one million aggregate planning units and staged output cells, eight
million charged text characters and 200,000 metadata nodes. Literal source
rectangles are bound to their declared worksheet without defined-name lookup.
Malformed stored-record XML is rejected before any live refresh mutation.
Empty report expansion is supported;
expansion into populated ordinary cells, existing report page selectors, or
intersecting final shared-report footprints is rejected before mutation. Native
collision controls produce error-valued report states; those transitions remain
unsupported. Erroring event-handler behavior and
arbitrary refresh combinations remain unaudited. See the
[refresh core tests](../../tests/test_excel_pivot_refresh_core.py),
[native refresh replays](../../tests/test_excel_pivot_refresh_native.py) and
[unaltered native evidence](../../tests/fixtures/pivot_refresh).
`RefreshTable` adds one registration; checklist-verified totals remain unchanged.

Field-range measurements add **55 controls**: 32 generated axis layouts and
23 existing Excel workbooks, including compact, tabular, outline, grouped,
renamed and filtered reports. `PivotField.DataRange` and `LabelRange` select
single or separated areas from persisted report lines; `PivotTable.PageRange`
returns vertical or wrapped filter blocks in both fill orders. Eight wrapped-page
workbooks also check `TableRange2` includes every filter column. Native save/reopen
checks cover every case, including Excel's row-depth-dependent Values ranges.
Measure ranges omit grand totals while retaining displayed row subtotals; tests
also enumerate their actual cells. See the
[range regressions](../../tests/test_excel_pivot_field_ranges.py) and
[measurement script](../../scripts/measure_pivot_field_ranges.py).

Another **12 native controls** distinguish explicit `AddDataField`, which removes
its source from a row/column/page axis, from assigning `Orientation=xlDataField`,
which preserves the source axis. Invalid additions leave the existing report
unchanged. Axis edits clear obsolete filter cells and retain imported page-wrap
geometry. **Nine retained-reference controls** cover Values-axis moves and hiding,
and removed data fields returning VBA error 424 rather than Python exceptions.
See the [source-axis tests](../../tests/test_excel_pivot_source_axis.py) and
[proxy controls](../../tests/fixtures/pivots/retained_fields.json). These controls
do not complete arbitrary layout mutation, filtering/grouping, event interactions
or all mixed-layout range/error combinations. Registration and verified counts
remain unchanged because these are existing partial members.

Imported pivot edits also retain namespace declarations used by compatibility
attributes and extension QName values. This fixes native Excel refusing to open
rewritten reports/caches whose `mc:Ignorable` prefix became undeclared. The shared
XML reader/writer preserves local bindings, aliases and nested rebindings without
changing global serializer state; DTD/entity refusal remains in place. See the
[namespace regressions](../../tests/test_excel_xml_namespaces.py) and
[pivot XML security tests](../../tests/test_excel_pivot_xml_security.py).

Native conditional-formatting controls cover ordinary cell-value and expression
rules, priority and StopIfTrue composition, relative formula anchors, differential
font/fill/border/number formats, protection, partial deletion, clearing and sheet
copying. DisplayFormat reads live conditional styles without changing base cell
formats; native Excel permits its Font setters to change the base font. Range.Text
uses conditional number formats, while native DisplayFormat.NumberFormat reports
the base format. Reversed Between bounds and color-only conditional borders have
native controls. Displayed reads over more than 1,048,576 conditional cells are
explicitly unsupported. Color scales, data bars, icon sets, other advanced rule
families and structural edits still need work. Transfers outside the measured
copy/paste scope below and affected structural edits are rejected before mutation;
value-only pastes preserve the rules. Recursive UDF-driven display reads still
need auditing. See
[rule tests](../../tests/test_excel_conditional.py),
[display tests](../../tests/test_excel_conditional_display.py),
[native controls](../../tests/fixtures/conditional) and their
[replay script](../../scripts/measure_conditional_formatting.py).

Statistical conditional rules expose native Top10, AboveAverage and UniqueValues
objects, including top/bottom percentages, ties, inclusive average/deviation
boundaries and duplicate/unique selection. Native controls establish numeric,
logical, error, text, empty-string and blank distinctions, disjoint/overlapping
areas, hidden rows, 15-digit numeric comparisons and asymmetric wildcard criteria.
Lossless `ChrW` controls distinguish ordinary canonical Unicode equality from
wildcard matching over raw UTF-16 units. Measurement sources stay ASCII so the
VBA editor's source-code conversion cannot silently change the input text.
Measured numeric-text/date parsing and linguistic comparisons do not establish
complete locale fidelity. Macro-free native samples test imports; generated
workbooks test native rule inspection and displayed formats after copying.
Population summaries are shared within one range-property read and discarded
afterward, so later edits cannot reuse an old summary. Sparse whole-column rules
count blank partners analytically. Each distinct wildcard query can still scan
the text population, and repeated scalar queries can rebuild summaries. Pivot
scopes and recursive or rule-mutating statistical reads are explicitly unsupported.
See [API tests](../../tests/test_excel_conditional_statistics.py),
[statistical evaluator tests](../../tests/test_excel_conditional_stats_eval.py),
[native import tests](../../tests/test_excel_conditional_statistics_native_import.py),
[copy tests](../../tests/test_excel_conditional_statistics_copy.py) and
[population-read tests](../../tests/test_excel_conditional_statistics_reads.py).

Text, blank, error and time-period conditional rules share the existing lifecycle,
differential styles, displayed formatting and worksheet-copy paths. Generated
compatibility formulas use the bounding top-left applicability corner as their anchor and are
rebuilt when attribute-driven rules move. Ordinary expression formulas retain
their separate behavior. Calendar-relative native controls cover all ten date
periods, fractional-day boundaries, Sunday-based weeks and adjacent-month boundaries;
native-created files also exercise import and subsequent model saves. Date rules
currently use the 1900 workbook epoch; 1904 workbook fidelity remains incomplete.
Typed controls distinguish spaces, tabs, nonbreaking spaces, formula-empty text,
logical values and errors. Native nonblank rules include error cells despite the
stored compatibility formula returning an error for that input. Measured modern
worksheet text functions count supplementary Unicode characters as one character;
SEARCH wildcard matching still consumes UTF-16 units, with character-based start
and result positions. Wildcard SEARCH evaluations requiring more than 8,000,000
matching states are explicitly unsupported; literal searches use a separate path.
These controls do not establish every locale or Excel compatibility-version mode.
Native string controls distinguish plain rule-text XML attributes from escaped
formula text, including literal escape-looking strings, CR/LF, paired and unpaired
surrogates, and the UTF-16 input-length boundary. Saves normalize rule-text CR and
unpaired surrogates independently of the formula. Saving NUL-containing rule text
is explicitly unsupported: native Excel also failed to reopen its own equivalent
saved file. In-memory rule inspection still supports that input.
The predicate fixtures contain 179 replayable native controls, one recorded
native-open limitation and 23 macro-free native workbook samples. Generated
workbooks also pass 53 native copy and displayed-format checks.
See [predicate API tests](../../tests/test_excel_conditional_predicates.py),
[evaluation tests](../../tests/test_excel_conditional_predicates_eval.py),
[copy tests](../../tests/test_excel_conditional_predicates_copy.py) and
[native rendering/import tests](../../tests/test_excel_conditional_predicates_rendering.py).

Native UDF controls distinguish fresh displayed-format acquisition from retained
objects. A worksheet UDF receives error 1004 when acquiring Range.DisplayFormat;
objects acquired beforehand remain readable, including across workbooks.
Conditional-expression UDFs and Application/Worksheet Evaluate calls can acquire
fresh displayed formats. Evaluation origins are restored after nested calls,
errors and deferred calculation. A worksheet UDF reading its own Caller.Text
uses the current cached value, or a temporary zero for an empty result cache; it
does not store that provisional value. Ordinary Value/Value2 circular-reference
handling is unchanged. The native fixtures replay 32 controls. Self-recursive
conditional-UDF partial rendering is still incomplete; its separate raw native
control is retained without claiming conformance. See the
[UDF native regressions](../../tests/test_excel_conditional_udf.py) and
[calculation-context tests](../../tests/test_excel_conditional_udf_context.py).

Conditional-format range transfers have **13 native controls** covering
`Range.Copy Destination`, `Worksheet.Paste` and `PasteSpecial` with all or formats.
The measured scope is a disjoint rectangular source and landing in one workbook,
initially covering one incoming rule, which may cover multiple areas. Copies clip
its applicability, shift its formula from the original rule anchor and give it
an independent differential style. Destination rules lose the entire landing
area, including positions without a source rule. New rules take priority one;
retained destination-sheet priorities increase without closing existing gaps.
Statistical populations remain one rule across copied fragments, and formats
pastes with `SkipBlanks` retain their full conditional applicability. Independent
native reopen controls cover both macro-free and macro-enabled workbooks.
Worksheet UDF attempts return `False` without changing the workbook; ordinary
macro calls still perform the copy. Overlaps,
tiling, transposition, shared-group or ambiguous loaded sources, merging paste modes, cross-book
transfers and affected mixed-feature combinations remain guarded. Narrow-column
`Range.Text` width truncation remains incomplete; these controls use measured
20-character columns to isolate conditional formatting behavior. See the
[transfer regressions](../../tests/test_excel_conditional_transfer_core.py) and
[native transfer/import tests](../../tests/test_excel_conditional_transfer_native.py).

A further **8 native controls** cover fresh multi-rule transfers. Incoming
priorities follow source priority, while serialized clone order follows source
creation order. Each rule retains its own clipped applicability and statistical
population; destination priorities increase by the number of intersecting source
rules. Held source-style edits do not change copied differential styles. Shared XML containers are validated once per transfer, and cloned containers omit sibling rule payloads; [count-based regressions](../../tests/test_excel_conditional_transfer_groups.py) verify this for 8- and 48-rule groups. Loaded rules with distinct validated group identities now have a separate measured scope below; shared-group sources and inconsistent copied identity metadata remain guarded.
Another **15 native controls** distinguish specialized property contexts:
`UniqueValues.DupeUnique` reads the active sheet's matching creation identity,
returns error 440 when that identity is absent, and reports duplicate mode for
the measured foreign ordinary, Top10 and AboveAverage rules. Its setter changes
the owning rule. The measured Top10 and AboveAverage scalar getters and setters
remain owner-bound. See the [multiple-rule and property regressions](../../tests/test_excel_conditional_multiple.py)
and [native evidence](../../tests/fixtures/conditional_multiple/native.json).

Loaded conditional groups have **36 native ordering controls** and **6 native
multi-rule copy controls**. Plain loads compare the complete ordered applicability
population by each area's column interval, then row interval; an exact-prefix
population with more areas sorts first. Equivalent projected populations in
compatible plain containers share one loaded identity. VBA-bearing loads retain
raw container and child order. Each rule keeps its own root, differential style
and API values; foreign statistical border colors resolve the matching group
and captured child position, with the first child serving as representative. Opaque container metadata and incompatible namespace bindings stay
separate and are preserved.

A further **20 native controls** distinguish load mode from FileFormat. A
FileFormat52 package without an actual VBA project uses the measured plain-load
projection, group coalescing and initial priority-based child order. A VBA-bearing
FileFormat52 package retains raw ordering even when opened with `with_vba=False`.
The mode is captured before macro loading. Generic API projection remains a
separate mutable state: `ModifyAppliesToRange` keeps newly supplied union order.
Saving after adding a VBA project materializes the captured plain order, including
when the output remains FileFormat52; unchanged parts remain untouched when no
conversion or edit requires rewriting them.

**18 shared-group lifecycle controls** cover individual and bulk deletion,
priority changes and range splits. Initial child order is captured once, so a
priority edit does not silently replace the representative on save. Individual
deletion keeps the target's tombstone identity and moves remaining children
together to a new identity; bulk deletion preserves the prior representative.
Moving one shared child retains the untouched group and appends the moved rule,
including a move to spatially earlier cells. **Two generic-union copy controls**
map the source's projected area order while preserving its bounding formula
anchor. **Three ModifyAppliesToRange controls** cover ordinary relative-formula
rebasing and raw union order, including a fresh-rule native save/reopen.

Shared-group range-copy sources remain explicitly unsupported. Initial
active-sheet lookup context and some retained native proxy rebinding still need
work; their raw observations are not treated as passing full-context controls. See the
[loaded-group tests](../../tests/test_excel_conditional_loaded.py),
[mode/lifecycle/rebase tests](../../tests/test_excel_conditional_loaded_lifecycle.py)
and [native evidence](../../tests/fixtures/conditional_loaded).
These changes add behavior to existing members; registration and verified totals
are unchanged.

Worksheet.Copy identity behavior has **8 fresh-rule native probes** and
**4 loaded-rule native probes** covering distinct and shared groups in both
plain and VBA-bearing workbooks. Hashed native JSON evidence retains its exact
bytes across platform checkouts. Copied live groups receive dense independent
identities and a matching fresh-rule counter. Copied children recapture priority
order independently of a source VBA-bearing group's raw child order; original
identities, tombstones, roots and stored styles remain unchanged. Foreign
statistical border-color reads use the caller's compact position within its
captured group, while owner-active reads retain that rule's own style.

The passing scope includes **64 fresh and 16 loaded explicit-activation phases**.
Three initial post-copy snapshots have known context differences and are retained
as separate raw evidence rather than treated as full-conformance results.
Two additional narrow native controls preserve creation order when rule geometry
is reversed and confirm a coalesced fresh same-population copied group through
a safe receiver sheet and its next new-rule identity. These controls do not
establish full same-population context fidelity: foreign reads of an original
source rule after copying that population still differ from native Excel.
That original-source lookup and the corresponding broader retained-proxy
observations remain deferred; the safe copied-sheet new-rule scope passes.
One [successful post-deletion follow-up](../../tests/fixtures/conditional_sheet_copy/post_delete_observation.json)
observes compact surviving member rank; [additional diagnostic observations](../../tests/fixtures/conditional_sheet_copy/post_delete_diagnostic_partials.json)
leave missing-child defaults after deletion inconclusive. Foreign lookups of an
absent child inside an existing group therefore report unsupported; a missing
group retains error 440. A separate successful native copied-sheet control covers
valid shared-group `UniqueValues.DupeUnique` child positions, with setters still
owning their original rule. Broader retained-proxy and post-deletion contexts
remain partial. Shared-group Range.Copy guards are unchanged.
A worksheet-local member/rank index avoids rescanning groups on repeated foreign
reads; mutation invalidation and failed-preflight preservation have count-based
and atomicity regressions. See the
[worksheet-copy context tests](../../tests/test_excel_conditional_sheet_copy.py)
and [native evidence](../../tests/fixtures/conditional_sheet_copy).
No public member registrations or checklist statuses change in this increment.

Another **10 native controls** establish conditional-format API projections:
number-format getters remove measured unnecessary space/hyphen escapes outside
quoted literals. Plain-loaded generic rules enumerate applicability areas by row,
then column; specialized statistical rules retain serialized area order.
VBA-bearing macro-enabled imports retain serialized order for both families. Runtime-created
ranges retain their supplied order through a save. Converting a loaded macro-free
workbook preserves its family-specific order in the resulting macro-enabled file.
Getter projections preserve stored XML and
differential styles. Eight additional native controls establish the bounding
top-left formula anchor for ordinary expression rules over unsorted unions,
including a bounding corner outside the union. Stored formulas and displayed
evaluation use that anchor independently of the API's area enumeration order.
These controls do not establish every number-format spelling or all union-formula
mutation behavior.
See the [API projection tests](../../tests/test_excel_conditional_api_projection.py).
Ten further family controls distinguish generic and statistical area ordering;
four typed union controls establish bounding anchors for text, blank, error and
date predicates through rendering and native imports. The
[format conversion tests](../../tests/test_excel_conditional_format_modes.py)
also verify that saving does not reorder the current in-memory rule objects. Separate [native XLSM reopen controls](../../tests/fixtures/conditional_api_projection/initial_reopen52.json) preserve raw applicability order for the existing ModifyAppliesToRange and two-dimensional deletion cases.

Statistical-rule border colors have **24 native controls** across active
sheet changes, priority gaps, reordering, deletion and range copying. Native
color getters can resolve another sheet's stable creation slot, or raise error
440 when that slot is absent. The engine reproduces these transient getters
without changing the rule's stored style, displayed border, line style or
weight. Additional loaded-group controls distinguish their identities from
priority and displayed area order. Color assignment creates a thin line when no visible style exists and
preserves an existing dashed line. Loaded identities follow the file's range
ordering and restart after persisted deletions. Imported style edits can
also change native bindings, but their transition rules remain incomplete;
the raw observations are retained separately from passing controls. Conditional
border theme/tint setters and broader imported-context behavior still need work.
See the [border-context regressions](../../tests/test_excel_conditional_border_context.py).

## Table-owned sorting

`ListObject.Sort` exposes shared fields/options through fresh Sort views with the
ListObject as Parent and measured Application/Creator values. Ordinary local
Apply sorts stable whole data rows, excluding headers/totals, retaining row
styles and shifting moved formulas. Planning validates keys, moved formulas and
replacement XML before publishing cells. Sort serialization is preflighted before package edits. Fixed-width height growth and appending
a row preserve held Sort/SortField views and original key spans. Table SetRange
fails with native error 1004.

Evidence retains **28 raw native controls and 24 original macro-free workbooks**.
Fourteen complete Table.Sort runtime controls and four additional worksheet-sort
persistence controls replay without projections. Eight whole Table.Sort controls
and two non-default worksheet-sort persistence combinations remain deferred.
The Table.Sort and worksheet-persistence suites pass **120 tests**, including native
input snapshots, public metadata rewrites, generated save/reopen, resource limits
and atomic unsupported paths. All eight engine-generated outputs passed native
Excel reopening in one closed owned session: five original table outputs and
three post-Apply/imported configuration outputs matched independent snapshots.
Native controls confirm protected-table sorting succeeds and worksheet-UDF Apply
sorts rows while returning Empty. Automatic Apply emits Calculate; manual Apply
emits none; disabled events suppress the callback.

Direct Sort.Rng and SortField.Key ranges have Sort and SortField as Parent.
Their Parent.Name raises 438; the range getters succeed. Descendant-range owner
propagation remains unmeasured. Native persistence distinguishes unapplied table
fields/options, which are dropped, from independently configured Worksheet.Sort
settings, which are retained. After a previous table Apply or import, field Order,
MatchCase and SortMethod metadata edits persist without another sort; Clear removes
the persisted group. DataOption remains a runtime setting and reopens as zero.
New table fields after Clear require Apply before the next save in this subset.
Native worksheet controls additionally cover loaded
field Order/MatchCase edits, DataOption resetting on reopen, Clear removing Rng,
and replacement fields without a new range. The worksheet reader/writer preserves
untouched imported XML and writes current measured configuration rather than stale
applied settings. Rich/wrapped imported sort metadata and non-default worksheet
orientation/method persistence are explicitly refused. Worksheet sort hydration streams the original document and limits only selected
sort metadata to four million bytes and 65 elements (64 persisted keys), without limiting unrelated
worksheet data size. The 64-key persistence limit is an engine admission bound,
not a measured Excel maximum; runtime Add retains its existing behavior. Selected metadata requires UTF-8 declarations and charges
complete opening and closing tags against the byte limit. Exact-boundary controls
are retained in `tests/test_opc_xml_selected.py`.

Value sorting admits single-column keys, ordinary formulas and cell styles.
It refuses active table/sheet filters, union/multi-column keys, custom/rich/stroke
sorts, annotations, conditional formatting, validations, arrays/spills, unknown
table metadata and unsupported configured-table edits before mutation. The
current table writer admits native unprefixed, double-quoted metadata; alternate
lexical forms, comments and CDATA are conservatively refused for Apply. Whole-table
Cut, broader held-object movement/deletion, header visibility changes and filtered
sort metadata remain incomplete. Per-transaction limits are one million data
cells, four million comparison work units, four million scanned text characters
and 64 keys. These are engine limits, not native Excel capacity claims.

See [native controls](../../tests/test_excel_table_sort.py),
[transactional regressions](../../tests/test_excel_table_sort_core.py),
[worksheet persistence controls](../../tests/test_excel_worksheet_sort_persistence.py),
[table metadata persistence](../../tests/test_excel_table_sort_persistence.py), and
[retained evidence](../../tests/fixtures/table_sort/manifest.json).

## Chart data admission

Embedded-chart source inference and supported setters prepare replacement data
before changing chart state. Failed source reads or value conversions preserve
prior series identities, plot orientation, title/legend settings and drawing
change flags. Existing native chart API controls remain unchanged; regression
coverage is in [chart staging tests](../../tests/test_excel_chart_staging.py).

Twenty additional native argument controls cover `SetSourceData`: fifteen replay
complete original source/result pairs, and five compare complete scalar-case
before/result/after bodies from an explicitly projected batch that later stopped
on a separate diagnostic. Scalar sources return error 13; omitting Source while supplying PlotBy returns
450, and unsupported PlotBy enum values return the measured automation error.
Numeric-string and fractional PlotBy coercions follow the recorded controls.
All failed admitted operations preserve the complete recorded chart snapshot.
Raw evidence retains two deferred diagnostics: a successful native union source,
whose inference remains unsupported, and a Nothing source that disconnected the
held native chart. The latter is not treated as a passing in-memory lifetime
assertion. See [argument controls](../../tests/test_excel_chart_source_arguments.py)
and their [dispositions](../../tests/fixtures/chart_source_arguments/manifest.json).

Engine resource bounds admit at most **one million cells/points** per source,
series input or aggregate reference read, and **1,024 inferred series** per
`SetSourceData` call. References are bounded to 32,768 characters before parsing
or storing them. Imported literal-cache point indexes must be nonnegative
integers below the point bound before dense allocation. These are engine
admission bounds, not measured Excel maxima. Cross-workbook source ranges and
cross-workbook or multi-area series ranges are explicitly unsupported rather
than silently binding to a local sheet or only the first area. Imported chart
content mutation, axes, series lifecycle and broader chart fidelity remain
incomplete. No public member registration or checklist status changes here.

## Inventory scope and maintenance

The reference is the Microsoft Excel 16.0 object-library dump in
[pyVBAReference at cecb1c1](https://github.com/WilliamSmithEdward/pyVBAReference/tree/cecb1c1515dfeb0d23f31727122ecbbe4c8b699a/reference/excel/json).
The raw dump contains **13,876 member entries across 746 types**, including paired
COM interfaces such as `IWorksheetFunction` and `WorksheetFunction`. This summary
omits an `Iâ€¦` interface when its corresponding non-prefixed type exists, and
counts unique member names per remaining type. Every remaining type is assigned
to one functional area. Properties, methods and events count as member names;
enum constants are outside this table.

The checklist contains **2,756 entries across 70 implemented model types**:
147 verified, 635 partial, 424 unassessed, 1,547 missing and 3 excluded. Its 1,206
registered entries include four helpers/entries without a matching reference
member; the table's numerator excludes those four. Entirely unimplemented types
are included in this table's denominator even though they are absent from the
selected-type checklist.

Functional areas own their object types. For example, `Range.Font` belongs to
Cells and ranges, while the `Font` object's members belong to Cell formatting
and styles. This avoids counting an entry point twice. Shared Office, VBIDE and
MSForms libraries are outside the Excel type-library totals, even where Excel
automation uses them.

When support changes, regenerate `docs/excel_checklist.csv` with
`python scripts/build_excel_checklist.py`, update the operation evidence and
refresh the affected functional rows and totals here. An area's status becomes
complete only after its required behavior, error cases, interactions and
persistence have been auditedâ€”not merely when its member count reaches 100%.

## Embedded chart series lifecycle

Ordinary engine-created embedded charts support `Series.Delete` and `PlotOrder`.
Deletion returns the measured Boolean result, compacts surviving order and keeps
held surviving series attached to their original data. Deleted series reads
through `CallByName` raise the measured automation error, including after a
replacement series is added. Direct deleted `Name` access reports the measured
invalid-object HRESULT; a narrow host dispatch hook preserves that distinction
without changing ordinary VBA objects. Reordering moves the existing entry rather than copying its state;
invalid measured zero/out-of-range orders leave entries and dirty flags intact.
Argument coercion revalidates ownership and membership before applying an order,
including when a VBA default getter deletes or reorders another series.
`XValues` returns independent array storage, so reading or modifying a returned
array cannot clear or change chart categories.

Two independent native controls retain their complete raw source/results and
explicitly project chart type/count, every series observation, operation result
and held-series observation. Axis and parent-graph observations remain in the raw
evidence and are explicitly deferred. The same projected final snapshots match
two native-produced workbook imports and two model-generated save/reopens.
This is partial lifecycle evidence, not a complete chart or member audit.
See `tests/test_excel_chart_lifecycle.py` and `tests/fixtures/chart_lifecycle/`.

Imported chart content mutation remains unsupported. Mutations and series access through a held chart after its
owning drawing/sheet/workbook is detached are refused explicitly pending native
lifetime measurements. Parent graphs, argument/coercion variants, styles after
reordering, protection/UDF interactions, automatic axes/scales and broader chart
editing remain incomplete. All five generated lifecycle workbooks passed native
Excel readback (5/5, 5.93 seconds) in one closed owned session.

Two further completed native `PlotOrder` matrices cover fractional rounding,
numeric and invalid strings, Boolean/Empty/Null, range errors and numeric
overflow, including complete before/after/reset series snapshots. Their replay
explicitly defers the parent graph and a separate `Series.Index` late-binding
discrepancy; those raw observations remain unchanged. Null assignment now raises
the measured type-mismatch error. A third combined stale-series deletion control
timed out without a complete returned value; it remains diagnostic only, and its
eleven unexecuted follow-ups are not counted as passing evidence.

Three additional complete native controls establish stable unnamed-series IDs,
reordering/deletion and explicit/reference-name collisions. Default names follow
the smallest unused series ID rather than reserving displayed text. Their five
native-produced/imported and model-generated save/reopen pairs (including the
two initial lifecycle controls) retain exact projected final snapshots. The
writer preserves `c:idx` separately from `c:order`, leaves default formula names
blank and retains identity-based default colors. The reader sorts ordinary
series by their declared order, retains IDs and refuses ambiguous or malformed
identity/order metadata before integer expansion. The XML identity/order values
are compared directly with the three independent native outputs.

Native measurement paths retain their original Windows provenance. Chart fixture
replays resolve those basenames explicitly so the same recorded inputs work on
other platforms.
