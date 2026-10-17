# Excel API coverage

Snapshot: **17 October 2026**, covering the current source tree, including the
worksheet-source PivotCache creation, compact PivotTable axes, and ordinary
conditional formatting rules and display evaluation. This is a functionality summary for the
in-memory Excel engine. **No functional area is complete yet.** Common worksheet
operations have substantially more support than advanced Excel automation.

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
| Application, workbooks and worksheets | Partial | 125 / 752 | 6 | Core containers, opening/saving, sheet operations and settings. Many application services, workbook features and host-specific behaviors remain. |
| Cells and ranges | Partial | 112 / 215 | 29 | Reads/writes, formulas, references, copying, moving and many editing operations. Native probes cover Currency precision and SpreadsheetML text/formula persistence for literal escapes, controls and Unicode. Full argument, error, structural-edit and mixed-feature fidelity remains incomplete. |
| Formulas and calculation | Partial | 388 / 463 | 0 | Broad worksheet-function exposure, dependency calculation, UDFs, arrays and calculation modes. Registered functions still need behavioral audits; remaining functions and Excel numerical edge cases are incomplete. |
| Defined names | Partial | 15 / 30 | 0 | Workbook/sheet names and reference evaluation. Full expression, relative-reference, external-reference and lifecycle behavior remains incomplete. |
| Cell formatting and styles | Partial | 103 / 206 | 20 | Fonts, fills, borders, number formats, alignment and named styles. Ordinary conditional DisplayFormat views are supported. Rich text, gradients, advanced display formats and remaining formatting APIs need work. |
| Conditional formatting | Partial | 22 / 212 | 0 | Ordinary cell-value/expression rules, priorities, applicability ranges, differential styles, displayed formatting and conditional number formats. Native controls cover relative formulas, rule composition, clear and worksheet copy. Advanced rule families, range copy/paste and structural edits remain incomplete. |
| Excel tables | Partial | 45 / 155 | 25 | ListObjects, rows/columns, structured references, resizing and totals. Table-style objects and broader table-operation fidelity remain incomplete. |
| Sorting, filtering and outlines | Partial | 28 / 66 | 6 | Common sorts and AutoFilter operations. Advanced criteria, sort variants and outline/grouping APIs remain incomplete. |
| Data validation | Partial | 19 / 21 | 14 | Validation rules, formulas, input/error messages and persistence. Remaining members and full rule/error behavior need auditing. |
| Protection and permissions | Partial | 11 / 32 | 0 | Worksheet/workbook protection and enforcement for supported edits. Allow-edit ranges, user permissions and full protection semantics remain. |
| Comments and hyperlinks | Partial | 35 / 74 | 25 | Classic notes and cell hyperlinks, including many edit/move/save paths. Threaded comments, richer note geometry and remaining hyperlink behavior are incomplete. |
| Shapes and worksheet controls | Partial | 54 / 2,025 | 0 | Basic shapes/text boxes and nine Forms control types, with measured values, links and persistence. Rich drawing APIs, grouping/copying, pictures and newer controls remain incomplete. |
| Charts and sparklines | Partial | 30 / 951 | 0 | Embedded charts, basic series/titles and eight persisted chart kinds. Axes, richer formatting, imported-content mutation, chart sheets and sparklines remain incomplete. |
| PivotTables and slicers | Partial | 65 / 637 | 0 | Imported report queries, GETPIVOTDATA, cache metadata, empty-report creation and data-only AddDataField totals using all eleven aggregate functions, with measured events, field ranges and saved-report edits. Native probes cover date/time serials, Currency, typed values, errors, numerical edge cases and immutable cache snapshots across saves. Compact row/column/page layouts, nested subtotals, Values-axis placement and field removal now aggregate cached records. Field/page range queries support native measured compact, tabular, outline, grouped and multi-area reports; sparse layouts avoid scanning all records for each blank intersection. Refresh, arbitrary layout edits, filtering, grouping and slicers remain incomplete; general linked rich-data creation is unsupported. |
| Queries and external data | Partial | 8 / 278 | 0 | Workbook queries plus local Power Query evaluation, refresh and sheet loading. Connection objects, remote connectors, QueryTables and full query semantics remain incomplete. |
| Data model and OLAP | Missing | 0 / 261 | 0 | No in-memory Power Pivot model, relationships/measures or cube automation. Reading some external-pivot metadata does not implement these operations. |
| XML mapping | Missing | 0 / 75 | 0 | No XmlMaps, schema/namespace binding or mapped-data import/export API. |
| Excel events | Partial | 0 / 123† | 0† | Change, calculation, selection, activation, lifecycle and custom WithEvents paths have native measurements. Remaining events and full signature/dispatch behavior need work. |
| Windows, views and dialogs | Partial | 38 / 304 | 23 | Selected window, view, freeze/split and visibility state. Full interactive UI, dialogs, protected views and additional view APIs remain incomplete. |
| Printing and publishing | Missing | 0 / 152 | 0 | No full PageSetup, print rendering, page-break or publishing automation. |
| Add-ins and embedded automation | Missing | 0 / 124 | 0 | No full Excel AddIns, embedded OLE/ActiveX or external automation-host implementation. |
| Legacy features and utilities | Missing | 0 / 386 | 0 | Legacy menus/toolbars, dialog/module-sheet utilities, SmartTags, scenarios and assorted web/research services remain. |
| **Total** | **Incomplete** | **1,098 / 7,542** | **148** | **14.6% of reference members have registrations; 2.0% are checklist-verified. These are API inventory ratios, not compatibility scores.** |

† Excel events use the interpreter's event dispatcher and sinks rather than
ordinary registered object members. The registration counter therefore does not
measure the event behavior already implemented. Event status comes from the
native traces described in the [host coverage plan](../host_completeness.md).

## Excel engine behavior beyond member counts

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
generated workbook and reproduce the intended behavior. Passing the current
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
families, range copy/paste and structural edits still need work. Affected
format-bearing transfers and structural edits are rejected before mutation;
value-only pastes preserve the rules. Recursive UDF-driven display reads still
need auditing. See
[rule tests](../../tests/test_excel_conditional.py),
[display tests](../../tests/test_excel_conditional_display.py),
[native controls](../../tests/fixtures/conditional) and their
[replay script](../../scripts/measure_conditional_formatting.py).

## Inventory scope and maintenance

The reference is the Microsoft Excel 16.0 object-library dump in
[pyVBAReference at cecb1c1](https://github.com/WilliamSmithEdward/pyVBAReference/tree/cecb1c1515dfeb0d23f31727122ecbbe4c8b699a/reference/excel/json).
The raw dump contains **13,876 member entries across 746 types**, including paired
COM interfaces such as `IWorksheetFunction` and `WorksheetFunction`. This summary
omits an `I…` interface when its corresponding non-prefixed type exists, and
counts unique member names per remaining type. Every remaining type is assigned
to one functional area. Properties, methods and events count as member names;
enum constants are outside this table.

The checklist contains **2,642 entries across 65 implemented model types**:
148 verified, 528 partial, 426 unassessed, 1,537 missing and 3 excluded. Its 1,102
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
persistence have been audited—not merely when its member count reaches 100%.
