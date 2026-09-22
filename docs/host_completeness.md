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
| VBA WorksheetFunction | 9 | 406 | Separate implementation from the cell formula engine |

The cell formula engine registers 122 names; its current inventory lists
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
UNASSESSED. Nothing is VERIFIED yet: no area has finished its
operation-level audit.

## Excel work areas

All rows below require an operation-level audit before they can become
VERIFIED. This is the initial work breakdown, not an exhaustive list of
Excel features. Reference members and file-format features must be
assigned to these areas or added as new areas, with every unresolved
entry tracked explicitly.

| Area | Current evidence | Remaining audit and implementation |
| --- | --- | --- |
| Packages and formats | VBA editing covers xlsm, xlsb, xlam and xls; runtime loads xlsx, xlsm and xlam | Inventory creation, open, save, conversion and persistence by format; xlsb and xls cell models are absent |
| VBA projects and forms | In-scope MS-OVBA gates, designer round trips and live file gates exist; shared library-reference add/remove operations pass native Excel/Word/PowerPoint checks and Access persistence tests | Reconcile every operation with its format and protection/signing limits; distinguish form design from form runtime; linked VBA-project creation and arbitrary library runtime resolution remain separate gaps |
| VBA language and runtime | Measured semantics, classes, errors and control flow exist | Audit complete syntax, type system, lifetime, references and intrinsics; track file I/O, external calls and known typing gaps explicitly |
| Application, workbooks and sheets | Core collections, cells, names and queries exist | Map all reference members; identify stubs, missing collections and incomplete arguments |
| Ranges and editing | Values, formulas, areas and basic movement exist; single-area Find/FindNext/FindPrevious pass 48 native cases; merged-range geometry/value operations pass 21 native cases and four persistence gates; FormulaR1C1 scalar and array writes/reads pass 27 native cases and five persistence gates; A1 Formula array writes share the implementation and pass 15 native cases plus four persistence gates | Complete multi-area/comment/format/MatchByte searches, complex across/multi-area merges and merge formatting; audit insertion, deletion, copying, sorting and filtering |
| Named ranges | Workbook/worksheet CRUD, validation, scope precedence, rename/retarget/delete, visibility/comments and Python snapshots share one implementation; 19 native cases and three persistence gates | Active-cell-relative name semantics, locale translation, Excel 4 macro metadata and arbitrary named-formula evaluation remain incomplete |
| Multiple workbooks and sheet copying | Add/open/activate/close; VBA sheet Copy/Move before/after or into a new book, matching native held-object invalidation on cross-book moves; inter-book range copies import used names and aliases with reuse/rename/error policies, backed by native probes and persistence gates | External workbook links, independent VBA projects, sheet-copy/move name conflicts, relative/other-sheet name imports, transferred drawings/controls, sheet modules and advanced worksheet metadata remain incomplete |
| Structural edits | Whole-row/column insertion/deletion and formula/name reference updates pass 12 native cases and four persistence gates; occupied-edge insertion is atomic | Partial-cell reference rewriting, relative name context, merges/shapes, formatting inheritance/CopyOrigin and table/validation/chart metadata updates remain incomplete |
| Range copying | Explicit-destination copies pass 17 native cases and five persistence gates: overlap, blanks, repeated blocks, formulas, cross-sheet copying and measured formatting | Merged-cell and multi-area copies, destinations larger than 1,048,576 cells, clipboard/paste behavior and full formatting fidelity remain incomplete |
| Formatting and layout | Cell formats are complete for Font, Interior, Borders, alignment, protection, NumberFormat and ClearFormats: 161 live probes, an Excel-authored workbook read back value for value, the border storage rules, a live gate for model-written formats, and a stylesheet Excel's for the same edits except one font-order quirk | Row and column formats and whole-row/column formatting, named cell styles (`Range.Style`, `Styles`), conditional formatting, validation, dimensions (AutoFit, widths, heights, hidden) and page setup |
| Tables and structured references | Power Query reads workbook tables and writes load targets | Complete the ListObjects/Table API and structured-reference semantics; preserve and edit table metadata |
| Formulas and calculation | Formula parser, dependency invalidation, manual mode and 122 registrations exist | Complete function and argument coverage, dynamic arrays, reference forms, errors, circular/iterative calculation and cached-value persistence |
| VBA WorksheetFunction and Evaluate | Nine WorksheetFunction bindings and VBA evaluation exist | Audit consistency with cell calculation, Excel Evaluate semantics, argument coercion and error translation |
| Events | EnableEvents state exists | Connect document handlers and WithEvents; measure event ordering, target ranges, suppression and reentrancy |
| Shapes and controls | Public SheetView snapshots and add/edit/delete APIs cover measured AutoShapes, text boxes and nine Forms control types; control edits reach supporting parts; checkbox/list/numeric/radio control values, numeric bounds/increments, A1 and workbook/local named sources/links (including CHOOSE, IF, single-area INDEX, OFFSET and A1 INDIRECT), inline items, whole-list arrays, range-source conversion and multi/extended selections and rebinding replay live measurements and pass persistence gates | Complete other reference formulas and external control links, broader named-reference cases, grouping, copying, pictures, geometry, formatting, anchors and macro links |
| Charts | Existing chart parts can be preserved and shapes identified | Audit chart objects, creation, series, axes, formatting, data links and round trips |
| Pivots, caches and data model | Unmodelled content is preserved in existing files | Inventory and implement headless operations; preservation alone is not runtime support |
| Power Query | Packaging, M evaluation, local refresh and sheet load targets exist | Complete language/library conformance, data types, errors, query dependency behavior, connectors and refresh settings |
| Other workbook features | Coverage varies and needs classification | Inventory names, links, comments/notes, hyperlinks, protection, metadata, embedded objects, connections and remaining reference/file features |
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
