# Measuring a design: control slots, property codes, property slots

A form or report is a stream of property records:

    <u32 id><u16 code><u32 value type><u32 width><u32 length><value>

The `code` names the property. The `id` is the record's slot in **that
object type's own schema**, so the same property sits at a different id on
a label than on a combo box, and a type whose ids have not been measured
cannot be written at all. These scripts measure both.

They need Windows, desktop Access and `pyvbaharness`, and they are
dev-time tools: pyOpenVBA itself never touches COM.

## The three methods, weakest to strongest

**Positional pairing is not good enough.** `SaveAsText` writes the same
properties with their names, but the blob carries records the text does
not write, so a straight walk of the two drifts and starts naming codes
wrongly a few records in. It reproduced only 14 of the 40 codes already
named, and contradicted one.

**Value matching** (`name_property_codes.py`) is decisive where values
are distinctive: a property whose value appears exactly once on each side
of one object can only be that record. Build the form with
`build_rich_form.py`, which sets every property VBA will set, each to a
value nothing else holds. Distinct values matter -- with `FontSize`,
`BorderWidth` and `DecimalPlaces` all set to 2, none of the three can be
told apart; with 17, 5 and 7 all three fall out.

**Differencing** (`name_by_differencing.py`) is the only thing that
settles a property whose values are small integers every other property
also uses. Build the same form twice, identical but for one property, and
see which record moved. Two records always differ -- Access mints a new
GUID each time it builds a control -- so filter that one out and what is
left is the property.

Differencing is what caught the worst error in the table: `TextAlign` was
recorded as code 379, which is really `IMESentenceMode`. Access writes 3
there on every new text box, so nothing looked wrong until someone tried
to set the alignment and it did not change. `TextAlign` is code 136.

Every method checks itself the same way: it re-derives the codes already
named and reports any it contradicts.

## Control slots

`extract_control_slots.py` reads a database in which Access has built one
control of each type and prints the `(id, code, value type, width)` of
every record, ready to paste into `CONTROL_SLOTS`.

Build the database with `CreateControl` for each `acControlType`. A page
needs its tab control as the `Parent` argument, which is the fourth, not
the third:

    CreateControl(form, 124, 0, tabs.Name, "", 0, 0, 100, 100)

What the measurement showed, per type:

- A tab control has no left or top of its own; a page break has only a
  top; an image has no overlap flags; a combo box carries its GUID ahead
  of its name; a subform has no `ControlSource` at all.
- A caption sits at id 221 on a label and a command button, 231 on a
  toggle button and 232 on a page. Copying one type's id to another
  writes the text into whatever property that id names instead, and
  Access does not complain -- it opens the form with the caption missing
  and something else changed.
- Code 261 is the tab index, confirmed against `Control.TabIndex`. Access
  omits the record when it is 0.
- Code 0 holds `-1` on exactly the controls whose `PictureData` and
  `ImageData` read back as `-1`, and on no others.
- The word an `0xFF` marker carries is **how many objects the group it
  opens holds**, the opener included. A form with eleven controls carries
  `0xFF 11` twice: once over the prototypes and the detail section, once
  over the controls. A wrong count is not refused -- Access opens the form
  and shows only that many controls.

## Property slots

`measure_property_slots.py` sweeps every database under a directory and
builds `PROPERTY_SLOTS`: for each object type, where each named property
sits. Point it only at databases **Access** built. A database pyOpenVBA
wrote carries whatever slots pyOpenVBA believed at the time, so mixing
the two turns the script's own check -- that no slot moved between
sources -- from evidence into noise.

Across the seven databases the shipped table came from, 959 slots agreed
on id, code and value type, and none moved. Only the strings' lengths
differ, which is their text's length, not the slot's.
