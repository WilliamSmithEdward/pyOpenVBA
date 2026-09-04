# Measuring a design: control slots and property codes

A form or report is a stream of property records:

    <u32 id><u16 code><u32 value type><u32 width><u32 length><value>

The `code` names the property. The `id` is the record's slot in **that
control type's own schema**, so the same property sits at a different id
on a label than on a combo box, and a type whose ids have not been
measured cannot be written at all. These scripts measure both.

They need Windows, desktop Access and `pyvbaharness`, and they are
dev-time tools: pyOpenVBA itself never touches COM.

## Control slots

`extract_control_slots.py` reads a database in which Access has built one
control of each type and prints the `(id, code, value type, width)` of
every record, ready to paste into `CONTROL_SLOTS` in
`src/pyopenvba/access/_designs.py`.

Build the database first with `CreateControl` for each `acControlType`.
A page needs its tab control named as the `Parent` argument, which is the
fourth, not the third:

    CreateControl(form, 124, 0, tabs.Name, "", 0, 0, 100, 100)

What the measurement showed, per type:

- A tab control has no left or top of its own; a page break has only a
  top; an image has no overlap flags; a combo box carries its GUID ahead
  of its name.
- Code 261 is the tab index, confirmed against `Control.TabIndex` for
  every control that has one. Access omits the record when it is 0.
- Code 0 holds `-1` on exactly the controls whose `PictureData` and
  `ImageData` read back as `-1`, and on no others.
- The word an `0xFF` marker carries is **how many objects the group it
  opens holds**, the opener included. A form with eleven controls carries
  `0xFF 11` twice: once over the prototypes and the detail section, once
  over the controls. Access does not refuse a wrong count -- it opens the
  form and shows only that many controls.

## Property codes

`build_rich_form.py` makes a form with one control of every type and sets
as many properties as VBA will set, each to a value nothing else on the
form holds. `name_property_codes.py` then pairs the blob against
`SaveAsText`, which writes the same properties with their names.

Pair on the **value**, not the position. The blob carries records
`SaveAsText` does not write, so a straight walk of the two drifts and
starts naming codes wrongly a few records in. A property whose value
appears exactly once on each side of one object can only be that record;
a name is kept only when every object carrying the code agrees.

The check that the method worked is built in: it re-derives the codes
already named and reports any it contradicts. Both runs reproduced all
seventeen testable ones and contradicted none, which is what makes the
twenty-eight new names evidence rather than a guess.

Distinct values matter. With `FontSize`, `BorderWidth` and
`DecimalPlaces` all set to 2, none of the three can be told apart; with
17, 5 and 7 all three fall out.
