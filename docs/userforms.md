# UserForm properties: VBA's names and the stored fields

pyOpenVBA reads and writes a form's design as [MS-OFORMS] stores it.
`properties()`, `get()` and `set_property()` name the fields a control's
record stores, which are not always the properties VBA shows. Where the
two differ, this page says where Excel's and Word's designers store a
property set through VBA, what setting it changes besides, and how they
lay out a MultiPage's pages. For the properties, both applications wrote
the same records, but for the alignment padding, which holds whatever
was in memory.

The measurements are `scripts/measure_form_properties.py` and
`scripts/measure_uncoupled_edge.py`, which `tests/test_form_properties.py`
replays, and `scripts/measure_multipage_resize.py`, which
`tests/test_multipage_resize.py` replays. Each statement below about
properties and pages is checked against one of them.

## Where a property is stored

| VBA property | Stored as |
|---|---|
| `TakeFocusOnClick` (CommandButton) | `TakeFocusOnClick`: PropMask bit 9 with no data, set when the property is False |
| `TextAlign` | `Font.ParagraphAlign`: left 1, the default and not stored; center 3; right 2 |
| `Alignment` (CheckBox, OptionButton) | `VariousPropertyBits` bit 13, set for fmAlignmentLeft (0) |
| `Style` (ComboBox) | `DisplayStyle`: 3 for fmStyleDropDownCombo, 7 for fmStyleDropDownList |
| `TripleState` | `MultiSelect`: 1 |
| `ScrollBars`, `KeepScrollBarsVisible` (form, Frame) | `ScrollBars`: the bars, plus KeepScrollBarsVisible shifted left two bits; 12 when not stored |
| `Font.Bold` (form, Frame) | the form's StdFont: weight 700 |
| `Font.Italic`, `.Underline`, `.Strikethrough` (form, Frame) | the StdFont's flags: bits 1, 2 and 3 |
| `Font.Bold`, `.Italic`, `.Underline`, `.Strikethrough` (control) | `Font.FontEffects` bits 0 to 3, and `Font.FontWeight` 700 for bold |
| `ColumnWidths` (ListBox, ComboBox) | `cColumnInfo`, and after the TextProps one ColumnInfo record per column, last column first |

A Label takes no focus, and the designers site one with its flags stored
as 0x32, the default 0x33 less TabStop.

## What setting a property changes besides

`set_property` writes the one field it names. The designers change others
with some properties, and a caller who wants their result sets those too:

- `Enabled = False` sets `Font.FontEffects` bit 13, dimming the text,
  except on a ListBox.
- Any font effect comes with `Font.FontEffects` bit 30, which keeps the
  text colour automatic.
- Bold sets `Font.FontWeight` to 700 as well as effect bit 0.
- A new font face drops `Font.FontCharSet`.
- `BorderStyle = 1` and a `SpecialEffect` clear each other on a TextBox,
  and on an Image, whichever is set second winning.
- `Enabled = False` on a ScrollBar or SpinButton sets `PrevEnabled` and
  `NextEnabled` to 0.
- `Min` above the `Position` moves the `Position` up to it.

A TextBox given `BorderStyle = 1` alone keeps its sunken effect, a state
the designers never write. Excel and Word load it as it is: they report
both at run time, draw only the sunken effect, and save both back when
their designer rewrites the form. Set `SpecialEffect` to 0 as well for a
single border.

## A MultiPage's pages

A MultiPage's TabStrip is as large as the MultiPage, and the designers
lay out only the page it shows: two pixels inside the TabStrip, under
tabs as tall as the MultiPage's font makes them. The font is the one set
on the MultiPage, which its TabStrip keeps, or else the one it shows from
its container. Every other page keeps the layout it was added with,
which is always the default size's, so a MultiPage that VBA sizes after
adding it keeps its second page laid out for the default size.
`add_control` and `add_page` lay pages out the same way. A running form
lays out again each page it shows, and shows that second page whole.

The designers round the TabStrip to whole pixels at 96 DPI and scale
each pixel edge back by the TabStrip's own HIMETRIC per pixel, so the
page's top moves a unit or two with the MultiPage's size. Word differs
from Excel in one case: setting a font on a MultiPage puts its TabStrip
back to the default size and keeps it there whatever size the MultiPage
takes, and the page shown follows the TabStrip.

Setting a MultiPage's size with `set_property` writes the size alone,
where the designers would also size the TabStrip and lay out the page
shown again.

## A form's two captions

A form keeps a caption in its record, which `Designer.Caption` reads, and
one in its designer header, the `\x03VBFrame` text, which the running form
shows and the editor's property sheet edits. `add_form` writes both, and so
does `set_property("Caption", ...)` on a form. `StartUpPosition` and
`ShowModal` are in the header alone.

## Font sizes follow the display

The designers store a form's or Frame's font size in whole pixels of the
display they run on: at 96 DPI in steps of 0.75 point, so 10 point is stored
as 9.75, and at 192 DPI in steps of 0.375. A control added to it takes that
size in twips, cut to a whole twip.
