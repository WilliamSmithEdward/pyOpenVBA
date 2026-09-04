# Macros in an `.accdb`, decoded

> **Status: shipped, 2026-09-03.** All of this is now in
> `pyopenvba.access._macros` and the `macros()` / `create_macro()` /
> `delete_macro()` methods on `AccessDatabase`. It stays here as the
> record of how it was measured -- against macros Access itself created
> through `Application.LoadFromText`.

Access stores a macro as a small binary blob, not as the XML the modern
macro designer shows. What a macro costs the file, measured by diffing a
blank database against the same database with one macro in it:

* three `MSysAccessStorage` rows: a numbered folder under `Scripts`
  (storage id 7), a `Blob` under it holding the macro, and a
  `\x03DirData` under `Scripts` listing the names
* an `MSysObjects` row of type **-32766** under the `Scripts` container,
  carrying an `LvProp` property blob with `PublishToWeb`
* an `MSysNavPaneObjectIDs` row of type **32770**

No `MSysNavPaneGroupToObjects` row, which a module does get.

**Object ids step by one for a macro**, where a module's step by four.
The step is what an object reserves for itself, not a global stride.

The storage folder follows the same rule as a module's: lowest free name
from `chr(0x30 + <rows in the container that are not folders>)`. `Scripts`
starts empty, so the first macro's folder is `0`; `\x03DirData` is created
alongside it, so the second is `1`, and they run on from there.

## The blob

```
00 00 00 00  ff ff ff ff  ff ff ff ff  00 00 02 00     header, 32 bytes
ff*16
<u16 4> "33" UTF-16 <u16 0>                            a length-prefixed
                                                       string, always "33"
```

Then one record per action:

```
<u16 action id>
<u16 row number, 1-based>
<14 u16 argument slots>        ff ff when absent, else a byte offset
                               into the string area
<u16 string-area length>
<strings, each UTF-16 and NUL-terminated>
<u16 0>
```

Arguments occupy slots from **4** upward, one per argument in order; an
argument left empty takes no slot. Slots 0 to 3 are `ff ff` in every
macro measured. A `Beep` costs 36 bytes and a `MsgBox` with four
arguments 62.

Worked example, `MsgBox "Hi", "No", 1, "Title"` as the second action:

```
16 00 02 00                     action 22, row 2
ff ff ff ff ff ff ff ff         slots 0-3, unused
00 00 06 00 0c 00 10 00         slots 4-7: offsets 0, 6, 12, 16
ff ff ... ff ff                 slots 8-13
1c 00                           28 bytes of strings
48 00 69 00 00 00               "Hi"
4e 00 6f 00 00 00               "No"
31 00 00 00                     "1"
54 00 69 00 74 00 6c 00 65 00 00 00   "Title"
00 00
```

## The action ids

Measured by loading one macro per action and pairing each storage folder
with its `MSysObjects` row in id order, which is creation order. `Beep`,
`Echo` and `MsgBox` agree with a separate five-macro fixture where the
correspondence was known independently.

| id | action | id | action |
|---:|---|---:|---|
| 4 | Beep | 33 | RunCode |
| 5 | CancelEvent | 35 | RunSQL |
| 6 | Close | 40 | SetValue |
| 9 | Echo | 41 | SetWarnings |
| 15 | GoToRecord | 44 | StopAllMacros |
| 17 | Hourglass | 45 | StopMacro |
| 22 | MsgBox | 46 | OpenReport |
| 23 | OpenForm | 71 | SingleStep |
| 24 | OpenQuery | 72 | ClearMacroError |
| 25 | OpenTable | 73 | OnError |
| 27 | Quit | 76 | SetTempVar |
| 28 | Requery | 78 | RemoveTempVar |

`MaximizeWindow` is refused by `LoadFromText` on this build, so it has no
id here.

## Getting ground truth

`Application.LoadFromText acMacro, name, path` still accepts the **Access
2003 text format**, which is far easier to write than the modern XML:

```
Version =196611
ColumnsShown =0
Begin
    Action ="MsgBox"
    Argument ="Hi"
    Argument ="No"
    Argument ="1"
    Argument ="Title"
End
```

ANSI or UTF-16, either is accepted. The modern
`<UserInterfaceMacros>` XML was refused with "encountered errors while
importing".

Two traps. The `\x03DirData` order under `Scripts` is **not** the folder
order, so pairing a blob with a name through it gives a shifted table;
pair through `MSysObjects` ids instead. And Access refuses to attach or
import some file types outright -- `bas`, `accdb`, `cab`, `exe`, `iso`
and `msi` for an attachment column -- so a probe over many extensions
has to tolerate failures rather than stopping at the first.

## The catalog row's properties

A macro's `MSysObjects.LvProp` is an ordinary `MR2` property blob holding
one property, `PublishToWeb`, true. The engine's own property writer
reproduces it byte for byte, so nothing new was needed.

## What is left

Nothing for macros. Forms and reports are the part of phase 8 still
unstarted, and they are a different problem: their designs are much
larger blobs on the `MSysObjects` row itself rather than a handful of
action records.
