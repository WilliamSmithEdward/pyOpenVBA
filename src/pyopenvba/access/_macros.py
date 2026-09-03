"""Macros.

Access stores a macro as a small binary blob, not as the XML its designer
shows.  The blob lives in `MSysAccessStorage` as `Scripts/<folder>/Blob`
and holds one record per action::

    00 00 00 00  ff*8  00 00 02 00  ff*16     a 32-byte header
    <u16 4> "33" UTF-16 <u16 0>               always "33"
    per action:
      <u16 action id> <u16 row number, 1-based>
      <14 u16 argument slots>                 ff ff when absent, else a
                                              byte offset into the strings
      <u16 string-area length>
      <strings, each UTF-16 and NUL-terminated>
      <u16 0>

Arguments occupy slots from **4** upward, one per argument in order; an
argument left empty takes no slot, so a gap in the middle reads back as
an empty string.  Slots 0 to 3 are `ff ff` in every macro measured.

The action ids were measured by loading one macro per action through
`Application.LoadFromText` and pairing each storage folder with its
`MSysObjects` row in id order, which is creation order.  See
`docs/research/access_macros/README.md`.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pyopenvba.access_read import AccessError

#: `<u32 0>`, eight `ff`, `00 00 02 00`, sixteen `ff`.
HEADER = bytes.fromhex("00000000") + b"\xff" * 8 + bytes.fromhex("00000200") + b"\xff" * 16
#: A length-prefixed string that reads `33` in every macro measured.
PREAMBLE = (4).to_bytes(2, "little") + "33".encode("utf-16-le") + (0).to_bytes(2, "little")
#: Fourteen `u16` slots per action, the first four never used.
SLOTS = 14
FIRST_ARGUMENT = 4
EMPTY = 0xFFFF
MAX_ARGUMENTS = SLOTS - FIRST_ARGUMENT

#: Action name to the id the blob carries.  Measured, not guessed; every
#: one of these was created by Access and read back.
ACTION_IDS: dict[str, int] = {
    "Beep": 4,
    "CancelEvent": 5,
    "Close": 6,
    "Echo": 9,
    "GoToRecord": 15,
    "Hourglass": 17,
    "MsgBox": 22,
    "OpenForm": 23,
    "OpenQuery": 24,
    "OpenTable": 25,
    "Quit": 27,
    "Requery": 28,
    "RunCode": 33,
    "RunSQL": 35,
    "SetValue": 40,
    "SetWarnings": 41,
    "StopAllMacros": 44,
    "StopMacro": 45,
    "OpenReport": 46,
    "SingleStep": 71,
    "ClearMacroError": 72,
    "OnError": 73,
    "SetTempVar": 76,
    "RemoveTempVar": 78,
}
ACTION_NAMES: dict[int, str] = {value: key for key, value in ACTION_IDS.items()}

#: `MSysObjects.Type` and `MSysNavPaneObjectIDs.Type` for a macro.  A
#: macro's object id steps by one where a module's steps by four.
OBJECT_MACRO = -32766
NAV_MACRO_TYPE = 32770
OBJECT_ID_STEP = 1


@dataclass(frozen=True)
class MacroAction:
    """One action and its arguments, in the order the designer shows."""

    name: str
    arguments: tuple[str, ...] = ()

    @property
    def action_id(self) -> int:
        try:
            return ACTION_IDS[self.name]
        except KeyError:
            raise AccessError(
                f"unknown macro action {self.name!r}; known: {', '.join(sorted(ACTION_IDS))}"
            ) from None


@dataclass(frozen=True)
class Macro:
    """A macro and the actions it runs."""

    name: str
    actions: tuple[MacroAction, ...] = field(default_factory=tuple)


def parse_macro(blob: bytes) -> tuple[MacroAction, ...]:
    """The actions a macro blob holds."""
    at = len(HEADER) + len(PREAMBLE)
    if blob[: len(HEADER)] != HEADER:
        raise AccessError("the macro blob does not start with its header")
    out: list[MacroAction] = []
    while at + 4 + 2 * SLOTS + 2 <= len(blob):
        action = int.from_bytes(blob[at : at + 2], "little")
        slots = [
            int.from_bytes(blob[at + 4 + 2 * i : at + 6 + 2 * i], "little") for i in range(SLOTS)
        ]
        strings_at = at + 4 + 2 * SLOTS
        length = int.from_bytes(blob[strings_at : strings_at + 2], "little")
        area = blob[strings_at + 2 : strings_at + 2 + length]
        used = [slot for slot in slots[FIRST_ARGUMENT:] if slot != EMPTY]
        arguments: list[str] = []
        for slot in slots[FIRST_ARGUMENT:]:
            if slot == EMPTY:
                arguments.append("")
                continue
            # The terminator has to be looked for two bytes at a time: a
            # character can hold a zero byte of its own.
            stop = slot
            while stop + 2 <= len(area) and area[stop : stop + 2] != b"\x00\x00":
                stop += 2
            arguments.append(area[slot:stop].decode("utf-16-le"))
        while arguments and arguments[-1] == "":
            arguments.pop()
        out.append(
            MacroAction(ACTION_NAMES.get(action, f"Action{action}"), tuple(arguments))
        )
        at = strings_at + 2 + length + 2
        if not used and length == 0 and at >= len(blob):
            break
    return tuple(out)


def build_macro(actions: tuple[MacroAction, ...]) -> bytes:
    """The blob for a macro that runs `actions`."""
    out = bytearray(HEADER + PREAMBLE)
    for number, action in enumerate(actions, start=1):
        if len(action.arguments) > MAX_ARGUMENTS:
            raise AccessError(
                f"an action takes at most {MAX_ARGUMENTS} arguments, not {len(action.arguments)}"
            )
        slots = [EMPTY] * SLOTS
        area = bytearray()
        for position, argument in enumerate(action.arguments):
            if not argument:
                continue
            slots[FIRST_ARGUMENT + position] = len(area)
            area += argument.encode("utf-16-le") + b"\x00\x00"
        out += action.action_id.to_bytes(2, "little") + number.to_bytes(2, "little")
        for slot in slots:
            out += slot.to_bytes(2, "little")
        out += len(area).to_bytes(2, "little") + area + (0).to_bytes(2, "little")
    return bytes(out)
