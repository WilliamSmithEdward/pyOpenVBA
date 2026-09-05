"""Build a demo Access database: a form whose buttons drive VBA.

The form ``Calculator`` has text boxes for a quantity and a unit price, a
check box for express delivery, two buttons and a list of the lines
added.  Its code-behind calls ``Pricing.Extended`` in a standard module,
which applies a quantity discount, and keeps the running total in an
instance of the class module ``Basket``.  Everything is written with
pyOpenVBA and nothing else: the controls, their fonts and colours, the
code.  Access compiles the project the first time it opens the file.

    python examples/access_form_demo.py [path.accdb]

Open the result in Access: the form opens with it (the database's
``StartUpForm`` property).  Type a quantity and a price and press "Add
line".  ``tests/test_live_access_design_gate.py`` builds the same database
and has Access drive the form.
"""

from __future__ import annotations

import sys
from pathlib import Path

from pyopenvba import AccessDatabase, VBAModuleKind
from pyopenvba.access import AccessForm

PRICING = """Option Compare Database
Option Explicit

' Quantity discounts: 5% from five units, 10% from ten.
Public Function DiscountRate(ByVal qty As Long) As Double
    If qty >= 10 Then
        DiscountRate = 0.1
    ElseIf qty >= 5 Then
        DiscountRate = 0.05
    Else
        DiscountRate = 0
    End If
End Function

Public Function Extended(ByVal qty As Long, ByVal price As Currency, ByVal express As Boolean) As Currency
    Extended = qty * price * (1 - DiscountRate(qty))
    If express Then Extended = Extended + 4.5
End Function
"""

BASKET = """Option Compare Database
Option Explicit

Private mTotal As Currency
Private mLines As Long

Public Property Get Total() As Currency
    Total = mTotal
End Property

Public Property Get Lines() As Long
    Lines = mLines
End Property

Public Sub Add(ByVal amount As Currency)
    mTotal = mTotal + amount
    mLines = mLines + 1
End Sub

Public Sub Clear()
    mTotal = 0
    mLines = 0
End Sub
"""

FORM_CODE = """Option Compare Database
Option Explicit

Private basket As New Basket

Private Sub AddLine_Click()
    AddCurrent
End Sub

Private Sub Reset_Click()
    basket.Clear
    Me.Lines.RowSource = ""
    Me.TotalLabel.Caption = "Total: " & Format(0, "Currency")
End Sub

' Adds the quantity and price on the form as one line; returns the total.
Public Function AddCurrent() As Currency
    Dim amount As Currency
    amount = Pricing.Extended(CLng(Nz(Me.Qty, 0)), CCur(Nz(Me.Price, 0)), CBool(Nz(Me.Express, False)))
    basket.Add amount
    Me.Lines.AddItem Me.Qty & " x " & Format(Me.Price, "Currency") & " = " & Format(amount, "Currency")
    Me.TotalLabel.Caption = "Total: " & Format(basket.Total, "Currency") & " (" & basket.Lines & " lines)"
    AddCurrent = basket.Total
End Function
"""


def rgb(red: int, green: int, blue: int) -> int:
    """A colour as Access stores it: red in the low byte, blue in the high."""
    return red | (green << 8) | (blue << 16)


# One accent, a neutral base: the palette the form is drawn with.
ACCENT = rgb(31, 78, 121)
ACCENT_HOVER = rgb(46, 117, 182)
ACCENT_PRESSED = rgb(21, 54, 84)
WHITE = rgb(255, 255, 255)
INK = rgb(38, 38, 38)
MUTED = rgb(89, 89, 89)
ON_ACCENT_MUTED = rgb(220, 230, 242)
HAIRLINE = rgb(191, 191, 191)
HOVER_GREY = rgb(242, 242, 242)
PRESSED_GREY = rgb(217, 217, 217)
FONT = "Segoe UI"
#: TextAlign values.
LEFT, RIGHT = 1, 3
#: BackStyle values.
TRANSPARENT, OPAQUE = 0, 1
#: OldBorderStyle (the property sheet's Border Style) values.
NO_BORDER, SOLID = 0, 1
#: SpecialEffect values.
FLAT = 0

# Twips: 1440 to the inch.  The form is five inches wide.
WIDTH = 7200
MARGIN = 360
INNER = WIDTH - 2 * MARGIN


def text(form: AccessForm, name: str, size: int, colour: int, *, bold: bool = False, align: int = LEFT) -> None:
    """Typography for a label, text box or list box."""
    control = form.control(name)
    control.set_property("FontName", FONT)
    control.set_property("FontSize", size)
    control.set_property("FontWeight", 700 if bold else 400)
    control.set_property("ForeColor", colour)
    if control.kind != "ListBox":
        control.set_property("TextAlign", align)


def build(path: str | Path) -> Path:
    """Write the demo database at ``path`` and return the path."""
    with AccessDatabase.create_new(path) as db:
        db.add_module("Pricing", PRICING)
        db.add_module("Basket", BASKET, kind=VBAModuleKind.other)

        form = db.add_form("Calculator", caption="Order calculator", width=WIDTH, height=5640)

        # A band of the accent colour across the top carries the title.
        form.add_control("Rectangle", "Banner", left=0, top=0, width=WIDTH, height=960)
        banner = form.control("Banner")
        banner.set_property("BackStyle", OPAQUE)
        banner.set_property("BackColor", ACCENT)
        banner.set_property("OldBorderStyle", NO_BORDER)
        banner.set_property("SpecialEffect", FLAT)
        form.add_control("Label", "Title", left=MARGIN, top=180, width=INNER, height=440, caption="Order calculator")
        form.add_control(
            "Label",
            "Subtitle",
            left=MARGIN,
            top=620,
            width=INNER,
            height=260,
            caption="Five units earn 5% off, ten earn 10%; express delivery adds 4.50.",
        )
        for name in ("Title", "Subtitle"):
            form.control(name).set_property("BackStyle", TRANSPARENT)
        text(form, "Title", 16, WHITE, bold=True)
        text(form, "Subtitle", 9, ON_ACCENT_MUTED)

        # Two inputs side by side, a check box beside them.
        form.add_control("Label", "QtyLabel", left=MARGIN, top=1200, width=1600, height=280, caption="Quantity")
        form.add_control("TextBox", "Qty", left=MARGIN, top=1500, width=1600, height=420)
        form.add_control("Label", "PriceLabel", left=2160, top=1200, width=1600, height=280, caption="Unit price")
        form.add_control("TextBox", "Price", left=2160, top=1500, width=1600, height=420)
        form.add_control("CheckBox", "Express", left=4080, top=1580, width=260, height=260)
        form.add_control(
            "Label", "ExpressLabel", left=4400, top=1560, width=2440, height=300, caption="Express delivery (+4.50)"
        )
        for name in ("QtyLabel", "PriceLabel", "ExpressLabel"):
            text(form, name, 10, MUTED)
        for name in ("Qty", "Price"):
            box = form.control(name)
            text(form, name, 12, INK, align=RIGHT)
            box.set_property("BorderColor", HAIRLINE)
            box.set_property("SpecialEffect", FLAT)
        form.control("Qty").set_property("DefaultValue", "1")
        form.control("Price").set_property("Format", "Currency")
        form.control("Express").set_property("DefaultValue", "False")

        # A filled primary button and an outlined secondary one.  A colour
        # set here lands with the records Access writes beside it (the theme
        # index turned off, a button's gradient off), so the properties read
        # as they would in the property sheet.  UseTheme stays on: turned
        # off, Access draws the Windows button instead and ignores the fill.
        form.add_control("CommandButton", "AddLine", left=MARGIN, top=2160, width=1800, height=480, caption="Add line")
        form.add_control("CommandButton", "Reset", left=2280, top=2160, width=1480, height=480, caption="Reset")
        for name, back, fore, border, hover, pressed in (
            ("AddLine", ACCENT, WHITE, ACCENT, ACCENT_HOVER, ACCENT_PRESSED),
            ("Reset", WHITE, ACCENT, HAIRLINE, HOVER_GREY, PRESSED_GREY),
        ):
            button = form.control(name)
            for colour, value in (
                ("Back", back),
                ("Border", border),
                ("Fore", fore),
                ("Hover", hover),
                ("Pressed", pressed),
                ("HoverFore", fore),
                ("PressedFore", fore),
            ):
                button.set_property(f"{colour}Color", value)
            button.set_property("FontName", FONT)
            if name == "AddLine":
                button.set_property("FontWeight", 700)
            button.set_property("OnClick", "[Event Procedure]")

        # The lines added so far, under a hairline rule, and the total.
        form.add_control("Line", "Rule", left=MARGIN, top=2880, width=INNER, height=0)
        rule = form.control("Rule")
        rule.set_property("BorderColor", PRESSED_GREY)
        form.add_control("Label", "LinesLabel", left=MARGIN, top=3000, width=INNER, height=280, caption="Lines")
        text(form, "LinesLabel", 10, MUTED)
        form.add_control("ListBox", "Lines", left=MARGIN, top=3300, width=INNER, height=1560)
        lines = form.control("Lines")
        text(form, "Lines", 10, INK)
        lines.set_property("BorderColor", HAIRLINE)
        lines.set_property("SpecialEffect", FLAT)
        lines.set_property("RowSourceType", "Value List")
        form.add_control("Label", "TotalLabel", left=MARGIN, top=5040, width=INNER, height=440, caption="Total: $0.00")
        text(form, "TotalLabel", 14, ACCENT, bold=True, align=RIGHT)

        form.set_code(FORM_CODE)
        # Access opens this form with the database.
        db.set_database_properties({"StartUpForm": "Calculator"})
        return db.save()


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "access_form_demo.accdb"))
