"""inject_xls_demo.py — inject Module1 into an .xls (legacy) workbook."""
import shutil
from pathlib import Path
from pyopenvba import ExcelFile

SRC = Path(__file__).parent.parent / "tests" / "live_excel_testing" / "xls_test.xls"
OUT = Path(__file__).parent / "output" / "xls_test_injected.xls"

NEW_SOURCE = (
    "Option Explicit\r\n"
    "\r\n"
    "Sub RunDemo()\r\n"
    "    MsgBox \"Hello from pyOpenVBA! (.xls)\", vbInformation, \"pyOpenVBA\"\r\n"
    "End Sub\r\n"
)


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SRC, OUT)

    with ExcelFile(OUT) as wb:
        print("before:", wb.module_names())
        wb.set_module("Module1", NEW_SOURCE)
        wb.save()

    with ExcelFile(OUT) as wb:
        names = wb.module_names()
        mod1 = wb.get_module("Module1")

    print("after :", names)
    print("Module1 source:")
    print(mod1)
    print("Written:", OUT)


if __name__ == "__main__":
    main()
