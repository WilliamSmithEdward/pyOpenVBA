# pyOpenVBA Demo

A tiny end-to-end demo: pyOpenVBA writes a real VBA module into a real
Excel workbook, without opening Excel.

## Files

- `test_macro_workbook.xlsm` -- the workbook the script writes into.
- `push_demo_module.py` -- adds a `DemoShowcase` module with three macros:
  `RainbowGrid`, `FibSeries`, `BubbleSortRace`.

## Run it

Install pyOpenVBA (one of):

```powershell
pip install pyOpenVBA
pip install -e .   # from a clone of this repo
```

From this folder:

```powershell
python push_demo_module.py
```

Then open `test_macro_workbook.xlsm` in Excel, press `Alt+F11`, and run any
of the three subs from the VBA editor (`F5`).
