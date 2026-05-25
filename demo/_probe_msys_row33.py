"""Parse row 33 of MSysObjects (the 'M' module) bit-by-bit to verify
the Jet 4 row layout, including null bitmap and variable-column jump
table."""
from pathlib import Path

PAGE = 4096
data = Path("tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb").read_bytes()
pg = data[17 * PAGE : 18 * PAGE]
row = pg[0x032C : 0x032C + 65]

print("Row 33 (M) raw bytes:")
for i in range(0, len(row), 16):
    chunk = row[i : i + 16]
    h = " ".join(f"{b:02x}" for b in chunk)
    a = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
    print(f"  {i:02x}  {h:<48}  {a}")

print()
print(f"col_count    = {int.from_bytes(row[0:2], 'little')}")
print(f"Id           = 0x{int.from_bytes(row[2:6], 'little'):08x}")
print(f"ParentId     = 0x{int.from_bytes(row[6:10], 'little'):08x}")
print(f"Type         = {int.from_bytes(row[10:12], 'little', signed=True)}")
print(f"Flags        = 0x{int.from_bytes(row[28:32], 'little'):08x}")

v = row[32:]
print(f"Var section  = bytes [32..{len(row)})  (len={len(v)})")
print(f"Var hex:     {v.hex()}")

null_mask = v[-3:]
var_col_count = int.from_bytes(v[-5:-3], "little")
print(f"Null mask    = {null_mask.hex()}")
present = [bool(null_mask[b // 8] & (1 << (b % 8))) for b in range(17)]
print(f"Present bits = {present}")
print(f"Var col count= {var_col_count}")

# Jump table: var_col_count u16 entries (each = start offset of that var col)
# placed BEFORE the var_col_count u16 word.
jt_start = len(v) - 5 - 2 * var_col_count
print(f"Jump table @ var offset {jt_start} (row offset {32 + jt_start}):")
for i in range(var_col_count):
    off = int.from_bytes(v[jt_start + 2 * i : jt_start + 2 * i + 2], "little")
    print(f"  var[{i}] start_off = 0x{off:04x} = {off}")
