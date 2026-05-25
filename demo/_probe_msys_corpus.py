"""Verify the MSysObjects reader across every corpus sample."""
from pathlib import Path
from pyopenvba.access import AccessFile, MSYS_TYPE_MODULE

CORPUS = Path("tests/live_access_test/re_corpus/samples")
samples = sorted(CORPUS.glob("*.accdb"))

all_ok = True
print(f"Testing {len(samples)} corpus samples...\n")
for sample in samples:
    db = AccessFile(sample)
    objs = db.msys_objects()
    mods = list(db.iter_msys_modules())
    src_names = set(db.vba_module_names())
    msys_mod_names = {m.name for m in mods}

    matches = src_names == msys_mod_names
    status = "OK" if matches else "MISMATCH"
    all_ok = all_ok and matches
    print(
        f"  {sample.name:<60s} rows={len(objs):3d}  "
        f"vba_mods={len(src_names)}  msys_mods={len(mods)}  {status}"
    )
    if not matches:
        print(f"    src_names    = {sorted(src_names)}")
        print(f"    msys_names   = {sorted(msys_mod_names)}")

print()
print("ALL PASS" if all_ok else "SOME SAMPLES MISMATCH")
