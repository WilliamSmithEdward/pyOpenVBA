from pyopenvba.access import AccessFile, MSYS_TYPE_MODULE

db = AccessFile("tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb")
objs = db.msys_objects()
print(f"Found {len(objs)} MSysObjects rows")
for o in objs:
    print(
        f"  page={o.page} slot={o.slot:2d}  Id=0x{o.id_:08x}  "
        f"ParentId=0x{o.parent_id:08x}  Type={o.type_:6d}  name={o.name!r}"
    )

print()
mod = db.find_msys_module("M")
print(f'find_msys_module("M") = {mod}')

print()
mods = list(db.iter_msys_modules())
print(f"iter_msys_modules returned {len(mods)} VBA modules")
for m in mods:
    print(f"  {m.name!r}  id=0x{m.id_:08x}  parent=0x{m.parent_id:08x}")
