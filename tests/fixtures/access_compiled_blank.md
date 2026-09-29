# Historical compiled Access fixture

`access_compiled_blank.accdb` is the pre-6.3.3 `blank_database.accdb`.
It intentionally retains stale test-comment bytes in its compiled module
cache and unused pages. It is test evidence, not a template for distribution.

The Access source-replacement tests require a compiled input; the allocation
and compaction tests use the exact layout measured with DAO on this file.
The template regression also demonstrates why rewriting the module alone
is insufficient and rebuilding the database removes the old bytes.

Neither the wheel nor the source distribution includes this fixture.
