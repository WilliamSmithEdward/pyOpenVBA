# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/pyOpenVBA/security/advisories/new).
Only the maintainer sees it. Include the pyOpenVBA version, the Python
version and the operating system, and the smallest file or steps that
show it, with credentials and private data removed. A crafted file is the
most direct proof, so attach a minimal one if you can.

A confirmed vulnerability is fixed in a release on PyPI, and the advisory is
published with it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on PyPI receives security fixes. Older
releases are not maintained separately; update when a fix ships.

## Scope

pyOpenVBA reads and writes Office files that may come from anyone, so the
input to worry about is a file. These count as vulnerabilities:

- a file that makes the library hang, run out of memory or crash the
  interpreter;
- a file whose reading writes anything, or a save that touches a file
  other than the one it was asked to write;
- a way past the safety guards, such as a save that changes a
  password-protected project without `allow_protected=True`.

The package has no runtime dependencies. It never starts a COM server,
runs a command or opens a network connection, and neither do its VBA
runtime and Power Query evaluator.

### Running macros

The VBA runtime is not a sandbox. A macro can still open and save files
through the object model, `Workbooks.Open` and `SaveAs` among them, with
the permissions of the Python process. Run macros you do not trust only
where that is acceptable.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security also runs weekly, so new queries and rules reach code
that has not changed, and Malware scan runs daily, so new signatures and
rules reach files that have not changed.

- **Code:** CodeQL with GitHub's security-extended queries, for Python and
  GitHub Actions, and Semgrep with the default, Python, security-audit,
  secrets and GitHub Actions rule sets. Both scan the package, the
  workflows that build and publish it, and the scripts in
  `scripts/security` that judge the scans. A `nosemgrep` comment cannot
  hide a finding. Results go to the repository's code scanning.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** there is no dependency audit, because the package has
  no runtime dependencies. The tools the workflows install come from
  hash-locked files (see Pinning and updates).
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan every file the commit holds, test fixtures and the
  fuzz corpus included, and the wheel and sdist built from it with the
  hash-locked build tools, as a release builds them. YARA-X runs YARA
  Forge's full rule set, which gathers the public YARA rule collections
  into one. A scan error fails the report as a match does.
- **Fuzzing:** Atheris drives the parsers that read untrusted input, seven
  targets in `fuzz/fuzz_parsers.py`: the compound file reader, VBA
  decompression, the `dir`, `PROJECT` and `PROJECTwm` streams, Power
  Query's mashup container and the M formula parser. Each target starts
  from its seeds in `tests/fuzz_corpus`, and a parser must succeed or
  raise one of the errors `tests/fuzz_corpus/README.md` lists for it.
  The Fuzz workflow runs on every change to the package, the fuzz targets
  or the corpus, and daily. It is not a gate: a finding becomes a
  regression test with its fix, a seed that `tests/test_gates.py` replays
  on every CI run.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Its Code-Review and Contributors checks assume more than one
  maintainer, such as a second person approving every change, so a
  single-maintainer project cannot score full marks on them.

## Accepted findings

A finding is fixed, or accepted with a written reason in
[.github/security/accepted.toml](.github/security/accepted.toml) for
CodeQL and Semgrep, or
[.github/security/malware-accepted.toml](.github/security/malware-accepted.toml)
for ClamAV and YARA-X. An entry matches on the tool, the rule and the
file, and in accepted.toml also the text of the flagged line, so an edited
line needs another review, and an entry that no longer matches fails the
report. zizmor keeps its exceptions in `.github/zizmor.yml` or inline
beside the line they excuse, each with its reason.

The current entries:

- Semgrep `use-defused-xml`, in four modules that parse an XML part with
  the standard library's ElementTree, which loads no external entity or
  DTD; defusedxml would be the package's first dependency.
- Semgrep `insecure-hash-algorithm-sha1`, in `compute_v3_content_hash()`,
  which fingerprints module source to detect change and protects nothing.
- A Semgrep parse notice for `excel-macro-test.yml`, whose PowerShell step
  two GitHub Actions rules read as Bash.
- YARA-X `ARKBIRD_SOLG_TA505_Maldoc_21Nov_2`, on binary `.xls` and `.doc`
  fixtures and fuzz seeds made from them, which match on the reference
  paths every Office VBA project records.
- YARA-X `SIGNATURE_BASE_Powershell_Case_Anomaly`, on the README and two
  test files that name PowerShell in ordinary prose and commands.
- zizmor's `self-repository` and `superfluous-actions` rules are turned
  off in `.github/zizmor.yml`, each with its reason and when it comes back.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, and the YARA-X engine and YARA Forge rules to a
release and its SHA-256. The package has no dependencies of its own.
ClamAV's signatures change too often to pin, so freshclam fetches and
verifies them on every run.

Dependabot proposes updates to the GitHub Actions, the Semgrep and ClamAV
images, and the hash-locked tools in `.github/requirements` once a version
is a week old, and at once for a security advisory. The Update YARA rules
workflow proposes new YARA pins in `.github/security/yara.json` each week.
A minor or patch update, and the YARA pull request, merges itself once CI,
Security and Malware scan pass; a third-party major version waits for
review.

## Releases

A pushed `v*.*.*` tag builds the sdist and wheel, checks that the tag
matches the version in `pyproject.toml`, and runs Security and Malware
scan on the tagged commit. Nothing is published unless both pass. The
distributions go to PyPI through Trusted Publishing, so no upload token
exists to leak. The GitHub release carries the distributions,
`pyopenvba-<version>-security-report.md` and
`pyopenvba-<version>-malware-report.md` beside the scan results they were
made from, and the provenance bundle. Started by hand, the Publish
workflow is always a dry run and publishes nothing.

### Verifying a download

Every file on PyPI carries PyPI's own provenance, which names this
repository's `publish.yml` as the publisher; the file's page on PyPI shows it.
Releases published after 2026-09-30 also carry a GitHub build provenance
attestation, which you can check against any copy of the file, from PyPI or
from the GitHub release:

```
pip download pyopenvba --no-deps -d check
gh attestation verify check/<file> --owner WilliamSmithEdward
```

The output names the commit and workflow run that built the file. The
signed bundle is also attached to the GitHub release as
`pyopenvba-<version>.sigstore.json`, so the check works without asking
GitHub for it: add `--bundle pyopenvba-<version>.sigstore.json`.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
