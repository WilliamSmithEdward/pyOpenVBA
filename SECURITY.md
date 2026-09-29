# Security policy

## Reporting a vulnerability

Please report a vulnerability privately, not in a public issue. Use
[Report a vulnerability](https://github.com/WilliamSmithEdward/pyOpenVBA/security/advisories/new)
on the repository's Security tab. It opens a draft advisory that only
you and the maintainer can see.

A useful report names the pyOpenVBA version, the Python version and the
operating system, and includes a file or a short script that shows the
problem. A crafted file is the most direct proof, so attach a minimal
one if you can.

## Supported versions

Only the latest release on PyPI is supported. A fix ships in a new
release, and earlier versions do not get one.

## What to report

pyOpenVBA reads and writes Office files that may come from anyone, so
the input to worry about is a file. For example:

- a file that makes the library hang, run out of memory or crash the
  interpreter;
- a file whose reading writes anything, or a save that touches a file
  other than the one it was asked to write;
- a way past the safety guards, such as a save that changes a
  password-protected project without `allow_protected=True`.

The VBA runtime is not a sandbox. It never starts a COM server, runs a
command or opens a network connection, and neither does the Power Query
evaluator. A macro can still open and save files through the object
model, `Workbooks.Open` and `SaveAs` among them, with the permissions of
the Python process. Run macros you do not trust only where that is
acceptable.

## How the code is checked

The package has no runtime dependencies. CodeQL and Semgrep scan it,
with the workflows that build and publish it, on every push to main,
every pull request, every week, and before every release. A finding
fails the scan unless
[.github/security/accepted.toml](.github/security/accepted.toml) lists it
with the reason it is accepted, and an entry there that no longer
matches fails it too.

A release is published only after its commit passes the scan, and it
carries the report as `pyopenvba-<version>-security-report.md`, beside
the SARIF the report was made from. Every action the workflows use is
pinned to a commit, and Dependabot proposes updates to the actions and
to the pinned Semgrep.
