"""Move the malware scan's YARA pins to the newest releases they may take.

YARA Forge's newest release is taken as soon as it is out: its rules are
detection content, and the pull request the updater opens is scanned
before anyone merges it. A YARA-X release is taken once it is a week old.
Each pin's SHA-256 is the digest GitHub records for the release asset,
never one worked out from a download.

Usage::

    python scripts/security/yara_update.py .github/security/yara.env

Needs the gh CLI with a token. Rewrites the file when a pin moved, prints
what moved, and writes ``changed`` and ``title`` to $GITHUB_OUTPUT when set.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

FORGE_ASSET = "yara-forge-rules-full.zip"
COOLDOWN = dt.timedelta(days=7)
PIN = re.compile(r"^(?P<key>[A-Z0-9_]+)=(?P<value>.*)$", re.MULTILINE)


def digest(release: dict[str, Any], name: str) -> str:
    """The SHA-256 GitHub records for one asset of `release`."""
    for asset in release.get("assets", []):
        if asset["name"] == name:
            recorded = asset.get("digest") or ""
            if not recorded.startswith("sha256:"):
                raise ValueError(f"GitHub records no SHA-256 for {name} in {release['tag_name']}")
            return recorded.removeprefix("sha256:")
    raise ValueError(f"{release['tag_name']} has no asset {name}")


def forge_pin(release: dict[str, Any]) -> dict[str, str]:
    return {"YARA_FORGE_RELEASE": release["tag_name"], "YARA_FORGE_SHA256": digest(release, FORGE_ASSET)}


def yara_x_pin(releases: list[dict[str, Any]], now: dt.datetime) -> dict[str, str]:
    """The newest full release published at least a week before `now`."""
    for release in releases:
        published = dt.datetime.fromisoformat(release["published_at"].replace("Z", "+00:00"))
        if release.get("draft") or release.get("prerelease") or now - published < COOLDOWN:
            continue
        tag = release["tag_name"]
        return {
            "YARA_X_VERSION": tag,
            "YARA_X_SHA256": digest(release, f"yara-x-{tag}-x86_64-unknown-linux-gnu.tar.gz"),
        }
    raise ValueError("no YARA-X release is a week old yet")


def read_pins(text: str) -> dict[str, str]:
    return {match["key"]: match["value"] for match in PIN.finditer(text)}


def with_pins(text: str, pins: dict[str, str]) -> str:
    """`text` with each pin's value replaced, comments and order kept."""
    missing = set(pins) - set(read_pins(text))
    if missing:
        raise ValueError(f"the pin file has no {sorted(missing)}")
    return PIN.sub(lambda m: f"{m['key']}={pins.get(m['key'], m['value'])}", text)


def moves(old: dict[str, str], new: dict[str, str]) -> list[str]:
    """What moved, one line per release."""
    lines: list[str] = []
    for name, key in (("YARA Forge", "YARA_FORGE_RELEASE"), ("YARA-X", "YARA_X_VERSION")):
        if old.get(key) != new[key]:
            lines.append(f"{name} {old.get(key)} -> {new[key]}")
    return lines


def gh_api(path: str) -> Any:
    out = subprocess.run(["gh", "api", path], capture_output=True, text=True, check=True).stdout
    return json.loads(out)


def main(argv: list[str] | None = None) -> int:
    target = Path((argv or sys.argv[1:])[0])
    text = target.read_text(encoding="utf-8")
    old = read_pins(text)
    new = {
        **forge_pin(gh_api("repos/YARAHQ/yara-forge/releases/latest")),
        **yara_x_pin(gh_api("repos/VirusTotal/yara-x/releases?per_page=20"), dt.datetime.now(dt.timezone.utc)),
    }
    moved = moves(old, new)
    if moved:
        target.write_text(with_pins(text, new), encoding="utf-8", newline="\n")
    print("\n".join(moved) or "The YARA pins are current.")
    output = os.environ.get("GITHUB_OUTPUT")
    if output:
        with open(output, "a", encoding="utf-8") as handle:
            handle.write(f"changed={'true' if moved else 'false'}\n")
            handle.write(f"title=ci: {', '.join(moved)}\n" if moved else "title=\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
