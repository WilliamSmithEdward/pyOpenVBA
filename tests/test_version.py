"""Guard against version drift between pyproject.toml and the package.

pyopenvba.__version__ went stale once (it reported 2.0.0 while PyPI
shipped 3.0.1); this test pins the two sources together so a release
bump cannot miss one of them again.
"""

from __future__ import annotations

from importlib.metadata import version

import pyopenvba


def test_dunder_version_matches_installed_metadata() -> None:
    assert pyopenvba.__version__ == version("pyOpenVBA")
