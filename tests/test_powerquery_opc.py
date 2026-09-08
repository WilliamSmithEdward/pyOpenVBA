"""The package container, and the one entry Excel will not tolerate.

Excel's own file recovery parks the parts it threw out under `[trash]`,
and an OPC part name cannot open a segment with a bracket.  A workbook
carrying one does not open in Excel at all -- measured both ways against
the same fixture, which opens before the entry is added and raises after
-- so preserving it faithfully would hand back a file that stays broken.
"""

from __future__ import annotations

import shutil
import warnings
import zipfile
from pathlib import Path

import pytest

from pyopenvba.powerquery import PowerQueryWorkbook
from pyopenvba.powerquery._opc import OpcFile

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"
TRASH = "[trash]/0000.dat"


def recovered(source: Path, target: Path, *names: str) -> Path:
    """A copy of `source` carrying the entries Excel's recovery leaves."""
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            out.writestr(info, src.read(info.filename))
        for name in names or (TRASH,):
            out.writestr(name, bytes(64))
    return target


def test_the_container_reads_a_package_that_carries_one(tmp_path: Path) -> None:
    path = recovered(FIXTURES / "three_queries.xlsx", tmp_path / "read.xlsx")
    assert TRASH in OpcFile.parse(path.read_bytes()).names()


def test_dropping_says_what_it_took(tmp_path: Path) -> None:
    path = recovered(FIXTURES / "three_queries.xlsx", tmp_path / "drop.xlsx")
    package = OpcFile.parse(path.read_bytes())

    assert package.drop_reserved() == [TRASH]
    assert TRASH not in package.names()
    assert package.drop_reserved() == []


def test_dropping_leaves_every_real_part_alone(tmp_path: Path) -> None:
    original = OpcFile.parse((FIXTURES / "three_queries.xlsx").read_bytes()).names()
    path = recovered(FIXTURES / "three_queries.xlsx", tmp_path / "keep.xlsx")
    package = OpcFile.parse(path.read_bytes())
    package.drop_reserved()

    assert package.names() == original


def test_saving_drops_them_and_warns(tmp_path: Path) -> None:
    path = recovered(FIXTURES / "three_queries.xlsx", tmp_path / "saved.xlsx")

    with pytest.warns(UserWarning, match=r"\[trash\]"):
        PowerQueryWorkbook(path).save()

    with zipfile.ZipFile(path) as package:
        assert not [name for name in package.namelist() if name.startswith("[trash]/")]
    assert PowerQueryWorkbook(path).query_names() == ["Numbers", "Doubled", "Count Of Rows"]


def test_more_than_one_is_reported_together(tmp_path: Path) -> None:
    path = recovered(
        FIXTURES / "three_queries.xlsx",
        tmp_path / "several.xlsx",
        "[trash]/0000.dat",
        "[trash]/0001.dat",
    )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        PowerQueryWorkbook(path).save()

    assert len(caught) == 1
    assert "0000.dat" in str(caught[0].message) and "0001.dat" in str(caught[0].message)


def test_an_ordinary_workbook_warns_about_nothing(tmp_path: Path) -> None:
    out = tmp_path / "clean.xlsx"
    shutil.copyfile(FIXTURES / "three_queries.xlsx", out)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        PowerQueryWorkbook(out).save()

    assert caught == []
