"""No live probe names a cryptographic provider, or touches encryption options, rights, signing, certificates, sensitivity labels or the work identity. Naming
a provider can reach hardware: on 2026-11-07 a probe that named "Microsoft Base Smart Card Crypto Provider" through
SetPasswordEncryptionOptions raised Windows Security popups asking for a smartcard. The model's answers for these members come from the files
it reads and from fixtures already measured (tests/fixtures/workbook_encryption_options.json), never from a new live run. A script's
docstring may name them, to say so; its code may not."""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

SCRIPTS = sorted((Path(__file__).parents[1] / "scripts").glob("measure_*.py"))
#: The members, each reached with a dot or named alone in a list a probe walks.
BANNED = ("SetPasswordEncryptionOptions", "Permission", "Signatures", "SignatureSet", "SignatureLine", "AddSignatureLine", "Certificate",
          "SensitivityLabel", "SensitivityLabelInitInfo", "IRM", "WorkIdentity")
#: What names a provider in a probe's text.
_PROVIDER = re.compile(r"(?i)crypt\w*\s+provider|\bcsp\b")
#: Setting the provider a book encrypts with names one. Reading it does not. WorksheetFunction.Sign is arithmetic, so signing is caught by
#: Signatures and SignatureSet instead of by a bare .Sign.
_MEMBER = re.compile(r"\.EncryptionProvider\s*=")
_NAMED = re.compile(r"(?<![\w])(?:" + "|".join(BANNED) + r")(?![\w])")


def _code_strings(path: Path) -> list[str]:
    """Every string in a script's code, which is where a probe's VBA is; not its docstring, or a comment."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    docstring = tree.body[0].value if tree.body and isinstance(tree.body[0], ast.Expr) else None
    return [node.value for node in ast.walk(tree) if isinstance(node, ast.Constant) and isinstance(node.value, str) and node is not docstring]


def test_there_are_probes_to_check() -> None:
    assert len(SCRIPTS) > 100


@pytest.mark.parametrize("path", SCRIPTS, ids=lambda path: path.stem)
def test_a_probe_names_no_crypto_member(path: Path) -> None:
    strings = _code_strings(path)
    found = sorted({hit for text in strings for hit in (*_NAMED.findall(text), *_MEMBER.findall(text), *_PROVIDER.findall(text))})
    assert found == []


@pytest.mark.parametrize(("code", "caught"), [
    ('x = "wb.SetPasswordEncryptionOptions p"', True), ('x = ["Permission", "Build"]', True), ('x = "TypeName(wb.Signatures)"', True),
    ('x = \'"Microsoft Base Smart Card Crypto Provider"\'', True), ('x = "Microsoft Enhanced RSA and AES Cryptographic Provider"', True),
    ('x = "wb.Signatures(1).Sign"', True), ('x = "f.Sign(-3)"', False), ('x = "wb.EncryptionProvider = p"', True), ('x = "r = wb.SensitivityLabel"', True),
    ('x = "r = wb.EncryptionProvider"', False), ('x = "r = wb.PasswordEncryptionKeyLength"', False), ('x = "ws.Range(\\"A1\\").Value"', False),
    ('"""It leaves out SetPasswordEncryptionOptions."""', False), ('x = "Signed off"', False)])
def test_the_check_catches_what_it_should(code: str, caught: bool, tmp_path: Path) -> None:
    path = tmp_path / "measure_x.py"
    path.write_text(code, encoding="utf-8")
    if caught:
        with pytest.raises(AssertionError):
            test_a_probe_names_no_crypto_member(path)
    else:
        test_a_probe_names_no_crypto_member(path)
