"""No live probe touches a member that could send mail from the account Excel runs under, whose results are not worth that
risk: such a member stays unmeasured. A script's docstring may name them, to say so; its code may not."""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

SCRIPTS = sorted((Path(__file__).parents[1] / "scripts").glob("measure_*.py"))
#: The members, each reached with a dot or named alone in a list a probe walks.
BANNED = ("SendMail", "SendForReview", "ReplyWithChanges", "ReplyAll", "EndReview", "RouteWorkbook", "RoutingSlip", "HasRoutingSlip",
          "HasMailer", "Mailer", "MailEnvelope", "SendFaxOverInternet", "MailLogon", "MailLogoff", "MailSession", "MailSystem",
          "xlDialogSendMail")
#: Words too common to look for alone: only with a dot, as a member. A pivot field may be called Route.
BANNED_AFTER_A_DOT = ("Route", "Reply")
_NAMED = re.compile(r"(?<![\w])(?:" + "|".join(BANNED) + r")(?![\w])")
_MEMBER = re.compile(r"\.(?:" + "|".join(BANNED_AFTER_A_DOT) + r")(?![\w])")
_FOLLOWED = re.compile(r"\.Follow(?:Hyperlink)?(?![\w])")


def _code_strings(path: Path) -> list[str]:
    """Every string in a script's code, which is where a probe's VBA is; not its docstring, or a comment."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    docstring = tree.body[0].value if tree.body and isinstance(tree.body[0], ast.Expr) else None
    return [node.value for node in ast.walk(tree) if isinstance(node, ast.Constant) and isinstance(node.value, str) and node is not docstring]


def test_there_are_probes_to_check() -> None:
    assert len(SCRIPTS) > 100


@pytest.mark.parametrize("path", SCRIPTS, ids=lambda path: path.stem)
def test_a_probe_names_no_mail_member(path: Path) -> None:
    strings = _code_strings(path)
    found = sorted({hit for text in strings for hit in (*_NAMED.findall(text), *_MEMBER.findall(text))})
    assert found == []
    text = "\n".join(strings)
    assert not (_FOLLOWED.search(text) and "mailto:" in text), "follows a link in a script that holds a mailto: address"


@pytest.mark.parametrize(("code", "caught"), [
    ('x = "wb.SendMail"', True), ('x = ["MailSession", "Build"]', True), ('x = f"r = {A}.MailSystem"', True), ('x = "wb.Reply"', True),
    ('x = "ws.Hyperlinks(1).Follow"\ny = "mailto:a@example.com"', True), ('x = "a Route"', False), ('x = "ws.Hyperlinks(1).Follow"', False),
    ('"""It leaves out SendMail."""', False), ('x = "ws.Hyperlinks.Add c, \'mailto:a@example.com\'"', False)])
def test_the_check_catches_what_it_should(code: str, caught: bool, tmp_path: Path) -> None:
    path = tmp_path / "measure_x.py"
    path.write_text(code, encoding="utf-8")
    if caught:
        with pytest.raises(AssertionError):
            test_a_probe_names_no_mail_member(path)
    else:
        test_a_probe_names_no_mail_member(path)
