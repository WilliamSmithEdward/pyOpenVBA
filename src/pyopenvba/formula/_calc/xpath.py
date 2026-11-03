"""XPath 1.0 over a small XML document, for FILTERXML.

Excel answers FILTERXML from a full XPath 1.0 engine, so this is one: every axis, union, predicates and the core function
library, read against tests/fixtures/range_missing.json. A document with a DOCTYPE fails, as it does in Excel, which also
keeps entity declarations out of reach of what a macro passes in.
"""
from __future__ import annotations

import math
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from xml.parsers import expat


class XPathError(ValueError):
    """The document or the path is not one that can be answered: #VALUE! to the formula."""


@dataclass(eq=False)
class Node:
    """A node of the document, in the order it came in."""

    kind: str
    local: str = ""
    namespace: str = ""
    prefix: str = ""
    text: str = ""
    parent: Node | None = None
    children: list[Node] = field(default_factory=lambda: [])
    attributes: list[Node] = field(default_factory=lambda: [])
    order: int = 0

    @property
    def name(self) -> str:
        return f"{self.prefix}:{self.local}" if self.prefix else self.local

    def value(self) -> str:
        """The XPath string value: an element's is all the text inside it."""
        if self.kind in ("text", "attribute"):
            return self.text
        return "".join(child.value() for child in self.children)

    def own_text(self) -> str:
        """What FILTERXML answers for a node: an element's own text, and nothing for the text of those inside it."""
        if self.kind == "attribute":
            return self.text.strip()
        if self.kind != "element":
            return ""
        return "".join(child.text for child in self.children if child.kind == "text").strip()


Value = list[Node] | float | str | bool


def parse_document(text: str) -> Node:
    root = Node("root")
    stack = [root]
    counter = [0]

    def numbered(node: Node) -> Node:
        counter[0] += 1
        node.order = counter[0]
        return node

    def named(raw: str, kind: str) -> Node:
        parts = raw.split(" ")
        if len(parts) == 1:
            return numbered(Node(kind, local=parts[0], parent=stack[-1]))
        return numbered(Node(kind, namespace=parts[0], local=parts[1], prefix=parts[2] if len(parts) > 2 else "", parent=stack[-1]))

    def start(raw: str, attributes: dict[str, str]) -> None:
        element = named(raw, "element")
        stack[-1].children.append(element)
        stack.append(element)
        for key, value in attributes.items():
            attribute = named(key, "attribute")
            attribute.text = value
            element.attributes.append(attribute)

    def end(_: str) -> None:
        stack.pop()

    def characters(data: str) -> None:
        last = stack[-1].children[-1] if stack[-1].children else None
        if last is not None and last.kind == "text":
            last.text += data
        else:
            stack[-1].children.append(numbered(Node("text", text=data, parent=stack[-1])))

    def refuse(*_: object) -> None:
        raise XPathError("a document with a DOCTYPE")

    parser = expat.ParserCreate(namespace_separator=" ")
    parser.namespace_prefixes = True
    parser.ordered_attributes = False
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = characters
    parser.StartDoctypeDeclHandler = refuse
    try:
        parser.Parse(text, True)
    except expat.ExpatError as failure:
        raise XPathError(str(failure)) from failure
    _drop_blank_text(root)
    return root


def _drop_blank_text(node: Node) -> None:
    node.children = [child for child in node.children if child.kind != "text" or child.text.strip()]
    for child in node.children:
        _drop_blank_text(child)


# --- the path, as tokens --------------------------------------------------------------------------------------------

_TOKEN = re.compile(r"""\s*(?:(?P<number>\d+(?:\.\d*)?|\.\d+)|(?P<string>"[^"]*"|'[^']*')|(?P<name>[^\W\d][\w.\-]*(?::[^\W\d][\w.\-]*|:\*)?)
                    |(?P<op>::|//|\.\.|!=|<=|>=|[/()\[\]@,|+\-=<>*.$]))""", re.X)
_AXES = {"ancestor", "ancestor-or-self", "attribute", "child", "descendant", "descendant-or-self", "following", "following-sibling",
         "namespace", "parent", "preceding", "preceding-sibling", "self"}
_NODE_TYPES = {"comment", "text", "processing-instruction", "node"}


def _tokens(path: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    at, end = 0, len(path.rstrip())
    while at < end:
        match = _TOKEN.match(path, at)
        if match is None:
            raise XPathError(f"cannot read the path at {at}")
        kind = match.lastgroup or ""
        text = match.group(kind)
        # After something that ends a value, * multiplies and these names are operators.
        if found and found[-1][0] != "op" or (found and found[-1][1] in (")", "]", ".", "..")):
            if text == "*" or (kind == "name" and text in ("and", "or", "div", "mod")):
                kind = "op"
        found.append((kind, text))
        at = match.end()
    return found


@dataclass
class _Context:
    node: Node
    position: int
    size: int


_Expression = Callable[[_Context], Value]


class _Parser:
    def __init__(self, path: str) -> None:
        self.tokens = _tokens(path)
        self.at = 0
        if not self.tokens:
            raise XPathError("an empty path")

    def peek(self, ahead: int = 0) -> str:
        place = self.at + ahead
        return self.tokens[place][1] if place < len(self.tokens) else ""

    def peek_kind(self, ahead: int = 0) -> str:
        place = self.at + ahead
        return self.tokens[place][0] if place < len(self.tokens) else ""

    def take(self, wanted: str | None = None) -> str:
        text = self.peek()
        if self.at >= len(self.tokens) or (wanted is not None and text != wanted):
            raise XPathError(f"expected {wanted} in the path")
        self.at += 1
        return text

    def whole(self) -> _Expression:
        found = self.binary(0)
        if self.at != len(self.tokens):
            raise XPathError("more path than it could read")
        return found

    _LEVELS: tuple[tuple[str, ...], ...] = (("or",), ("and",), ("=", "!="), ("<", "<=", ">", ">="), ("+", "-"), ("*", "div", "mod"))

    def binary(self, level: int) -> _Expression:
        if level == len(self._LEVELS):
            return self.unary()
        left = self.binary(level + 1)
        while self.peek_kind() == "op" and self.peek() in self._LEVELS[level]:
            operator = self.take()
            left = _operation(operator, left, self.binary(level + 1))
        return left

    def unary(self) -> _Expression:
        if self.peek_kind() == "op" and self.peek() == "-":
            self.take()
            inner = self.unary()
            return lambda context: -_number(inner(context))
        left = self.path()
        while self.peek() == "|":
            self.take()
            right = self.path()
            left = _union(left, right)
        return left

    def path(self) -> _Expression:
        start = self.peek()
        kind = self.peek_kind()
        filtered = start in ("(", "$") or kind in ("number", "string") or (
            kind == "name" and self.peek(1) == "(" and start not in _NODE_TYPES)
        if not filtered:
            return self.location()
        primary = self.primary()
        while self.peek() == "[":
            primary = _filtered(primary, self.predicate())
        if self.peek() in ("/", "//"):
            return self.steps(primary)
        return primary

    def primary(self) -> _Expression:
        kind, text = self.peek_kind(), self.peek()
        if text == "(":
            self.take()
            inner = self.binary(0)
            self.take(")")
            return inner
        if kind == "number":
            self.take()
            number = float(text)
            return lambda _: number
        if kind == "string":
            self.take()
            value = text[1:-1]
            return lambda _: value
        if kind == "name" and self.peek(1) == "(":
            return self.call()
        raise XPathError(f"cannot read {text} in the path")

    def call(self) -> _Expression:
        name = self.take()
        self.take("(")
        arguments: list[_Expression] = []
        if self.peek() != ")":
            arguments.append(self.binary(0))
            while self.peek() == ",":
                self.take()
                arguments.append(self.binary(0))
        self.take(")")
        return _function(name, arguments)

    def location(self) -> _Expression:
        if self.peek() in ("/", "//"):
            def root(context: _Context) -> Value:
                node = context.node
                while node.parent is not None:
                    node = node.parent
                return [node]

            if self.peek() == "/" and not self.starts_step(1):
                self.take()
                return root
            return self.steps(root)
        first = self.step()
        return self.steps(_applied(lambda context: [context.node], first))

    def starts_step(self, ahead: int) -> bool:
        return self.peek_kind(ahead) == "name" or self.peek(ahead) in ("*", "@", ".", "..")

    def steps(self, left: _Expression) -> _Expression:
        while self.peek() in ("/", "//"):
            if self.take() == "//":
                left = _applied(left, _step("descendant-or-self", _any_node, []))
            left = _applied(left, self.step())
        return left

    def step(self) -> _Stepper:
        text = self.peek()
        if text == ".":
            self.take()
            return _step("self", _any_node, [])
        if text == "..":
            self.take()
            return _step("parent", _any_node, [])
        axis = "child"
        if text == "@":
            self.take()
            axis = "attribute"
        elif self.peek_kind() == "name" and self.peek(1) == "::":
            axis = self.take()
            self.take("::")
            if axis not in _AXES or axis == "namespace":
                raise XPathError(f"no axis {axis}")
        test = self.node_test(axis)
        predicates: list[_Expression] = []
        while self.peek() == "[":
            predicates.append(self.predicate())
        return _step(axis, test, predicates)

    def node_test(self, axis: str) -> Callable[[Node], bool]:
        kind, text = self.peek_kind(), self.peek()
        principal = "attribute" if axis == "attribute" else "element"
        if text == "*":
            self.take()
            return lambda node: node.kind == principal
        if kind != "name":
            raise XPathError(f"cannot read {text} as a step")
        self.take()
        if self.peek() == "(" and text in _NODE_TYPES:
            self.take("(")
            if text == "processing-instruction" and self.peek_kind() == "string":
                self.take()
            self.take(")")
            if text == "node":
                return _any_node
            return (lambda node: node.kind == "text") if text == "text" else (lambda _: False)
        if ":" in text:
            # No prefix is bound in the path, so a prefixed name cannot be resolved.
            raise XPathError(f"the prefix of {text} is not bound")
        return lambda node: node.kind == principal and node.local == text and not node.namespace

    def predicate(self) -> _Expression:
        self.take("[")
        inner = self.binary(0)
        self.take("]")
        return inner


def _any_node(_: Node) -> bool:
    return True


_Stepper = Callable[[Node], list[Node]]


def _following_in_document(node: Node) -> list[Node]:
    found: list[Node] = []
    current = node
    while current.parent is not None:
        siblings = current.parent.children
        if current.kind != "attribute":
            for sibling in siblings[siblings.index(current) + 1:]:
                found.append(sibling)
                found.extend(_descendants(sibling))
        current = current.parent
    return found


def _preceding_in_document(node: Node) -> list[Node]:
    found: list[Node] = []
    current = node
    while current.parent is not None:
        siblings = current.parent.children
        if current.kind != "attribute":
            for sibling in reversed(siblings[:siblings.index(current)]):
                found.extend(reversed(_descendants(sibling)))
                found.append(sibling)
        current = current.parent
    return found


def _descendants(node: Node) -> list[Node]:
    found: list[Node] = []
    for child in node.children:
        found.append(child)
        found.extend(_descendants(child))
    return found


def _ancestors(node: Node) -> list[Node]:
    found: list[Node] = []
    while node.parent is not None:
        node = node.parent
        found.append(node)
    return found


def _siblings(node: Node, after: bool) -> list[Node]:
    if node.parent is None or node.kind == "attribute":
        return []
    siblings = node.parent.children
    place = siblings.index(node)
    return siblings[place + 1:] if after else list(reversed(siblings[:place]))


_ALONG: dict[str, _Stepper] = {
    "child": lambda node: node.children,
    "descendant": _descendants,
    "descendant-or-self": lambda node: [node, *_descendants(node)],
    "self": lambda node: [node],
    "parent": lambda node: [] if node.parent is None else [node.parent],
    "ancestor": _ancestors,
    "ancestor-or-self": lambda node: [node, *_ancestors(node)],
    "attribute": lambda node: node.attributes,
    "following-sibling": lambda node: _siblings(node, True),
    "preceding-sibling": lambda node: _siblings(node, False),
    "following": _following_in_document,
    "preceding": _preceding_in_document,
}


def _step(axis: str, test: Callable[[Node], bool], predicates: list[_Expression]) -> _Stepper:
    along = _ALONG[axis]

    def run(node: Node) -> list[Node]:
        found = [one for one in along(node) if test(one)]
        for predicate in predicates:
            found = _kept(found, predicate)
        return found

    return run


def _kept(nodes: list[Node], predicate: _Expression) -> list[Node]:
    """What a predicate keeps, in the order it was given them: a number is a position."""
    size = len(nodes)
    kept: list[Node] = []
    for place, node in enumerate(nodes, 1):
        answer = predicate(_Context(node, place, size))
        wanted = answer == float(place) if isinstance(answer, (int, float)) and not isinstance(answer, bool) else _boolean(answer)
        if wanted:
            kept.append(node)
    return kept


def _in_order(nodes: list[Node]) -> list[Node]:
    return sorted({id(node): node for node in nodes}.values(), key=lambda node: node.order)


def _node_set(value: Value) -> list[Node]:
    if not isinstance(value, list):
        raise XPathError("a step needs nodes to start from")
    return value


def _applied(left: _Expression, step: _Stepper) -> _Expression:
    def run(context: _Context) -> Value:
        found: list[Node] = []
        for node in _node_set(left(context)):
            found.extend(step(node))
        return _in_order(found)

    return run


def _filtered(left: _Expression, predicate: _Expression) -> _Expression:
    return lambda context: _kept(_in_order(_node_set(left(context))), predicate)


def _union(left: _Expression, right: _Expression) -> _Expression:
    return lambda context: _in_order([*_node_set(left(context)), *_node_set(right(context))])


# --- values ---------------------------------------------------------------------------------------------------------


def _number_text(value: float) -> str:
    if math.isnan(value):
        return "NaN"
    if math.isinf(value):
        return "Infinity" if value > 0 else "-Infinity"
    return str(int(value)) if value == int(value) else repr(value)


def _string(value: Value) -> str:
    if isinstance(value, list):
        return _in_order(value)[0].value() if value else ""
    if isinstance(value, bool):
        return "true" if value else "false"
    return value if isinstance(value, str) else _number_text(value)


def _number(value: Value) -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, float):
        return value
    text = _string(value).strip()
    if re.fullmatch(r"-?(\d+(\.\d*)?|\.\d+)", text):
        return float(text)
    return math.nan


def _boolean(value: Value) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (list, str)):
        return len(value) > 0
    return not math.isnan(value) and value != 0.0


def _compared(operator: str, left: float | str | bool, right: float | str | bool) -> bool:
    if operator in ("=", "!="):
        if isinstance(left, bool) or isinstance(right, bool):
            same = _boolean(left) == _boolean(right)
        elif isinstance(left, float) or isinstance(right, float):
            same = _number(left) == _number(right)
        else:
            same = left == right
        return same if operator == "=" else not same
    one, other = _number(left), _number(right)
    return {"<": one < other, "<=": one <= other, ">": one > other, ">=": one >= other}[operator]


def _comparison(operator: str, left: Value, right: Value) -> bool:
    lefts: list[float | str | bool]
    rights: list[float | str | bool]
    if isinstance(left, list) and isinstance(right, bool):
        return _compared(operator, _boolean(left), right)
    if isinstance(right, list) and isinstance(left, bool):
        return _compared(operator, left, _boolean(right))
    lefts = [node.value() for node in left] if isinstance(left, list) else [left]
    rights = [node.value() for node in right] if isinstance(right, list) else [right]
    return any(_compared(operator, one, other) for one in lefts for other in rights)


def _operation(operator: str, left: _Expression, right: _Expression) -> _Expression:
    if operator == "or":
        return lambda context: _boolean(left(context)) or _boolean(right(context))
    if operator == "and":
        return lambda context: _boolean(left(context)) and _boolean(right(context))
    if operator in ("=", "!=", "<", "<=", ">", ">="):
        return lambda context: _comparison(operator, left(context), right(context))

    def run(context: _Context) -> Value:
        one, other = _number(left(context)), _number(right(context))
        if operator == "+":
            return one + other
        if operator == "-":
            return one - other
        if operator == "*":
            return one * other
        if operator == "mod":
            return math.nan if other == 0 or math.isinf(one) else math.fmod(one, other)
        if other == 0:
            return math.nan if one == 0 or math.isnan(one) else math.copysign(math.inf, one) * math.copysign(1.0, other)
        return one / other

    return run


def _round(value: float) -> float:
    return value if math.isnan(value) or math.isinf(value) else float(math.floor(value + 0.5))


def _substring(text: str, start: float, length: float | None) -> str:
    first = _round(start)
    last = math.inf if length is None else first + _round(length)
    if math.isnan(first) or math.isnan(last):
        return ""
    return "".join(char for place, char in enumerate(text, 1) if first <= place < last)


def _translate(text: str, source: str, target: str) -> str:
    mapping: dict[str, str] = {}
    for place, char in enumerate(source):
        mapping.setdefault(char, target[place] if place < len(target) else "")
    return "".join(mapping.get(char, char) for char in text)


def _first_node(arguments: list[Value], context: _Context) -> Node | None:
    nodes = _in_order(_node_set(arguments[0])) if arguments else [context.node]
    return nodes[0] if nodes else None


def _function(name: str, given: list[_Expression]) -> _Expression:
    limits: dict[str, tuple[int, int]] = {
        "last": (0, 0), "position": (0, 0), "count": (1, 1), "name": (0, 1), "local-name": (0, 1), "namespace-uri": (0, 1), "string": (0, 1),
        "concat": (2, 99), "starts-with": (2, 2), "contains": (2, 2), "substring-before": (2, 2), "substring-after": (2, 2), "substring": (2, 3),
        "string-length": (0, 1), "normalize-space": (0, 1), "translate": (3, 3), "boolean": (1, 1), "not": (1, 1), "true": (0, 0), "false": (0, 0),
        "number": (0, 1), "sum": (1, 1), "floor": (1, 1), "ceiling": (1, 1), "round": (1, 1),
    }
    limit = limits.get(name)
    if limit is None or not limit[0] <= len(given) <= limit[1]:
        raise XPathError(f"no function {name} of {len(given)} arguments")

    def run(context: _Context) -> Value:
        arguments = [one(context) for one in given]

        def text(place: int = 0) -> str:
            return _string(arguments[place]) if len(arguments) > place else context.node.value()

        if name == "last":
            return float(context.size)
        if name == "position":
            return float(context.position)
        if name == "count":
            return float(len(_node_set(arguments[0])))
        if name in ("name", "local-name", "namespace-uri"):
            node = _first_node(arguments, context)
            if node is None:
                return ""
            return {"name": node.name, "local-name": node.local, "namespace-uri": node.namespace}[name]
        if name == "string":
            return text()
        if name == "concat":
            return "".join(_string(one) for one in arguments)
        if name == "starts-with":
            return text(0).startswith(text(1))
        if name == "contains":
            return text(1) in text(0)
        if name == "substring-before":
            return text(0).partition(text(1))[0] if text(1) in text(0) else ""
        if name == "substring-after":
            return text(0).partition(text(1))[2] if text(1) in text(0) else ""
        if name == "substring":
            return _substring(text(0), _number(arguments[1]), _number(arguments[2]) if len(arguments) > 2 else None)
        if name == "string-length":
            return float(len(text()))
        if name == "normalize-space":
            return " ".join(text().split())
        if name == "translate":
            return _translate(text(0), text(1), text(2))
        if name == "boolean":
            return _boolean(arguments[0])
        if name == "not":
            return not _boolean(arguments[0])
        if name in ("true", "false"):
            return name == "true"
        if name == "number":
            return _number(arguments[0]) if arguments else _number(context.node.value())
        if name == "sum":
            return float(sum(_number(node.value()) for node in _node_set(arguments[0])))
        number = _number(arguments[0])
        if math.isnan(number) or math.isinf(number):
            return number
        if name == "floor":
            return float(math.floor(number))
        return float(math.ceil(number)) if name == "ceiling" else _round(number)

    return run


def select(document: str, path: str) -> list[Node]:
    """The nodes ``path`` selects in ``document``, in document order. A path that answers anything else fails, as
    FILTERXML has it: true(), a number, count()."""
    root = parse_document(document)
    try:
        answer = _Parser(path).whole()(_Context(root, 1, 1))
    except RecursionError as failure:
        raise XPathError("too deep") from failure
    if not isinstance(answer, list):
        raise XPathError("the path does not select nodes")
    return answer
