"""XPath 1.0 over a custom XML part's tree, as SelectNodes and SelectSingleNode answer it (MSXML's engine;
tests/fixtures/custom_xml*.json).

A name with no prefix is in no namespace, so a part with a default namespace is reached through a prefix its
NamespaceManager holds -- ns0 for that namespace from the start. A prefix it does not hold, an expression that is not
a node-set and one that stops short are error -2147467259 with MSXML's text, the expression quoted with the place
marked; other mistakes, variables and the namespace axis report themselves.
"""
from __future__ import annotations

import math
import re
from collections.abc import Callable, Sequence

from pyopenvba.apps.excel._custom_xml_tree import (ATTRIBUTE, COMMENT, ELEMENT, FAILED, INSTRUCTION, TEXT, Node,
                                                  document_order)
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

Value = list[Node] | float | str | bool
_TOKEN = re.compile(r"\s*(?:(?P<number>\d+(?:\.\d*)?|\.\d+)|(?P<literal>\"[^\"]*\"|'[^']*')|(?P<op>//|::|\.\.|!=|<=|>="
                    r"|[/\[\]()@,|.=<>+*-])|(?P<name>[A-Za-z_][\w.\-]*(?::(?:[A-Za-z_][\w.\-]*|\*))?)|(?P<bad>\S))")
_AXES = {'child', 'descendant', 'descendant-or-self', 'parent', 'ancestor', 'ancestor-or-self', 'following-sibling',
         'preceding-sibling', 'following', 'preceding', 'attribute', 'self'}
_REVERSE = {'parent', 'ancestor', 'ancestor-or-self', 'preceding-sibling', 'preceding'}
_KIND_TESTS = {'node', 'text', 'comment', 'processing-instruction'}
_OPERATOR_NAMES = {'and', 'or', 'div', 'mod'}


def _marked(text: str, start: int, end: int) -> str:
    return text[:start] + ('-->' + text[start:end] if end > start else '') + '<--'


class _Tokens:
    def __init__(self, text: str) -> None:
        self.text = text
        self.items: list[tuple[str, str, int, int]] = []
        place = 0
        while place < len(text):
            found = _TOKEN.match(text, place)
            if found is None or found.lastgroup is None:
                if text[place:].strip():
                    raise VBAUnsupportedError(f'the XPath {text!r} is not implemented')
                break
            group = kind = found.lastgroup
            if kind == 'bad':
                raise VBAUnsupportedError(f'the XPath {text!r} is not implemented')
            value = found.group(kind)
            start = found.start(kind)
            if kind == 'name' or (kind == 'op' and value == '*'):
                # After something that ends an operand, * multiplies and and/or/div/mod are operators.
                previous = self.items[-1] if self.items else None
                operand_ended = previous is not None and not (
                    previous[0] == 'op' and previous[1] not in (')', ']', '.', '..')) and previous[0] != 'operator'
                if operand_ended and (value == '*' or value in _OPERATOR_NAMES):
                    kind = 'operator'
                elif value == '*':
                    kind = 'name'
            self.items.append((kind, value, start, found.end(group)))
            place = found.end()
        self.at = 0

    def peek(self, offset: int = 0) -> tuple[str, str, int, int] | None:
        place = self.at + offset
        return self.items[place] if place < len(self.items) else None

    def take(self) -> tuple[str, str, int, int]:
        token = self.peek()
        if token is None:
            raise error(FAILED, 'Expression expected.\n\n' + _marked(self.text, len(self.text), len(self.text)))
        self.at += 1
        return token

    def accept(self, value: str) -> bool:
        token = self.peek()
        if token is not None and token[1] == value and token[0] in ('op', 'operator'):
            self.at += 1
            return True
        return False

    def expect(self, value: str) -> None:
        if not self.accept(value):
            token = self.peek()
            if token is None:
                raise error(FAILED, 'Expression expected.\n\n' + _marked(self.text, len(self.text), len(self.text)))
            raise VBAUnsupportedError(f'the XPath {self.text!r} is not implemented')


Expression = Callable[['_Context'], Value]


class _Context:
    __slots__ = ('node', 'position', 'size')

    def __init__(self, node: Node, position: int = 1, size: int = 1) -> None:
        self.node, self.position, self.size = node, position, size


class _Compiler:
    def __init__(self, text: str, resolve: Callable[[str], str], order: dict[int, int]) -> None:
        self.tokens = _Tokens(text)
        self.resolve = resolve
        self.order = order

    # --- expressions ---------------------------------------------------------------------------------------------

    def expression(self) -> Expression:
        return self._binary(self._and, ('or',), lambda a, b: _boolean(a) or _boolean(b))

    def _and(self) -> Expression:
        return self._binary(self._equality, ('and',), lambda a, b: _boolean(a) and _boolean(b))

    def _equality(self) -> Expression:
        return self._binary(self._relational, ('=', '!='), None)

    def _relational(self) -> Expression:
        return self._binary(self._additive, ('<', '>', '<=', '>='), None)

    def _additive(self) -> Expression:
        return self._binary(self._multiplicative, ('+', '-'), None)

    def _multiplicative(self) -> Expression:
        return self._binary(self._unary, ('*', 'div', 'mod'), None)

    def _binary(self, operand: Callable[[], Expression], operators: tuple[str, ...],
                logic: Callable[[Value, Value], bool] | None) -> Expression:
        left = operand()
        while True:
            token = self.tokens.peek()
            if token is None or token[1] not in operators or token[0] not in ('op', 'operator'):
                return left
            self.tokens.take()
            right = operand()
            left = self._combined(left, right, token[1], logic)

    def _combined(self, left: Expression, right: Expression, operator: str,
                  logic: Callable[[Value, Value], bool] | None) -> Expression:
        if logic is not None:
            if operator == 'or':
                return lambda context: _boolean(left(context)) or _boolean(right(context))
            return lambda context: _boolean(left(context)) and _boolean(right(context))
        if operator in ('=', '!=', '<', '>', '<=', '>='):
            return lambda context: _compare(left(context), right(context), operator)
        arithmetic: dict[str, Callable[[float, float], float]] = {
            '+': lambda a, b: a + b, '-': lambda a, b: a - b, '*': lambda a, b: a * b,
            'div': lambda a, b: a / b if b else (math.nan if a == 0 or math.isnan(a) else math.copysign(math.inf, a)),
            'mod': lambda a, b: math.fmod(a, b) if b else math.nan}
        function = arithmetic[operator]
        return lambda context: function(_number(left(context)), _number(right(context)))

    def _unary(self) -> Expression:
        if self.tokens.accept('-'):
            inner = self._unary()
            return lambda context: -_number(inner(context))
        return self._union()

    def _union(self) -> Expression:
        first = self._path()
        parts = [first]
        while self.tokens.accept('|'):
            parts.append(self._path())
        if len(parts) == 1:
            return first

        def union(context: _Context) -> Value:
            found: dict[int, Node] = {}
            for part in parts:
                value = part(context)
                if not isinstance(value, list):
                    raise error(FAILED, 'Expression must evaluate to a node-set.\n\n'
                                + _marked(self.tokens.text, 0, len(self.tokens.text)))
                found.update((id(node), node) for node in value)
            return self._sorted(list(found.values()))
        return union

    # --- paths -----------------------------------------------------------------------------------------------------

    def _path(self) -> Expression:
        token = self.tokens.peek()
        if token is None:
            raise error(FAILED, 'Expression expected.\n\n' + _marked(self.tokens.text, len(self.tokens.text),
                                                                      len(self.tokens.text)))
        if token[1] in ('/', '//') and token[0] == 'op':
            return self._absolute()
        if self._starts_step(token):
            steps = self._relative()
            return lambda context: self._walk([context.node], steps)
        primary = self._filter()
        token = self.tokens.peek()
        if token is not None and token[0] == 'op' and token[1] in ('/', '//'):
            self.tokens.take()
            steps = ([self._descendant_or_self()] if token[1] == '//' else []) + self._relative()

            def filtered_path(context: _Context) -> Value:
                start = primary(context)
                if not isinstance(start, list):
                    raise error(FAILED, 'Expression must evaluate to a node-set.\n\n'
                                + _marked(self.tokens.text, 0, len(self.tokens.text)))
                return self._walk(start, steps)
            return filtered_path
        return primary

    def _starts_step(self, token: tuple[str, str, int, int]) -> bool:
        kind, value = token[0], token[1]
        if kind == 'op':
            return value in ('@', '.', '..')
        if kind != 'name':
            return False
        following = self.tokens.peek(1)
        if following is not None and following[1] == '::':
            return True
        if following is not None and following[1] == '(' and following[0] == 'op':
            return value in _KIND_TESTS
        return True

    def _absolute(self) -> Expression:
        leading = self.tokens.take()[1]
        steps: list[Callable[[list[Node]], list[Node]]] = []
        if leading == '//':
            steps = [self._descendant_or_self()] + self._relative()
        else:
            token = self.tokens.peek()
            if token is not None and self._starts_step(token):
                steps = self._relative()
        return lambda context: self._walk([_document(context.node)], steps)

    def _relative(self) -> list[Callable[[list[Node]], list[Node]]]:
        steps = [self._step()]
        while True:
            token = self.tokens.peek()
            if token is None or token[0] != 'op' or token[1] not in ('/', '//'):
                return steps
            self.tokens.take()
            if token[1] == '//':
                steps.append(self._descendant_or_self())
            steps.append(self._step())

    def _descendant_or_self(self) -> Callable[[list[Node]], list[Node]]:
        return lambda nodes: self._axis_nodes(nodes, 'descendant-or-self', lambda node: True, [])

    def _step(self) -> Callable[[list[Node]], list[Node]]:
        if self.tokens.accept('.'):
            return lambda nodes: nodes
        if self.tokens.accept('..'):
            return lambda nodes: self._axis_nodes(nodes, 'parent', lambda node: True, [])
        axis = 'child'
        if self.tokens.accept('@'):
            axis = 'attribute'
        else:
            token, following = self.tokens.peek(), self.tokens.peek(1)
            if token is not None and token[0] == 'name' and following is not None and following[1] == '::':
                axis = token[1]
                if axis not in _AXES:
                    raise VBAUnsupportedError(f'the XPath axis {axis!r} is not implemented')
                self.tokens.take()
                self.tokens.take()
        test = self._node_test(axis)
        predicates = self._predicates()
        return lambda nodes: self._axis_nodes(nodes, axis, test, predicates)

    def _node_test(self, axis: str) -> Callable[[Node], bool]:
        token = self.tokens.take()
        if token[0] != 'name':
            raise VBAUnsupportedError(f'the XPath {self.tokens.text!r} is not implemented')
        name = token[1]
        principal = ATTRIBUTE if axis == 'attribute' else ELEMENT
        if name in _KIND_TESTS and self.tokens.peek() is not None and self.tokens.peek()[1] == '(':  # type: ignore[index]
            self.tokens.expect('(')
            target = None
            if name == 'processing-instruction' and self.tokens.peek() is not None and \
                    self.tokens.peek()[0] == 'literal':  # type: ignore[index]
                target = self.tokens.take()[1][1:-1]
            self.tokens.expect(')')
            kinds = {'node': None, 'text': TEXT, 'comment': COMMENT, 'processing-instruction': INSTRUCTION}[name]
            return lambda node: (kinds is None or node.kind == kinds) and (target is None or node.name == target)
        if name == '*':
            return lambda node: node.kind == principal
        prefix, _, local = name.rpartition(':')
        uri = self.resolve(prefix) if prefix else ''
        if local == '*':
            return lambda node: node.kind == principal and node.uri == uri
        return lambda node: node.kind == principal and node.local == local and node.uri == uri

    def _predicates(self) -> list[Expression]:
        found: list[Expression] = []
        while self.tokens.accept('['):
            found.append(self.expression())
            self.tokens.expect(']')
        return found

    def _axis_nodes(self, nodes: list[Node], axis: str, test: Callable[[Node], bool],
                    predicates: list[Expression]) -> list[Node]:
        found: dict[int, Node] = {}
        for node in nodes:
            chosen = [one for one in _axis(node, axis) if test(one)]
            for predicate in predicates:
                chosen = _filtered(chosen, predicate, reverse=axis in _REVERSE)
            found.update((id(one), one) for one in chosen)
        return self._sorted(list(found.values()))

    def _walk(self, nodes: list[Node], steps: Sequence[Callable[[list[Node]], list[Node]]]) -> list[Node]:
        for step in steps:
            nodes = step(nodes)
        return nodes

    def _sorted(self, nodes: list[Node]) -> list[Node]:
        return sorted(nodes, key=lambda node: self.order.get(id(node), 0))

    # --- primaries ------------------------------------------------------------------------------------------------

    def _filter(self) -> Expression:
        primary = self._primary()
        predicates = self._predicates()
        if not predicates:
            return primary

        def filtered(context: _Context) -> Value:
            value = primary(context)
            if not isinstance(value, list):
                raise VBAUnsupportedError(f'the XPath {self.tokens.text!r} is not implemented')
            for predicate in predicates:
                value = _filtered(value, predicate, reverse=False)
            return value
        return filtered

    def _primary(self) -> Expression:
        token = self.tokens.take()
        kind, value = token[0], token[1]
        if kind == 'number':
            number = float(value)
            return lambda context: number
        if kind == 'literal':
            text = value[1:-1]
            return lambda context: text
        if kind == 'op' and value == '(':
            inner = self.expression()
            self.tokens.expect(')')
            return inner
        if kind == 'name' and self.tokens.accept('('):
            arguments: list[Expression] = []
            if not self.tokens.accept(')'):
                arguments.append(self.expression())
                while self.tokens.accept(','):
                    arguments.append(self.expression())
                self.tokens.expect(')')
            return _function(value, arguments)
        raise VBAUnsupportedError(f'the XPath {self.tokens.text!r} is not implemented')


def _document(node: Node) -> Node:
    while node.parent is not None:
        node = node.parent
    return node


def _descendants(node: Node) -> list[Node]:
    found: list[Node] = []
    for child in node.children:
        found.append(child)
        found += _descendants(child)
    return found


def _axis(node: Node, axis: str) -> list[Node]:
    parent = node.parent
    if axis == 'child':
        return list(node.children)
    if axis == 'attribute':
        return node.plain_attributes() if node.kind == ELEMENT else []
    if axis == 'self':
        return [node]
    if axis == 'parent':
        return [parent] if parent is not None else []
    if axis in ('ancestor', 'ancestor-or-self'):
        found = [node] if axis == 'ancestor-or-self' else []
        while parent is not None:
            found.append(parent)
            parent = parent.parent
        return found
    if axis == 'descendant':
        return _descendants(node)
    if axis == 'descendant-or-self':
        return [node, *_descendants(node)]
    if node.kind == ATTRIBUTE or parent is None:
        if axis in ('following', 'preceding') and node.kind == ATTRIBUTE and parent is not None:
            return _axis(parent, axis) if axis == 'preceding' else [*_descendants(parent), *_axis(parent, 'following')]
        return []
    siblings = parent.children
    place = siblings.index(node)
    if axis == 'following-sibling':
        return siblings[place + 1:]
    if axis == 'preceding-sibling':
        return list(reversed(siblings[:place]))
    found: list[Node] = []
    if axis == 'following':
        for sibling in siblings[place + 1:]:
            found += [sibling, *_descendants(sibling)]
        return found + _axis(parent, 'following')
    for sibling in reversed(siblings[:place]):
        found += list(reversed([sibling, *_descendants(sibling)]))
    return found + [one for one in _axis(parent, 'preceding')]


def _filtered(nodes: list[Node], predicate: Expression, *, reverse: bool) -> list[Node]:
    size = len(nodes)
    kept: list[Node] = []
    for place, node in enumerate(nodes, 1):
        value = predicate(_Context(node, place, size))
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if float(value) == float(place):
                kept.append(node)
        elif _boolean(value):
            kept.append(node)
    return list(reversed(kept)) if reverse else kept


def _text(value: Value) -> str:
    if isinstance(value, list):
        return value[0].text() if value else ''
    if isinstance(value, bool):
        return 'true' if value else 'false'
    if isinstance(value, (int, float)):
        number = float(value)
        if math.isnan(number):
            return 'NaN'
        if math.isinf(number):
            return 'Infinity' if number > 0 else '-Infinity'
        return str(int(number)) if number.is_integer() else repr(number)
    return value


def _number(value: Value) -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, float):
        return value
    text = _text(value).strip()
    try:
        return float(text) if re.fullmatch(r'-?(\d+(\.\d*)?|\.\d+)', text) else math.nan
    except ValueError:
        return math.nan


def _boolean(value: Value) -> bool:
    if isinstance(value, list):
        return bool(value)
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) != 0.0 and not math.isnan(float(value))
    return bool(value)


def _compare(left: Value, right: Value, operator: str) -> bool:
    if isinstance(left, list) or isinstance(right, list):
        lefts: list[Value] = [one.text() for one in left] if isinstance(left, list) else [left]
        rights: list[Value] = [one.text() for one in right] if isinstance(right, list) else [right]
        if isinstance(left, list) and not isinstance(right, list):
            lefts = [_as(one, right) for one in lefts]
        if isinstance(right, list) and not isinstance(left, list):
            rights = [_as(one, left) for one in rights]
        return any(_compare(a, b, operator) for a in lefts for b in rights)
    if operator in ('=', '!='):
        if isinstance(left, bool) or isinstance(right, bool):
            same = _boolean(left) == _boolean(right)
        elif isinstance(left, float) or isinstance(right, float):
            same = _number(left) == _number(right)
        else:
            same = left == right
        return same if operator == '=' else not same
    a, b = _number(left), _number(right)
    return {'<': a < b, '>': a > b, '<=': a <= b, '>=': a >= b}[operator]


def _as(text: Value, like: Value) -> Value:
    """A node's string compared with a number or a boolean as one."""
    if isinstance(like, bool):
        return _boolean(text)
    if isinstance(like, float):
        return _number(text)
    return text


def _function(name: str, arguments: list[Expression]) -> Expression:
    def one(context: _Context, place: int = 0) -> Value:
        return arguments[place](context) if len(arguments) > place else [context.node]

    def string(context: _Context, place: int = 0) -> str:
        return _text(one(context, place))

    table: dict[str, Callable[[_Context], Value]] = {
        'last': lambda context: float(context.size),
        'position': lambda context: float(context.position),
        'count': lambda context: float(len(_node_set(one(context)))),
        'local-name': lambda context: _named(one(context), 'local'),
        'name': lambda context: _named(one(context), 'name'),
        'namespace-uri': lambda context: _named(one(context), 'uri'),
        'string': lambda context: string(context),
        'concat': lambda context: ''.join(_text(argument(context)) for argument in arguments),
        'starts-with': lambda context: string(context).startswith(string(context, 1)),
        'contains': lambda context: string(context, 1) in string(context),
        'substring-before': lambda context: string(context).partition(string(context, 1))[0]
        if string(context, 1) in string(context) else '',
        'substring-after': lambda context: string(context).partition(string(context, 1))[2]
        if string(context, 1) in string(context) else '',
        'string-length': lambda context: float(len(string(context))),
        'normalize-space': lambda context: ' '.join(string(context).split()),
        'not': lambda context: not _boolean(one(context)),
        'true': lambda context: True,
        'false': lambda context: False,
        'boolean': lambda context: _boolean(one(context)),
        'number': lambda context: _number(one(context)),
        'sum': lambda context: float(sum(_number(node.text()) for node in _node_set(one(context)))),
        'floor': lambda context: float(math.floor(_number(one(context)))),
        'ceiling': lambda context: float(math.ceil(_number(one(context)))),
    }
    if name not in table:
        raise VBAUnsupportedError(f'the XPath function {name}() is not implemented')
    return table[name]


def _node_set(value: Value) -> list[Node]:
    if not isinstance(value, list):
        raise VBAUnsupportedError('an XPath function given what is not a node-set is not implemented')
    return value


def _named(value: Value, which: str) -> str:
    """local-name(), name() and namespace-uri() of the first node, '' for none or one with no name."""
    node = _first(value)
    if node is None or node.kind not in (ELEMENT, ATTRIBUTE):
        return ''
    return node.local if which == 'local' else node.name if which == 'name' else node.uri


def _first(value: Value) -> Node | None:
    return value[0] if isinstance(value, list) and value else None


def selected(text: str, context: Node, resolve: Callable[[str], str]) -> list[Node]:
    """The nodes ``text`` selects from ``context``, in document order."""
    order = {id(node): place for place, node in enumerate(document_order(_document(context)))}
    compiler = _Compiler(text, resolve, order)
    if not compiler.tokens.items:
        raise error(FAILED, 'Expression must evaluate to a node-set.\n\n' + _marked(text, 0, 0))
    expression = compiler.expression()
    leftover = compiler.tokens.peek()
    if leftover is not None:
        raise VBAUnsupportedError(f'the XPath {text!r} is not implemented')
    value = expression(_Context(context))
    if not isinstance(value, list):
        raise error(FAILED, 'Expression must evaluate to a node-set.\n\n' + _marked(text, 0, len(text)))
    return value


def undeclared(prefix: str) -> Exception:
    return error(FAILED, f"Reference to undeclared namespace prefix: '{prefix}'.\n")
