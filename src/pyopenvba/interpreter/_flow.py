"""Procedure-wide control flow, including labels inside structured statements."""
from __future__ import annotations

from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field, fields
from typing import cast

from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._lex import default_type, strip_suffix


@dataclass(slots=True)
class Instruction:
    kind: str
    statement: A.Stmt
    scopes: tuple[int, ...] = ()
    target: int = -1
    otherwise: int = -1
    owner: A.Stmt | None = None
    clauses: tuple[A.CaseClause, ...] = ()
    resume_next: int = -1


@dataclass(slots=True)
class Program:
    instructions: list[Instruction] = field(default_factory=lambda: [])
    labels: dict[str, int] = field(default_factory=lambda: {})
    arrays: list[A.VarDecl] = field(default_factory=lambda: [])


@dataclass(slots=True)
class Loop:
    kind: str
    statement: A.Stmt
    exits: list[Instruction] = field(default_factory=lambda: [])


def label_key(label: str) -> str:
    """Line-number labels are numeric; named labels are case insensitive."""
    return str(int(label)) if label.isdecimal() else label.lower()


def nested_bodies(statement: A.Stmt) -> Iterator[list[A.Stmt]]:
    if isinstance(statement, A.If):
        yield from (branch.body for branch in statement.branches)
        yield statement.otherwise
    elif isinstance(statement, (A.For, A.ForEach, A.DoLoop, A.WhileLoop, A.With)):
        yield statement.body
    elif isinstance(statement, A.SelectCase):
        yield from (case.body for case in statement.blocks)
        if statement.otherwise is not None:
            yield statement.otherwise


def compile_program(procedure: A.Procedure, where: str, global_declarations: Sequence[A.Dim] = (),
                    value_type: Callable[[str], bool] | None = None, defaults: Mapping[str, str] | None = None) -> Program:
    program = Program()
    loops: list[Loop] = []
    declarations = {item.name.lower(): (item.declared, item.is_array)
                    for group in global_declarations for item in group.decls}
    declarations.update({item.name.lower(): (item.declared, item.is_array or item.param_array) for item in procedure.params})

    def collect(body: Sequence[A.Stmt]) -> None:
        for statement in body:
            if isinstance(statement, A.Dim):
                declarations.update({item.name.lower(): (item.declared, item.is_array) for item in statement.decls})
            for child in nested_bodies(statement):
                collect(child)

    collect(procedure.body)
    if procedure.kind in ('function', 'get'):
        declarations[procedure.name.lower()] = (procedure.returns, procedure.returns_array)

    def declaration(name: str) -> tuple[str, bool]:
        return declarations.get(strip_suffix(name).lower(), (default_type(name, defaults or {}), False))

    used: set[str] = set()

    def read_names(value: object) -> None:
        if isinstance(value, A.Name):
            used.add(strip_suffix(value.name).lower())
        elif isinstance(value, (list, tuple)):
            for child in cast(Sequence[object], value):
                read_names(child)
        elif isinstance(value, (A.Expr, A.Argument, A.CaseClause, A.VarDecl)):
            for field_ in fields(value):
                if isinstance(value, A.Index) and field_.name == 'target' and isinstance(value.target, A.Name):
                    continue  # A call's bare name is not an implicit scalar.
                read_names(getattr(value, field_.name))

    def validate_arrays(body: Sequence[A.Stmt]) -> None:
        for statement in body:
            if isinstance(statement, A.ReDim):
                for index, item in enumerate(statement.decls):
                    if index < len(statement.targets) and statement.targets[index] is not None:
                        continue
                    key = item.name.lower()
                    known = declarations.get(key)
                    if known is None:
                        if key in used:
                            raise VBACompileError('Expected array', where=where)
                        declarations[key] = (item.declared, True)
                        program.arrays.append(A.VarDecl(name=item.name, declared=item.declared, is_array=True))
                    elif not known[1] and known[0] not in ('', 'Variant'):
                        raise VBACompileError('Expected array', where=where)
                    elif known[1] and item.explicit_type and known[0] != item.declared:
                        raise VBACompileError("Can't change data types of array elements", where=where)
            if isinstance(statement, A.If):
                for branch in statement.branches:
                    read_names(branch.condition)
                    validate_arrays(branch.body)
                validate_arrays(statement.otherwise)
            elif isinstance(statement, A.SelectCase):
                read_names(statement.subject)
                for case in statement.blocks:
                    read_names(case.clauses)
                    validate_arrays(case.body)
                if statement.otherwise is not None:
                    validate_arrays(statement.otherwise)
            else:
                for field_ in fields(statement):
                    if isinstance(statement, A.CallStmt) and field_.name == 'callee':
                        continue
                    if isinstance(statement, A.DoLoop) and statement.at_end and field_.name == 'condition':
                        continue
                    read_names(getattr(statement, field_.name))
                for child in nested_bodies(statement):
                    validate_arrays(child)
                if isinstance(statement, A.DoLoop) and statement.at_end:
                    read_names(statement.condition)

    validate_arrays(procedure.body)

    def emit(kind: str, statement: A.Stmt, scopes: tuple[int, ...], **kwargs: object) -> Instruction:
        instruction = Instruction(kind, statement, scopes)
        for name, value in kwargs.items():
            setattr(instruction, name, value)
        program.instructions.append(instruction)
        return instruction

    def block(body: Sequence[A.Stmt], scopes: tuple[int, ...]) -> None:
        for statement in body:
            if isinstance(statement, A.Label):
                key = label_key(statement.name)
                if key in program.labels:
                    raise VBACompileError(f"duplicate label {statement.name}", where=where)
                program.labels[key] = len(program.instructions)
                emit('exec', statement, scopes)
            elif isinstance(statement, A.If):
                endings: list[Instruction] = []
                for branch in statement.branches:
                    test = emit('if', statement, scopes)
                    # The branch expression is held in a synthetic If with
                    # just this branch, retaining the original source line.
                    test.statement = A.If(line=statement.line, branches=[branch])
                    test.target = len(program.instructions)
                    block(branch.body, scopes)
                    endings.append(emit('jump', statement, scopes))
                    test.otherwise = len(program.instructions)
                block(statement.otherwise, scopes)
                for ending in endings:
                    ending.target = len(program.instructions)
            elif isinstance(statement, A.SelectCase):
                initial = emit('select', statement, scopes)
                endings = []
                for case in statement.blocks:
                    test = emit('case', statement, scopes, clauses=tuple(case.clauses))
                    test.target = len(program.instructions)
                    block(case.body, scopes)
                    endings.append(emit('jump', statement, scopes))
                    test.otherwise = len(program.instructions)
                if statement.otherwise is not None:
                    block(statement.otherwise, scopes)
                for ending in endings:
                    ending.target = len(program.instructions)
                initial.resume_next = len(program.instructions)
                emit('select_end', statement, scopes)
            elif isinstance(statement, (A.For, A.ForEach)):
                if isinstance(statement, A.ForEach) and isinstance(statement.target, A.Name):
                    declared, array = declaration(statement.target.name)
                    if array or (value_type is not None and value_type(declared)):
                        raise VBACompileError('For Each control variable must be Variant or Object', where=where)
                    if isinstance(statement.collection, A.Name) and declaration(statement.collection.name)[1] \
                            and declared not in ('', 'Variant'):
                        raise VBACompileError('For Each control variable on arrays must be Variant', where=where)
                kind = 'for' if isinstance(statement, A.For) else 'each'
                initial = emit(kind + '_init', statement, scopes)
                initial.target = len(program.instructions)
                loop = Loop('for', statement)
                loops.append(loop)
                block(statement.body, scopes)
                following = emit(kind + '_next', statement, scopes, target=initial.target)
                loops.pop()
                end = len(program.instructions)
                initial.otherwise = following.otherwise = end
                if isinstance(statement, A.ForEach):
                    emit('each_end', statement, scopes)
                for ending in loop.exits:
                    ending.target = end
            elif isinstance(statement, (A.DoLoop, A.WhileLoop)):
                head = len(program.instructions)
                initial_test: Instruction | None = None
                if isinstance(statement, A.WhileLoop) or (statement.test and not statement.at_end):
                    initial_test = emit('loop_test', statement, scopes)
                    initial_test.target = len(program.instructions)
                body_start = len(program.instructions)
                loop = Loop('while' if isinstance(statement, A.WhileLoop) else 'do', statement)
                loops.append(loop)
                block(statement.body, scopes)
                loops.pop()
                if isinstance(statement, A.DoLoop) and statement.test and statement.at_end:
                    final_test = emit('loop_test', statement, scopes, target=body_start)
                    final_test.otherwise = len(program.instructions)
                else:
                    emit('jump', statement, scopes, target=head)
                if initial_test is not None:
                    initial_test.otherwise = len(program.instructions)
                for ending in loop.exits:
                    ending.target = len(program.instructions)
            elif isinstance(statement, A.With):
                emit('with_init', statement, scopes)
                block(statement.body, (*scopes, id(statement)))
                emit('with_end', statement, scopes)
            elif isinstance(statement, A.ExitStmt) and statement.what in ('for', 'do'):
                loop = next((loop for loop in reversed(loops) if loop.kind == statement.what), None)
                if loop is None:
                    raise VBACompileError(f"Exit {statement.what} is outside its loop", where=where)
                loop.exits.append(emit('exit_loop', statement, scopes, owner=loop.statement))
            elif isinstance(statement, A.ExitStmt):
                expected = 'property' if procedure.is_property else procedure.kind
                if statement.what != expected:
                    raise VBACompileError(f'Exit {statement.what} is outside its procedure', where=where)
                emit('exec', statement, scopes)
            elif isinstance(statement, A.GoTo):
                emit('goto', statement, scopes)
            elif isinstance(statement, A.OnGoto):
                emit('computed', statement, scopes)
            elif isinstance(statement, A.ReturnStmt):
                emit('return', statement, scopes)
            else:
                emit('exec', statement, scopes)

    block(procedure.body, ())
    for instruction in program.instructions:
        statement = instruction.statement
        labels = ([statement.label] if isinstance(statement, A.GoTo)
                  else statement.labels if isinstance(statement, A.OnGoto)
                  else [statement.label] if isinstance(statement, A.OnError) and statement.mode == 'goto'
                  else [statement.label] if isinstance(statement, A.Resume) and statement.mode == 'label'
                  else [])
        for label in labels:
            if label_key(label) not in program.labels:
                raise VBACompileError(f"label {label} is not in {procedure.name}", where=where)
    return program
