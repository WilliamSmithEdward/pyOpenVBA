"""Demand checks for type characters on statically resolved VBA references."""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import fields
from typing import TYPE_CHECKING, cast

from pyopenvba.exceptions import VBACompileError, VBARuntimeError
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._compile import Info, nodes, variables
from pyopenvba.interpreter._lex import SUFFIX_TYPES, default_type, strip_suffix
from pyopenvba.interpreter._values import type_name

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter, ModuleRuntime


def validate_type_characters(interpreter: Interpreter, module: ModuleRuntime, procedure: A.Procedure,
                             implicit_arrays: Sequence[A.VarDecl]) -> None:
    scope = module.program_scope
    cacheable = id(procedure) in module.programs
    if cacheable and module.reference_checks.get(id(procedure)) == scope:
        return
    body = list(nodes(procedure.body))
    if not any(isinstance(node, (A.Name, A.Member)) and node.name[-1:] in SUFFIX_TYPES for node in body):
        if cacheable:
            module.reference_checks[id(procedure)] = scope
        return
    local = {name: info for owner in interpreter.modules.values() if not owner.is_class and owner is not module
             for name, info in variables(owner).items() if name in owner.public_variables}
    local.update(variables(module))
    local.update({param.name.lower(): (param.declared, param.is_array, module) for param in procedure.params})
    if procedure.kind in ('function', 'get'):
        local[procedure.name.lower()] = (procedure.returns, procedure.returns_array, module)
    for node in body:
        if isinstance(node, A.Dim):
            local.update({decl.name.lower(): (decl.declared, decl.is_array, module) for decl in node.decls})
    for decl in implicit_arrays:
        local.setdefault(decl.name.lower(), (decl.declared, True, module))

    def runtime_for(name: str) -> ModuleRuntime | None:
        return next((owner for owner in interpreter.modules.values() if owner.name.lower() == name.lower()), None)

    def function_info(owner: ModuleRuntime, name: str) -> Info | None:
        found = owner.procedure(name, 'get') or owner.procedure(name)
        if found is not None:
            if found.kind in ('let', 'set') and found.params:
                return found.params[-1].declared, found.params[-1].is_array, owner
            return found.returns, found.returns_array, owner
        return None

    constants: dict[str, Info] = {}
    def constant_type(declared: str, value: A.Expr, owner: ModuleRuntime) -> str:
        if declared != 'Variant':
            return declared
        # Fold only pure literal operators, never execute a function or
        # initialize a module merely to infer a reference's type.
        if all(isinstance(node, (A.Literal, A.Unary, A.Binary)) for node in nodes(value)):
            try:
                return type_name(interpreter.evaluate_constant(value, owner))
            except VBARuntimeError:
                pass
        return ''  # General constant-expression typing remains incremental.

    owners = [owner for owner in interpreter.modules.values() if owner is not module] + [module]
    for owner in owners:
        if owner is module or not owner.is_class:
            for group in owner.parsed.constants:
                if owner is module or group.scope == 'public':
                    for name, declared, value in group.names:
                        constants[name.lower()] = constant_type(declared, value, owner), False, owner
    for node in body:
        if isinstance(node, A.ConstDecl):
            for name, declared, value in node.names:
                constants[name.lower()] = constant_type(declared, value, module), False, module

    def name_info(name: str, *, declare: bool) -> Info | None:
        key = strip_suffix(name).lower()
        if key == 'me' and module.is_class:
            return module.name, False, module
        if key in local:
            return local[key]
        if key in constants:
            return constants[key]
        found = function_info(module, key)
        if found is not None:
            return found
        for owner in interpreter.modules.values():
            called = owner.procedure(key)
            if not owner.is_class and called is not None and called.scope == 'public':
                return called.returns, called.returns_array, owner
        if declare and not module.parsed.option_explicit and runtime_for(key) is None:
            local[key] = default_type(name, module.parsed.default_types), False, module
            return local[key]
        return None

    def check(name: str, info: Info | None, line: int) -> None:
        suffix = SUFFIX_TYPES.get(name[-1:])
        if suffix is None or info is None:
            return
        declared = info[0]
        if not declared:
            return
        if declared.casefold().startswith('string*'):
            declared = 'String'
        if declared.casefold() != suffix.casefold():
            raise VBACompileError('Type-declaration character does not match declared data type',
                                  where=f'{module.name}.{procedure.name} line {line}')

    def member_info(expression: A.Member, context: Info | None) -> Info | None:
        key = strip_suffix(expression.name).lower()
        if isinstance(expression.target, A.Name):
            parent_key = strip_suffix(expression.target.name).lower()
            owner = runtime_for(parent_key) if parent_key not in local and parent_key not in constants else None
            if owner is not None and not owner.is_class:
                return variables(owner).get(key) or function_info(owner, key)
        parent = visit(expression.target, context) if expression.target is not None else context
        if parent is None or parent[1]:
            return None
        record = interpreter.record_type(parent[0], parent[2])
        if record is not None:
            definition, owner = record
            field = next((field for field in definition.fields if field.name.lower() == key), None)
            if field is not None:
                return field.declared, field.is_array, owner
        owner = runtime_for(parent[0])
        return (variables(owner).get(key) or function_info(owner, key)) if owner is not None else None

    def visit(value: object, context: Info | None = None, *, callee: bool = False) -> Info | None:
        if isinstance(value, A.NewExpr):
            return value.type_name, False, module
        if isinstance(value, A.Name):
            key = strip_suffix(value.name).lower()
            info = name_info(value.name, declare=not callee)
            suffix = SUFFIX_TYPES.get(value.name[-1:])
            # These measured library functions have compiler-enforced suffix
            # rules. Other core intrinsics can ignore arbitrary type characters.
            if info is None and callee and suffix is not None and key in ('trim', 'left', 'round'):
                if key == 'round' or suffix != 'String':
                    raise VBACompileError('Type-declaration character does not match declared data type',
                                          where=f'{module.name}.{procedure.name} line {value.line}')
            check(value.name, info, value.line)
            return info
        if isinstance(value, A.Member):
            info = member_info(value, context)
            check(value.name, info, value.line)
            return info
        if isinstance(value, A.Index):
            info = visit(value.target, context, callee=True)
            visit(value.args, context)
            return (info[0], False, info[2]) if info is not None and info[1] else info
        if isinstance(value, A.CallStmt):
            visit(value.callee, context, callee=True)
            visit(value.args, context)
        elif isinstance(value, A.With):
            parent = visit(value.subject, context)
            visit(value.body, parent)
        elif isinstance(value, (list, tuple)):
            for child in cast(Sequence[object], value):
                visit(child, context)
        elif isinstance(value, (A.Stmt, A.Expr, A.Argument, A.VarDecl, A.IfBranch, A.CaseBlock, A.CaseClause)):
            for attribute in fields(value):
                visit(getattr(value, attribute.name), context)
        return None

    visit(procedure.body)
    if cacheable:
        module.reference_checks[id(procedure)] = scope
