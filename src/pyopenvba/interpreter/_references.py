"""Demand checks for statically resolved VBA calls and type characters."""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import fields
from typing import TYPE_CHECKING, cast

from pyopenvba.exceptions import VBACompileError, VBARuntimeError
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._compile import Info, nodes, variables
from pyopenvba.interpreter._flow import Program
from pyopenvba.interpreter._financial import SIGNATURES as FINANCIAL_SIGNATURES
from pyopenvba.interpreter._lex import SUFFIX_TYPES, default_type, strip_suffix
from pyopenvba.interpreter._values import type_name

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter, ModuleRuntime


def validate_references(interpreter: Interpreter, module: ModuleRuntime, procedure: A.Procedure,
                        program: Program) -> None:
    scope = module.program_scope
    cacheable = id(procedure) in module.programs
    if cacheable and module.reference_checks.get(id(procedure)) == scope:
        return
    from pyopenvba.interpreter._runtime import UNRESOLVED, vba_constant

    body = list(nodes(procedure.body))
    program.late_bound_members.clear()
    local = {name: info for owner in interpreter.modules.values() if not owner.is_class and owner is not module
             for name, info in variables(owner).items() if name in owner.public_variables}
    local.update(variables(module))
    local.update({param.name.lower(): (param.declared, param.is_array, module) for param in procedure.params})
    if procedure.kind in ('function', 'get'):
        local[procedure.name.lower()] = (procedure.returns, procedure.returns_array, module)
    for node in body:
        if isinstance(node, A.Dim):
            local.update({decl.name.lower(): (decl.declared, decl.is_array, module) for decl in node.decls})
    for decl in program.arrays:
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
        if key in FINANCIAL_SIGNATURES:
            return 'Double', False, module
        constant = interpreter.host.constant(key)
        if constant is UNRESOLVED:
            constant = vba_constant(key)
        if constant is not UNRESOLVED:
            constants[key] = type_name(constant), False, module
            return constants[key]
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
        if parent is None or parent[0] in ('', 'Object', 'Variant'):
            program.late_bound_members.add(id(expression))
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

    def signature(expression: A.Expr | None, context: Info | None, *, explicit: bool = False,
                  setter: str = '') -> tuple[A.Procedure, ModuleRuntime] | None:
        owner = module
        key = ''
        if isinstance(expression, A.Name):
            key = strip_suffix(expression.name).lower()
            if key in local and not (explicit and key == procedure.name.lower()):
                return None
            found = owner.procedure(key, setter) if setter else owner.procedure(key, 'get') or owner.procedure(key)
            if found is not None:
                return found, owner
            for candidate in interpreter.modules.values():
                found = candidate.procedure(key, setter) if setter else candidate.procedure(key)
                if not candidate.is_class and found is not None and found.scope == 'public':
                    return found, candidate
            if not setter and key in FINANCIAL_SIGNATURES:
                return FINANCIAL_SIGNATURES[key], module
            return None
        if isinstance(expression, A.Member):
            key = strip_suffix(expression.name).lower()
            owner = None
            if isinstance(expression.target, A.Name):
                parent_key = strip_suffix(expression.target.name).lower()
                if parent_key not in local and parent_key not in constants:
                    candidate = runtime_for(parent_key)
                    if candidate is not None and not candidate.is_class:
                        owner = candidate
            if owner is None:
                parent = visit(expression.target, context) if expression.target is not None else context
                if parent is not None and not parent[1]:
                    owner = runtime_for(parent[0])
            if owner is not None:
                found = owner.procedure(key, setter) if setter else owner.procedure(key, 'get') or owner.procedure(key)
                if found is not None:
                    return found, owner
        return None

    def fail(message: str, line: int) -> None:
        raise VBACompileError(message, where=f'{module.name}.{procedure.name} line {line}')

    def byref_kind(info: Info) -> str:
        declared, _, owner = info
        if interpreter.is_enum_type(declared, owner):
            return 'Long'
        if declared.startswith('String*'):
            return 'String'
        if declared == 'LongPtr':
            return 'LongLong'  # The in-memory runtime models 64-bit VBA.
        return declared

    def check_call(expression: A.Expr | None, args: Sequence[A.Argument], context: Info | None,
                   *, explicit: bool = False, setter: str = '') -> None:
        line = expression.line if expression is not None else procedure.line
        named_seen = False
        for argument in args:
            if argument.name:
                named_seen = True
            elif named_seen:
                fail('Expected: named argument', line)
        called = signature(expression, context, explicit=explicit, setter=setter)
        if called is None:
            return
        target, owner = called
        params = target.params
        if any(parameter.param_array for parameter in params) and any(argument.name for argument in args):
            fail('Named arguments not allowed with ParamArray', line)
        if args and not params and target.kind in ('function', 'get') and not target.returns_array:
            if target.returns == 'Collection':
                if len(args) != 1 or (args[0].name and args[0].name.lower() != 'index') or args[0].omitted:
                    fail('Wrong number of arguments or invalid property assignment', line)
                return
            if target.returns in ('Object', 'Variant') or runtime_for(target.returns) is not None:
                if any(argument.name for argument in args):
                    fail('Named argument not found', line)
                return
        assigned: set[int] = set()
        for index, argument in enumerate(args):
            if argument.name:
                position = next((i for i, param in enumerate(params) if param.name.lower() == argument.name.lower()), -1)
                if position < 0 or params[position].param_array:
                    fail('Named argument not found', line)
            else:
                position = min(index, len(params) - 1) if params and params[-1].param_array else index
            if position >= len(params):
                fail('Wrong number of arguments or invalid property assignment', line)
            parameter = params[position]
            if position in assigned and not parameter.param_array:
                fail('Named argument already specified', line)
            assigned.add(position)
            if argument.omitted:
                if not parameter.optional:
                    fail('Argument not optional', line)
                continue
            if parameter.is_array and not parameter.param_array and argument.by_value:
                fail('Type mismatch: array or user-defined type expected', line)
            if parameter.by_val or parameter.param_array or argument.by_value:
                continue
            actual = visit(argument.value, context)
            storage = isinstance(argument.value, A.Name) and strip_suffix(argument.value.name).lower() in local
            if isinstance(argument.value, A.Member):
                member = argument.value
                parent = visit(member.target, context) if member.target is not None else context
                storage = parent is not None and interpreter.record_type(parent[0], parent[2]) is not None
                if isinstance(member.target, A.Name) and strip_suffix(member.target.name).lower() not in local:
                    candidate = runtime_for(member.target.name)
                    storage = candidate is not None and not candidate.is_class and strip_suffix(member.name).lower() in variables(candidate)
            if isinstance(argument.value, A.Index):
                parent = visit(argument.value.target, context, callee=True)
                storage = parent is not None and parent[1]
            if parameter.is_array and (actual is None or not actual[1]):
                fail('Type mismatch: array or user-defined type expected', line)
            if parameter.is_array and actual is not None and byref_kind((parameter.declared, True, owner)) != byref_kind(actual):
                fail('ByRef argument type mismatch', line)
            if actual is not None and storage:
                expected = byref_kind((parameter.declared, parameter.is_array, owner))
                actual_kind = byref_kind(actual)
                primitives = ('Byte', 'Integer', 'Long', 'LongLong', 'Single', 'Double', 'Currency', 'Date', 'Boolean', 'String', 'Variant')
                if parameter.is_array and expected != actual_kind:
                    fail('ByRef argument type mismatch', line)
                if (actual[1] and not parameter.is_array and expected != 'Variant') or (
                        expected in primitives and actual_kind in primitives and expected != 'Variant' and expected != actual_kind):
                    fail('ByRef argument type mismatch', line)
        for index, parameter in enumerate(params):
            if index not in assigned and not parameter.optional and parameter.default is None and not parameter.param_array:
                fail('Argument not optional', line)

    resolved: dict[tuple[int, Info | None, bool, bool], Info | None] = {}

    def visit(value: object, context: Info | None = None, *, callee: bool = False, lvalue: bool = False) -> Info | None:
        if isinstance(value, A.Expr):
            key = id(value), context, callee, lvalue
            if key not in resolved:
                resolved[key] = visit_node(value, context, callee=callee, lvalue=lvalue)
            return resolved[key]
        return visit_node(value, context, callee=callee, lvalue=lvalue)

    def visit_node(value: object, context: Info | None = None, *, callee: bool = False, lvalue: bool = False) -> Info | None:
        if isinstance(value, A.NewExpr):
            return value.type_name, False, module
        if isinstance(value, A.Name):
            key = strip_suffix(value.name).lower()
            info = name_info(value.name, declare=not callee or lvalue)
            suffix = SUFFIX_TYPES.get(value.name[-1:])
            # These measured library functions have compiler-enforced suffix
            # rules. Other core intrinsics can ignore arbitrary type characters.
            if info is None and callee and suffix is not None and key in ('trim', 'left', 'round'):
                if key == 'round' or suffix != 'String':
                    raise VBACompileError('Type-declaration character does not match declared data type',
                                          where=f'{module.name}.{procedure.name} line {value.line}')
            check(value.name, info, value.line)
            if not callee:
                check_call(value, [], context)
            return info
        if isinstance(value, A.Member):
            info = member_info(value, context)
            check(value.name, info, value.line)
            if not callee:
                check_call(value, [], context)
            return info
        if isinstance(value, A.Index):
            info = visit(value.target, context, callee=True)
            visit(value.args, context)
            check_call(value.target, value.args, context, explicit=True)
            return (info[0], False, info[2]) if info is not None and info[1] and value.args else info
        if isinstance(value, A.CallStmt):
            visit(value.callee, context, callee=True)
            visit(value.args, context)
            check_call(value.callee, value.args, context, explicit=True)
        elif isinstance(value, A.Assign):
            target = value.target
            arguments: list[A.Argument] = []
            if isinstance(target, A.Index):
                arguments = list(target.args)
                visit(arguments, context)
                target = target.target
            visit(target, context, callee=True, lvalue=True)
            visit(value.value, context)
            kind = 'set' if value.kind == 'set' else 'let'
            called = signature(target, context, setter=kind)
            if called is None:
                check_call(target, arguments, context, setter=kind)
            else:
                name = called[0].params[-1].name if called[0].params and any(arg.name for arg in arguments) else ''
                rhs = A.Argument(value=value.value, name=name, by_value=True)
                check_call(target, [*arguments, rhs], context, setter=kind)
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
