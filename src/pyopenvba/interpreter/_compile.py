"""Demand validation of statically typed calls before a procedure executes.

This is an incremental compiler surface, not a complete VBA type checker.
"""
from __future__ import annotations
from collections.abc import Iterator, Sequence
from dataclasses import fields, is_dataclass
from typing import TYPE_CHECKING, cast
from weakref import ref
from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._flow import nested_bodies
from pyopenvba.interpreter._lex import SUFFIX_TYPES, default_type, strip_suffix
if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter, ModuleRuntime

Info = tuple[str, bool, 'ModuleRuntime']
VARIANT_RECORD_ERROR = ('Only user-defined types defined in public object modules can be coerced to or from '
                        'a variant or passed to late-bound functions')

def nodes(value: object) -> Iterator[object]:
    if isinstance(value, (list, tuple)):
        for item in cast(Sequence[object], value):
            yield from nodes(item)
    elif is_dataclass(value) and not isinstance(value, type):
        yield value
        for attribute in fields(value):
            yield from nodes(getattr(value, attribute.name))

def variables(module: ModuleRuntime) -> dict[str, Info]:
    return {decl.name.lower(): (decl.declared, decl.is_array, module)
            for group in module.parsed.variables for decl in group.decls}

def record_byval_signature(interpreter: Interpreter, module: ModuleRuntime, procedure: A.Procedure) -> bool:
    return any(param.by_val and interpreter.record_type(param.declared, module) is not None
               for param in procedure.params)


def validate_numeric_loops(interpreter: Interpreter, module: ModuleRuntime, procedure: A.Procedure,
                           implicit_arrays: Sequence[A.VarDecl]) -> None:
    """A numeric counter must name numeric storage, not a property value.

    Keep this demand check before argument binding and execution, including
    counters in unreachable blocks. With can hold a temporary UDT as well
    as an ordinary variable; its fields still provide writable storage.
    """
    # program_for has just refreshed the project's symbol scope. Reuse its
    # snapshot instead of rebuilding weak references on every VBA call.
    scope = module.program_scope
    # Host-created synthetic procedures are not retained by the module and
    # can reuse ids; validate them without retaining a reusable-id cache key.
    cacheable = id(procedure) in module.programs
    if cacheable and module.numeric_checks.get(id(procedure)) == scope:
        return
    body = list(nodes(procedure.body))
    if not any(isinstance(node, A.For) for node in body):
        if cacheable:
            module.numeric_checks[id(procedure)] = scope
        return
    local = {name: info for runtime in interpreter.modules.values() if not runtime.is_class and runtime is not module
             for name, info in variables(runtime).items() if name in runtime.public_variables}
    local.update(variables(module))
    local.update({param.name.lower(): (param.declared, param.is_array, module) for param in procedure.params})
    if procedure.kind in ('function', 'get'):
        local[procedure.name.lower()] = (procedure.returns, procedure.returns_array, module)
    for node in body:
        if isinstance(node, A.Dim):
            local.update({decl.name.lower(): (decl.declared, decl.is_array, module) for decl in node.decls})
    for decl in implicit_arrays:
        local.setdefault(decl.name.lower(), (decl.declared, True, module))
    constants = {strip_suffix(name).lower() for node in body if isinstance(node, A.ConstDecl)
                 for name, _, _ in node.names}
    for runtime in interpreter.modules.values():
        if runtime is module or not runtime.is_class:
            constants.update(name.lower() for group in runtime.parsed.constants
                             if runtime is module or group.scope == 'public' for name, _, _ in group.names)
            constants.update(name.lower() for definition in runtime.parsed.enums
                             if runtime is module or definition.scope == 'public' for name, _ in definition.members)

    def runtime_for(name: str) -> ModuleRuntime | None:
        return next((runtime for runtime in interpreter.modules.values() if runtime.name.lower() == name.lower()), None)

    def member_info(expression: A.Member, context: Info | None) -> Info | None:
        if isinstance(expression.target, A.Name) and strip_suffix(expression.target.name).lower() not in local:
            owner = runtime_for(expression.target.name)
            if owner is not None and not owner.is_class:
                key = strip_suffix(expression.name).lower()
                if owner is module or key in owner.public_variables:
                    return variables(owner).get(key)
                return None
        parent = expression_info(expression.target, context) if expression.target is not None else context
        if parent is not None and not parent[1]:
            record = interpreter.record_type(parent[0], parent[2])
            if record is not None:
                definition, owner = record
                field = next((field for field in definition.fields if field.name.lower() == strip_suffix(expression.name).lower()), None)
                if field is not None:
                    return field.declared, field.is_array, owner
        # Class/document members, including public fields, are properties.
        return None

    def callable_info(expression: A.Expr | None) -> Info | None:
        name = ''
        owner = module
        if isinstance(expression, A.Name):
            name = strip_suffix(expression.name)
            if name.lower() in local:
                return None
            found = module.procedure(name)
            if found is not None:
                return found.returns, found.returns_array, module
            for candidate in interpreter.modules.values():
                found = candidate.procedure(name)
                if not candidate.is_class and found is not None and found.scope == 'public':
                    return found.returns, found.returns_array, candidate
        elif isinstance(expression, A.Member) and isinstance(expression.target, A.Name):
            key = strip_suffix(expression.target.name).lower()
            parent = local.get(key)
            candidate = runtime_for(parent[0] if parent is not None else key)
            if candidate is not None:
                owner, name = candidate, strip_suffix(expression.name)
        found = owner.procedure(name)
        return (found.returns, found.returns_array, owner) if found is not None else None

    def expression_info(expression: A.Expr | None, context: Info | None) -> Info | None:
        if isinstance(expression, A.Name):
            return local.get(strip_suffix(expression.name).lower()) or callable_info(expression)
        if isinstance(expression, A.Member):
            return member_info(expression, context) or callable_info(expression)
        if isinstance(expression, A.Index):
            parent = expression_info(expression.target, context)
            if parent is not None and parent[1]:
                return parent[0], False, parent[2]
            return callable_info(expression.target)
        return None

    def counter_info(expression: A.Expr | None, context: Info | None) -> Info | None:
        if isinstance(expression, A.Name):
            key = strip_suffix(expression.name).lower()
            if key in local:
                return local[key]
            if key in constants or callable_info(expression) is not None or module.parsed.option_explicit:
                return None
            return default_type(expression.name, module.parsed.default_types), False, module
        if isinstance(expression, A.Member):
            return member_info(expression, context)
        if isinstance(expression, A.Index):
            parent = counter_info(expression.target, context)
            if parent is not None and parent[1]:
                return parent[0], False, parent[2]
        return None

    implicit: set[str] = set()

    def infer_names(value: object) -> None:
        if isinstance(value, A.Name):
            key = strip_suffix(value.name).lower()
            suffix = SUFFIX_TYPES.get(value.name[-1:])
            if key in implicit and suffix is not None and local[key][0] != suffix:
                raise VBACompileError('Type-declaration character does not match declared data type',
                                      where=f'{module.name}.{procedure.name} line {value.line}')
            if (not module.parsed.option_explicit and key not in local and key not in constants
                    and runtime_for(key) is None and callable_info(value) is None):
                local[key] = default_type(value.name, module.parsed.default_types), False, module
                implicit.add(key)
        elif isinstance(value, A.ConstDecl):
            return  # A constant expression cannot introduce implicit storage.
        elif isinstance(value, (list, tuple)):
            for child in cast(Sequence[object], value):
                infer_names(child)
        elif isinstance(value, (A.Stmt, A.Expr, A.Argument, A.IfBranch, A.CaseBlock, A.CaseClause)):
            for attribute in fields(value):
                child = getattr(value, attribute.name)
                # Calls do not declare their bare function names as variables.
                if ((isinstance(value, A.Index) and attribute.name == 'target') or
                        (isinstance(value, A.CallStmt) and attribute.name == 'callee')) and isinstance(child, A.Name):
                    continue
                infer_names(child)

    infer_names(procedure.body)

    def validate(body: Sequence[A.Stmt], context: Info | None = None) -> None:
        for statement in body:
            if isinstance(statement, A.For):
                info = counter_info(statement.target, context)
                if info is None or info[1]:
                    raise VBACompileError("Variable required - can't assign to this expression",
                                          where=f'{module.name}.{procedure.name} line {statement.line}')
                declared, _, owner = info
                if declared not in ('', 'Variant', 'Byte', 'Integer', 'Long', 'LongLong', 'LongPtr',
                                     'Single', 'Double', 'Currency', 'Date') and not interpreter.is_enum_type(declared, owner):
                    raise VBACompileError('For control variable must be numeric',
                                          where=f'{module.name}.{procedure.name} line {statement.line}')
            if isinstance(statement, A.With):
                validate(statement.body, expression_info(statement.subject, context))
            else:
                for child in nested_bodies(statement):
                    validate(child, context)

    validate(procedure.body)
    if cacheable:
        module.numeric_checks[id(procedure)] = scope


def validate_record_arguments(interpreter: Interpreter, module: ModuleRuntime, procedure: A.Procedure) -> None:
    if not any(runtime.types for runtime in interpreter.modules.values()):
        return
    scope = tuple((runtime.name, ref(runtime)) for runtime in interpreter.modules.values())
    if module.record_checks.get(id(procedure)) == scope:
        return
    if record_byval_signature(interpreter, module, procedure):
        raise VBACompileError('User-defined type may not be passed ByVal', where=f'{module.name}.{procedure.name}')
    body = list(nodes(procedure.body))
    local = variables(module)
    local.update({param.name.lower(): (param.declared, param.is_array, module) for param in procedure.params})
    local[procedure.name.lower()] = (procedure.returns, procedure.returns_array, module)
    for node in body:
        if isinstance(node, A.Dim):
            local.update({decl.name.lower(): (decl.declared, decl.is_array, module) for decl in node.decls})
        elif isinstance(node, A.ReDim):
            # A ReDim of a record's field (item.B) declares no local.
            local.update({decl.name.lower(): (decl.declared, decl.is_array, module)
                          for index, decl in enumerate(node.decls)
                          if index >= len(node.targets) or node.targets[index] is None})
    # Declaration errors are compile errors even when their statement is
    # unreachable. Leave unused procedures deferred, as native VBA does.
    for declared, _, owner in local.values():
        interpreter.record_type(declared, owner)

    def runtime_for(name: str) -> ModuleRuntime | None:
        return next((runtime for runtime in interpreter.modules.values() if runtime.name.lower() == name.lower()), None)

    def record_for(name: str, owner: ModuleRuntime) -> A.TypeDef | None:
        record = interpreter.record_type(name, owner)
        return record[0] if record is not None else None

    def cannot_marshal(info: Info) -> bool:
        record = interpreter.record_type(info[0], info[2])
        return record is not None and (not record[1].is_class or record[0].scope != 'public')

    def callable_for(expression: A.Expr | None) -> tuple[A.Procedure, ModuleRuntime] | None:
        if isinstance(expression, A.Name) and (expression.name.lower() not in local or expression.name.lower() == procedure.name.lower()):
            found = module.procedure(expression.name)
            if found is not None:
                return found, module
            for runtime in interpreter.modules.values():
                found = runtime.procedure(expression.name)
                if not runtime.is_class and found is not None and found.scope == 'public':
                    return found, runtime
        if isinstance(expression, A.Member) and isinstance(expression.target, A.Name):
            name = expression.target.name.lower()
            target = runtime_for(local[name][0]) if name in local else runtime_for(name)
            if target is not None:
                found = target.procedure(expression.name)
                if found is not None:
                    return found, target
        return None

    def expression_info(expression: A.Expr | None) -> Info | None:
        if isinstance(expression, A.Name):
            key = expression.name.lower()
            if key in local:
                return local[key]
            for runtime in interpreter.modules.values():
                info = variables(runtime).get(key)
                if not runtime.is_class and info is not None:
                    return info
            call = callable_for(expression)
            if call is not None:
                return call[0].returns, False, call[1]
        if isinstance(expression, A.Member):
            if isinstance(expression.target, A.Name) and expression.target.name.lower() not in local:
                target = runtime_for(expression.target.name)
                if target is not None:
                    info = variables(target).get(expression.name.lower())
                    if info is not None:
                        return info
            parent = expression_info(expression.target)
            if parent is not None and not parent[1]:
                record = interpreter.record_type(parent[0], parent[2])
                if record is not None:
                    defined, owner = record
                    field = next((field for field in defined.fields if field.name.lower() == expression.name.lower()), None)
                    if field is not None:
                        return field.declared, field.is_array, owner
                target = runtime_for(parent[0])
                if target is not None:
                    info = variables(target).get(expression.name.lower())
                    if info is not None:
                        return info
            call = callable_for(expression)
            if call is not None:
                return call[0].returns, False, call[1]
        if isinstance(expression, A.Index):
            info = expression_info(expression.target)
            if info is not None and info[1]:
                return info[0], False, info[2]
            call = callable_for(expression.target)
            if call is not None:
                return call[0].returns, False, call[1]
        return None

    for node in body:
        if isinstance(node, A.Assign) and node.kind == 'let':
            target, value = expression_info(node.target), expression_info(node.value)
            if target is not None and value is not None:
                if ((target[0].lower() == 'variant' and cannot_marshal(value)) or
                        (value[0].lower() == 'variant' and cannot_marshal(target))):
                    raise VBACompileError(VARIANT_RECORD_ERROR, where=f'{module.name}.{procedure.name} line {node.line}')
            if target is not None and value is not None and not target[1] and not value[1]:
                target_record, value_record = record_for(target[0], target[2]), record_for(value[0], value[2])
                if target_record is not None and value_record is not None and target_record is not value_record:
                    raise VBACompileError('Type mismatch', where=f'{module.name}.{procedure.name} line {node.line}')
        if isinstance(node, A.CallStmt):
            callee, arguments = node.callee, node.args
        elif isinstance(node, A.Index):
            callee, arguments = node.target, node.args
        else:
            continue
        call = callable_for(callee)
        if call is None:
            if isinstance(callee, A.Name) and callee.name.lower() == 'typename':
                for argument in arguments:
                    info = expression_info(argument.value)
                    if info is not None and cannot_marshal(info):
                        raise VBACompileError(VARIANT_RECORD_ERROR, where=f'{module.name}.{procedure.name} line {getattr(node, "line", procedure.line)}')
            continue
        called, owner = call
        if record_byval_signature(interpreter, owner, called):
            raise VBACompileError('User-defined type may not be passed ByVal', where=f'{owner.name}.{called.name}')
        for index, argument in enumerate(arguments):
            parameter = next((param for param in called.params if param.name.lower() == argument.name.lower()), None) if argument.name else called.params[index] if index < len(called.params) else None
            if parameter is None or parameter.param_array:
                continue
            info = expression_info(argument.value)
            if info is not None and parameter.declared.lower() == 'variant' and cannot_marshal(info):
                raise VBACompileError(VARIANT_RECORD_ERROR, where=f'{module.name}.{procedure.name} line {getattr(node, "line", procedure.line)}')
            if parameter.by_val:
                continue
            expected = record_for(parameter.declared, owner)
            if expected is None:
                continue
            if info is not None:
                actual = record_for(info[0], info[2])
                if parameter.is_array and (not info[1] or argument.by_value):
                    raise VBACompileError('Type mismatch: array or user-defined type expected', where=f'{module.name}.{procedure.name} line {getattr(node, "line", procedure.line)}')
                if actual is not None and (actual is not expected or info[1] != parameter.is_array):
                    raise VBACompileError('ByRef argument type mismatch', where=f'{module.name}.{procedure.name} line {getattr(node, "line", procedure.line)}')
                if actual is not None and not info[1] and argument.by_value:
                    raise VBACompileError("Variable required - can't assign to this expression",
                                      where=f"{module.name}.{procedure.name} line {getattr(node, 'line', procedure.line)}")
    module.record_checks[id(procedure)] = scope
