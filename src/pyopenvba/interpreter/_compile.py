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
    local[procedure.name.lower()] = (procedure.returns, False, module)
    for node in body:
        if isinstance(node, (A.Dim, A.ReDim)):
            local.update({decl.name.lower(): (decl.declared, decl.is_array, module) for decl in node.decls})
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
