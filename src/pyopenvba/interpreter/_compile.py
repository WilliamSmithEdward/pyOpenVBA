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

Info = tuple[str, bool]

def nodes(value: object) -> Iterator[object]:
    if isinstance(value, (list, tuple)):
        for item in cast(Sequence[object], value):
            yield from nodes(item)
    elif is_dataclass(value) and not isinstance(value, type):
        yield value
        for attribute in fields(value):
            yield from nodes(getattr(value, attribute.name))

def variables(module: ModuleRuntime) -> dict[str, Info]:
    return {decl.name.lower(): (decl.declared, decl.is_array)
            for group in module.parsed.variables for decl in group.decls}

def validate_record_arguments(interpreter: Interpreter, module: ModuleRuntime, procedure: A.Procedure) -> None:
    if not any(runtime.types for runtime in interpreter.modules.values()):
        return
    scope = tuple((runtime.name, ref(runtime)) for runtime in interpreter.modules.values())
    if module.record_checks.get(id(procedure)) == scope:
        return
    body = list(nodes(procedure.body))
    local = variables(module)
    local.update({param.name.lower(): (param.declared, param.is_array) for param in procedure.params})
    local[procedure.name.lower()] = (procedure.returns, False)
    for node in body:
        if isinstance(node, (A.Dim, A.ReDim)):
            local.update({decl.name.lower(): (decl.declared, decl.is_array) for decl in node.decls})
    # Declaration errors are compile errors even when their statement is
    # unreachable. Leave unused procedures deferred, as native VBA does.
    for declared, _ in local.values():
        interpreter.record_type(declared, module)

    def runtime_for(name: str) -> ModuleRuntime | None:
        return next((runtime for runtime in interpreter.modules.values() if runtime.name.lower() == name.lower()), None)

    def record_for(name: str, owner: ModuleRuntime) -> A.TypeDef | None:
        record = interpreter.record_type(name, owner)
        return record[0] if record is not None else None

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
        if isinstance(expression, A.Member):
            if isinstance(expression.target, A.Name) and expression.target.name.lower() not in local:
                target = runtime_for(expression.target.name)
                if target is not None:
                    info = variables(target).get(expression.name.lower())
                    if info is not None:
                        return info
            parent = expression_info(expression.target)
            if parent is not None and not parent[1]:
                defined = record_for(parent[0], module)
                if defined is not None:
                    field = next((field for field in defined.fields if field.name.lower() == expression.name.lower()), None)
                    if field is not None:
                        return field.declared, field.is_array
                target = runtime_for(parent[0])
                if target is not None:
                    info = variables(target).get(expression.name.lower())
                    if info is not None:
                        return info
        if isinstance(expression, A.Index):
            info = expression_info(expression.target)
            if info is not None and info[1]:
                return info[0], False
            call = callable_for(expression.target)
            if call is not None:
                return call[0].returns, False
        return None

    for node in body:
        if isinstance(node, A.CallStmt):
            callee, arguments = node.callee, node.args
        elif isinstance(node, A.Index):
            callee, arguments = node.target, node.args
        else:
            continue
        if not any(argument.by_value for argument in arguments):
            continue
        call = callable_for(callee)
        if call is None:
            continue
        called, owner = call
        for index, argument in enumerate(arguments):
            if not argument.by_value:
                continue
            parameter = next((param for param in called.params if param.name.lower() == argument.name.lower()), None) if argument.name else called.params[index] if index < len(called.params) else None
            if parameter is None or parameter.by_val or parameter.is_array or parameter.param_array:
                continue
            if record_for(parameter.declared, owner) is None:
                continue
            info = expression_info(argument.value)
            if info is not None and not info[1] and record_for(info[0], module) is not None:
                raise VBACompileError("Variable required - can't assign to this expression",
                                      where=f"{module.name}.{procedure.name} line {getattr(node, 'line', procedure.line)}")
    module.record_checks[id(procedure)] = scope
