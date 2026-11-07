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
from pyopenvba.interpreter._intrinsics import INTRINSICS
from pyopenvba.interpreter._lex import SUFFIX_TYPES, default_type, strip_suffix
from pyopenvba.interpreter._type_library import CallShape, LibraryClass, LibraryType, find_class, indexed_refusal, let_refusal, refusal
from pyopenvba.interpreter._values import type_name

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter, ModuleRuntime


_FIXED_TYPES = frozenset({'Byte', 'Integer', 'Long', 'LongLong', 'LongPtr', 'Single', 'Double', 'Currency', 'Date', 'Boolean'})
#: Declared to answer one of those, where Date and Now answer a Variant and Len(Date) compiles (tests/fixtures/vba_everyday_17.json).
_FIXED_INTRINSICS = frozenset({'cbool', 'cbyte', 'ccur', 'cdate', 'cdbl', 'cint', 'clng', 'clnglng', 'clngptr', 'csng', 'timer'})


def validate_references(interpreter: Interpreter, module: ModuleRuntime, procedure: A.Procedure,
                        program: Program) -> None:
    scope = module.program_scope
    cacheable = id(procedure) in module.programs
    if cacheable and module.reference_checks.get(id(procedure)) == scope:
        return
    from pyopenvba.interpreter._runtime import UNRESOLVED, vba_constant

    body = list(nodes(procedure.body))
    program.late_bound_members.clear()
    program.early_bound_members.clear()
    program.stored_len_arguments.clear()
    library: str | None = getattr(interpreter.host, 'vba_type_library', None)
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

    def library_class(declared: str, owner: ModuleRuntime) -> LibraryClass | None:
        """The class a type means, unless the project defines that name itself."""
        if library is None or not declared or runtime_for(declared) is not None or interpreter.record_type(declared, owner) is not None:
            return None
        return find_class(library, declared)

    def library_info(found: str | None) -> Info | None:
        return None if found is None else (LibraryType(found), False, module)

    def global_member(key: str) -> Info | None:
        """A name no declaration claims that the host's library answers: Range, ActiveSheet, Selection."""
        found = find_class(library, '_Global') if library is not None else None
        if found is None:
            return None
        if key == 'thisworkbook' and runtime_for(key) is None:
            return library_info('Workbook')
        return library_info(found.result(key))

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
        found = global_member(key)
        if found is not None:
            return found
        if declare and not module.parsed.option_explicit and runtime_for(key) is None:
            local[key] = default_type(name, module.parsed.default_types), False, module
            return local[key]
        return None

    def check(name: str, info: Info | None, line: int) -> None:
        suffix = SUFFIX_TYPES.get(name[-1:])
        if suffix is None or info is None:
            return
        declared = info[0]
        if not declared or isinstance(declared, LibraryType):
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
        if owner is not None:
            return variables(owner).get(key) or function_info(owner, key)
        found = library_class(parent[0], parent[2])
        if found is None:
            return None
        program.early_bound_members[id(expression)] = found.name
        return library_info(found.result(key))

    def shape_of(args: Sequence[A.Argument]) -> CallShape:
        given = [argument for argument in args if not argument.name]
        while given and given[-1].omitted:
            given.pop()
        return CallShape(len(given), tuple(argument.name for argument in args if argument.name),
                         frozenset(at for at, argument in enumerate(given) if isinstance(argument.value, A.Literal)))

    def check_library_call(expression: A.Expr | None, args: Sequence[A.Argument], use: str) -> None:
        """Refuse what VBA's compiler refuses of a member call bound early, which the type library decides."""
        name = program.early_bound_members.get(id(expression)) if isinstance(expression, A.Member) else None
        if library is None or name is None:
            return
        assert isinstance(expression, A.Member)
        found = find_class(library, name)
        if found is None:
            return
        if expression.bracketed and strip_suffix(expression.name).lower() not in found.members and found.evaluates:
            return
        said = refusal(library, found, strip_suffix(expression.name), use, shape_of(args))
        if said is not None:
            fail(said, expression.line)

    def indexed_by_default(target: A.Expr | None, context: Info | None, info: Info) -> bool:
        """Whether the arguments after this go to what it answers: Cells(1, 1) is Cells then its default member."""
        if isinstance(target, A.Member) and id(target) in program.early_bound_members:
            parent = visit(target.target, context) if target.target is not None else context
            found = library_class(parent[0], parent[2]) if parent is not None else None
            return found is not None and not found.takes_arguments(strip_suffix(target.name))
        if isinstance(target, A.Name) and isinstance(info[0], LibraryType):
            found = find_class(library, '_Global') if library is not None else None
            return found is not None and not found.takes_arguments(strip_suffix(target.name))
        return signature(target, context, explicit=True) is None

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
    with_subjects: set[int] = set()

    def has_fixed_type(expression: A.Expr | None, context: Info | None) -> bool:
        """Whether the compiler knows this expression as a number, a date or a Boolean. What it cannot tell, it is not said to be."""
        if isinstance(expression, A.Literal):
            return type_name(expression.value) in _FIXED_TYPES
        if isinstance(expression, A.Unary):
            return has_fixed_type(expression.operand, context)
        if isinstance(expression, A.Binary):
            return expression.op != '&' and has_fixed_type(expression.left, context) and has_fixed_type(expression.right, context)
        if isinstance(expression, A.Index) and isinstance(expression.target, A.Name):
            key = strip_suffix(expression.target.name).lower()
            if key in _FIXED_INTRINSICS and name_info(key, declare=False) is None:
                return True
        if isinstance(expression, A.Name) and strip_suffix(expression.name).lower() in _FIXED_INTRINSICS:
            return name_info(expression.name, declare=False) is None
        if isinstance(expression, (A.Name, A.Member, A.Index)):
            info = visit(expression, context, callee=isinstance(expression, A.Name))
            if isinstance(expression, A.Member) and id(expression) not in program.early_bound_members:
                return False
            return info is not None and not info[1] and info[0] in _FIXED_TYPES
        return False

    def check_len(argument: A.Expr | None, context: Info | None, line: int) -> None:
        """Len(n) reads a variable's storage, and with no variable to read a number is refused: Len(5), Len(n + 1), Len(ws.Index)
        (tests/fixtures/vba_everyday_17.json). A Variant's answer goes by its text at run time, so Len(v + 1) compiles."""
        if isinstance(argument, A.Literal) and type_name(argument.value) == 'Nothing':
            fail('Invalid use of object', line)
        if isinstance(argument, A.Name):
            key = strip_suffix(argument.name).lower()
            if key == 'me' and not module.is_class:
                fail('Invalid use of Me keyword', line)
            if key in local:
                return
        if isinstance(argument, (A.Index, A.Member)):
            info = visit(argument, context)
            indexed = isinstance(argument, A.Index) and isinstance(argument.target, A.Name) and strip_suffix(argument.target.name).lower() in local
            field = isinstance(argument, A.Member) and id(argument) not in program.early_bound_members
            if (indexed or field) and info is not None and not info[1]:
                if info[0] in _FIXED_TYPES:
                    program.stored_len_arguments[id(argument)] = str(info[0])
                return
        if has_fixed_type(argument, context):
            fail('Variable required - can\'t assign to this expression', line)

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
            # A bare CurDir is the function, so a CurDir$ beside it is not a Variant's mismatched type character.
            info = name_info(value.name, declare=(not callee and key not in INTRINSICS) or lvalue)
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
                check_library_call(value, [], 'value')
            return info
        if isinstance(value, A.Index):
            with_subject = id(value) in with_subjects
            info = visit(value.target, context, callee=True)
            visit(value.args, context)
            if (info is None and isinstance(value.target, A.Name) and strip_suffix(value.target.name).lower() in ('len', 'lenb')
                    and len(value.args) == 1):
                check_len(value.args[0].value, context, value.line)
            if not callee and not lvalue and not with_subject:
                check_library_call(value.target, value.args, 'value')
            if isinstance(value.target, A.Name) and info is not None and not info[1]:
                held = library_class(info[0], info[2])
                if held is not None:
                    program.early_bound_members[id(value.target)] = held.name
                    # A variable of the class, not a name its library answers: Range("A1") is no default member's call.
                    if not isinstance(info[0], LibraryType) and library is not None and strip_suffix(value.target.name).lower() in local:
                        said = indexed_refusal(library, held, shape_of(value.args))
                        if said is not None:
                            fail(said, value.line)
            check_call(value.target, value.args, context, explicit=True)
            if info is not None and info[1] and value.args:
                return info[0], False, info[2]
            if info is not None and value.args and indexed_by_default(value.target, context, info):
                found = library_class(info[0], info[2])
                if found is not None:
                    return library_info(found.default_result)
            return info
        if isinstance(value, A.CallStmt):
            visit(value.callee, context, callee=True)
            visit(value.args, context)
            check_call(value.callee, value.args, context, explicit=True)
            check_library_call(value.callee, value.args, 'statement')
        elif isinstance(value, A.Assign):
            target = value.target
            arguments: list[A.Argument] = []
            if isinstance(target, A.Index):
                arguments = list(target.args)
                visit(arguments, context)
                target = target.target
            visit(target, context, callee=True, lvalue=True)
            visit(value.value, context)
            if value.kind == 'let':
                check_library_call(target, arguments, 'assign')
                if isinstance(target, A.Name) and not arguments:
                    declared = local.get(strip_suffix(target.name).lower())
                    held = library_class(declared[0], declared[2]) if declared is not None and not declared[1] else None
                    said = let_refusal(held) if held is not None else None
                    if said is not None:
                        fail(said, target.line)
            kind = 'set' if value.kind == 'set' else 'let'
            called = signature(target, context, setter=kind)
            if called is None:
                check_call(target, arguments, context, setter=kind)
            else:
                name = called[0].params[-1].name if called[0].params and any(arg.name for arg in arguments) else ''
                rhs = A.Argument(value=value.value, name=name, by_value=True)
                check_call(target, [*arguments, rhs], context, setter=kind)
        elif isinstance(value, A.With):
            subject = value.subject
            if isinstance(subject, A.Member):
                parent = visit(subject, context, callee=True)
                check_library_call(subject, [], 'with')
            elif isinstance(subject, A.Index):
                with_subjects.add(id(subject))
                parent = visit(subject, context)
                check_library_call(subject.target, subject.args, 'with')
            else:
                parent = visit(subject, context)
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
