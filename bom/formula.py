"""Safe formula evaluator for BOM quantity/wastage expressions.

Templates store consumption as formulas (e.g. ``piece_area_sqm * gsm / 1000``)
rather than fixed numbers, because the carton reference data showed
consumption is computed from SKU geometry, not a stored constant. This module
evaluates such expressions WITHOUT ``eval()``/``exec()``: it parses the
expression into an AST and walks it, rejecting any node type or name that
isn't on the allow-list below.

Never widen this allow-list to include Attribute, Subscript, Lambda,
comprehensions, Import, or arbitrary Call targets — those are exactly the
constructs that turn a "safe expression language" into arbitrary code
execution.
"""
from __future__ import annotations

import ast
import math
from decimal import Decimal, DivisionByZero, InvalidOperation

MAX_EXPRESSION_LENGTH = 500
MAX_EXPONENT = 12

ALLOWED_BINOPS = (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.FloorDiv, ast.Mod, ast.Pow)
ALLOWED_UNARYOPS = (ast.UAdd, ast.USub, ast.Not)
ALLOWED_BOOLOPS = (ast.And, ast.Or)
ALLOWED_CMPOPS = (ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE)

ALLOWED_NODES = (
    ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Name, ast.Load,
    ast.Compare, ast.BoolOp, ast.IfExp, ast.Call,
) + ALLOWED_BINOPS + ALLOWED_UNARYOPS + ALLOWED_BOOLOPS + ALLOWED_CMPOPS


class FormulaError(Exception):
    """Raised for an invalid or unsafe formula expression, or a runtime evaluation failure."""


def _to_decimal(value):
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return Decimal(1) if value else Decimal(0)
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    if value is None:
        return Decimal(0)
    raise FormulaError(f'Cannot use non-numeric value in a formula: {value!r}')


def _safe_sqrt(x):
    d = _to_decimal(x)
    if d < 0:
        raise FormulaError('sqrt() of a negative number')
    return Decimal(str(math.sqrt(float(d))))


def _safe_round(x, ndigits=0):
    return Decimal(str(round(float(_to_decimal(x)), int(ndigits))))


ALLOWED_FUNCTIONS = {
    'min': lambda *a: min(_to_decimal(x) for x in a),
    'max': lambda *a: max(_to_decimal(x) for x in a),
    'abs': lambda x: abs(_to_decimal(x)),
    'round': _safe_round,
    'ceil': lambda x: Decimal(math.ceil(_to_decimal(x))),
    'floor': lambda x: Decimal(math.floor(_to_decimal(x))),
    'sqrt': _safe_sqrt,
}


class _SafeEvaluator(ast.NodeVisitor):
    def __init__(self, namespace):
        self.namespace = namespace

    def visit(self, node):
        if not isinstance(node, ALLOWED_NODES):
            raise FormulaError(f'Disallowed expression element: {type(node).__name__}')
        method = 'eval_' + type(node).__name__
        visitor = getattr(self, method, None)
        if visitor is None:
            raise FormulaError(f'Unsupported expression element: {type(node).__name__}')
        return visitor(node)

    def eval_Expression(self, node):
        return self.visit(node.body)

    def eval_Constant(self, node):
        if isinstance(node.value, bool):
            return node.value
        if isinstance(node.value, (int, float)):
            return _to_decimal(node.value)
        raise FormulaError(f'Disallowed constant: {node.value!r}')

    def eval_Name(self, node):
        if node.id not in self.namespace:
            raise FormulaError(f'Unknown variable: {node.id}')
        return self.namespace[node.id]

    def eval_BinOp(self, node):
        if not isinstance(node.op, ALLOWED_BINOPS):
            raise FormulaError(f'Disallowed operator: {type(node.op).__name__}')
        left = _to_decimal(self.visit(node.left))
        right = _to_decimal(self.visit(node.right))
        try:
            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            if isinstance(node.op, ast.Mult):
                return left * right
            if isinstance(node.op, ast.Div):
                return left / right
            if isinstance(node.op, ast.FloorDiv):
                return left // right
            if isinstance(node.op, ast.Mod):
                return left % right
            if isinstance(node.op, ast.Pow):
                if abs(right) > MAX_EXPONENT:
                    raise FormulaError(f'Exponent too large: {right}')
                return left ** right
        except (DivisionByZero, ZeroDivisionError, InvalidOperation) as exc:
            raise FormulaError(f'Division by zero evaluating: {ast.dump(node)}') from exc
        raise FormulaError('Unreachable operator branch')

    def eval_UnaryOp(self, node):
        if not isinstance(node.op, ALLOWED_UNARYOPS):
            raise FormulaError(f'Disallowed unary operator: {type(node.op).__name__}')
        value = self.visit(node.operand)
        if isinstance(node.op, ast.UAdd):
            return _to_decimal(value)
        if isinstance(node.op, ast.USub):
            return -_to_decimal(value)
        if isinstance(node.op, ast.Not):
            return not bool(value)
        raise FormulaError('Unreachable unary branch')

    def eval_BoolOp(self, node):
        if not isinstance(node.op, ALLOWED_BOOLOPS):
            raise FormulaError('Disallowed boolean operator')
        values = [self.visit(v) for v in node.values]
        if isinstance(node.op, ast.And):
            result = values[0]
            for v in values[1:]:
                result = result and v
            return result
        result = values[0]
        for v in values[1:]:
            result = result or v
        return result

    def eval_Compare(self, node):
        left = self.visit(node.left)
        for op, comparator in zip(node.ops, node.comparators):
            if not isinstance(op, ALLOWED_CMPOPS):
                raise FormulaError(f'Disallowed comparison operator: {type(op).__name__}')
            right = self.visit(comparator)
            left_d, right_d = _to_decimal(left), _to_decimal(right)
            if isinstance(op, ast.Eq):
                ok = left_d == right_d
            elif isinstance(op, ast.NotEq):
                ok = left_d != right_d
            elif isinstance(op, ast.Lt):
                ok = left_d < right_d
            elif isinstance(op, ast.LtE):
                ok = left_d <= right_d
            elif isinstance(op, ast.Gt):
                ok = left_d > right_d
            elif isinstance(op, ast.GtE):
                ok = left_d >= right_d
            else:
                raise FormulaError('Unreachable comparison branch')
            if not ok:
                return False
            left = right
        return True

    def eval_IfExp(self, node):
        return self.visit(node.body) if self.visit(node.test) else self.visit(node.orelse)

    def eval_Call(self, node):
        if node.keywords:
            raise FormulaError('Keyword arguments are not allowed in formulas')
        if not isinstance(node.func, ast.Name):
            raise FormulaError('Only calls to a fixed set of named functions are allowed')
        func = ALLOWED_FUNCTIONS.get(node.func.id)
        if func is None:
            raise FormulaError(f'Unknown function: {node.func.id}')
        args = [self.visit(a) for a in node.args]
        try:
            return func(*args)
        except FormulaError:
            raise
        except Exception as exc:
            raise FormulaError(f'Error calling {node.func.id}(): {exc}') from exc


def _parse(expression):
    if not expression or not expression.strip():
        raise FormulaError('Empty formula')
    if len(expression) > MAX_EXPRESSION_LENGTH:
        raise FormulaError('Formula too long')
    try:
        tree = ast.parse(expression, mode='eval')
    except SyntaxError as exc:
        raise FormulaError(f'Invalid syntax: {exc}') from exc
    return tree


def validate_expression(expression):
    """Parse and structurally validate a formula without a namespace.

    Raises FormulaError if the expression cannot possibly be safe (bad syntax,
    disallowed node types) — used by forms to reject a bad formula at save
    time. Unknown-variable errors are not raised here since the namespace may
    depend on run-time context; use evaluate() with a real namespace for that.
    """
    tree = _parse(expression)
    evaluator = _SafeEvaluator(namespace={})

    class _StructureOnly(ast.NodeVisitor):
        def generic_visit(self, node):
            if not isinstance(node, ALLOWED_NODES):
                raise FormulaError(f'Disallowed expression element: {type(node).__name__}')
            if isinstance(node, ast.Call):
                if node.keywords:
                    raise FormulaError('Keyword arguments are not allowed in formulas')
                if not isinstance(node.func, ast.Name) or node.func.id not in ALLOWED_FUNCTIONS:
                    func_name = getattr(node.func, 'id', type(node.func).__name__)
                    raise FormulaError(f'Unknown function: {func_name}')
            super().generic_visit(node)

    _StructureOnly().visit(tree)
    return tree


def evaluate(expression, namespace):
    """Evaluate a formula expression against a namespace of Decimal/numeric values.

    Returns a Decimal. Raises FormulaError on anything unsafe or invalid,
    including references to names not present in ``namespace``.
    """
    tree = _parse(expression)
    safe_namespace = {k: (v if isinstance(v, Decimal) else _to_decimal(v) if isinstance(v, (int, float, bool)) or v is None else v)
                      for k, v in namespace.items()}
    evaluator = _SafeEvaluator(safe_namespace)
    result = evaluator.visit(tree)
    if isinstance(result, bool):
        return Decimal(1) if result else Decimal(0)
    return _to_decimal(result)
