from decimal import Decimal

from django.test import SimpleTestCase

from .formula import FormulaError, evaluate, validate_expression


class FormulaArithmeticTests(SimpleTestCase):
    def test_basic_arithmetic(self):
        self.assertEqual(evaluate('2 + 3 * 4', {}), Decimal('14'))
        self.assertEqual(evaluate('(2 + 3) * 4', {}), Decimal('20'))
        self.assertEqual(evaluate('10 / 4', {}), Decimal('2.5'))
        self.assertEqual(evaluate('10 // 4', {}), Decimal('2'))
        self.assertEqual(evaluate('10 % 3', {}), Decimal('1'))
        self.assertEqual(evaluate('-5', {}), Decimal('-5'))
        self.assertEqual(evaluate('2 ** 3', {}), Decimal('8'))

    def test_variables(self):
        self.assertEqual(evaluate('a * b', {'a': Decimal('2.5'), 'b': Decimal('4')}), Decimal('10.0'))

    def test_functions(self):
        self.assertEqual(evaluate('min(3, 1, 2)', {}), Decimal('1'))
        self.assertEqual(evaluate('max(3, 1, 2)', {}), Decimal('3'))
        self.assertEqual(evaluate('abs(-5)', {}), Decimal('5'))
        self.assertEqual(evaluate('round(3.14159, 2)', {}), Decimal('3.14'))
        self.assertEqual(evaluate('ceil(3.1)', {}), Decimal('4'))
        self.assertEqual(evaluate('floor(3.9)', {}), Decimal('3'))
        self.assertEqual(evaluate('sqrt(9)', {}), Decimal('3'))

    def test_conditional(self):
        self.assertEqual(evaluate('10 if 1 > 0 else 20', {}), Decimal('10'))
        self.assertEqual(evaluate('10 if 1 < 0 else 20', {}), Decimal('20'))

    def test_boolean_ops(self):
        self.assertEqual(evaluate('5 if (1 > 0 and 2 > 1) else 0', {}), Decimal('5'))
        self.assertEqual(evaluate('5 if (1 > 0 or 2 < 1) else 0', {}), Decimal('5'))

    def test_carton_reference_ratio(self):
        """Reproduces the verified carton reference ratio: starch = 0.0528 * flute_kg."""
        flute_kg = Decimal('0.882925')
        starch = evaluate('flute_kg * 0.0528', {'flute_kg': flute_kg})
        self.assertAlmostEqual(float(starch), float(flute_kg) * 0.0528, places=6)

    def test_division_by_zero_raises_formula_error(self):
        with self.assertRaises(FormulaError):
            evaluate('1 / 0', {})

    def test_unknown_variable_raises(self):
        with self.assertRaises(FormulaError):
            evaluate('unknown_var * 2', {})

    def test_empty_expression_raises(self):
        with self.assertRaises(FormulaError):
            evaluate('', {})

    def test_too_long_expression_raises(self):
        with self.assertRaises(FormulaError):
            evaluate('1+' * 400 + '1', {})

    def test_large_exponent_rejected(self):
        with self.assertRaises(FormulaError):
            evaluate('2 ** 999', {})


class FormulaSecurityTests(SimpleTestCase):
    """Every one of these must be rejected — this is the entire point of the module."""

    def test_rejects_import(self):
        with self.assertRaises(FormulaError):
            evaluate('__import__("os").system("echo hi")', {})

    def test_rejects_attribute_access(self):
        with self.assertRaises(FormulaError):
            evaluate('(1).__class__', {})

    def test_rejects_dunder_name_via_namespace(self):
        with self.assertRaises(FormulaError):
            evaluate('x.__class__.__bases__', {'x': Decimal(1)})

    def test_rejects_subscript(self):
        with self.assertRaises(FormulaError):
            evaluate('a[0]', {'a': Decimal(1)})

    def test_rejects_lambda(self):
        with self.assertRaises(FormulaError):
            evaluate('(lambda: 1)()', {})

    def test_rejects_list_comprehension(self):
        with self.assertRaises(FormulaError):
            evaluate('[x for x in range(10)]', {})

    def test_rejects_unknown_function_call(self):
        with self.assertRaises(FormulaError):
            evaluate('eval("1")', {})

    def test_rejects_exec_style_string(self):
        with self.assertRaises(FormulaError):
            evaluate('exec("import os")', {})

    def test_rejects_string_constant(self):
        with self.assertRaises(FormulaError):
            evaluate('"a" + "b"', {})

    def test_rejects_kwargs_in_call(self):
        with self.assertRaises(FormulaError):
            evaluate('round(1.5, ndigits=1)', {})

    def test_rejects_assignment_syntax(self):
        with self.assertRaises(FormulaError):
            evaluate('x = 1', {})

    def test_rejects_multiple_statements(self):
        with self.assertRaises(FormulaError):
            evaluate('1; 2', {})


class ValidateExpressionTests(SimpleTestCase):
    def test_validates_good_expression_without_namespace(self):
        validate_expression('piece_area_sqm * gsm / 1000')  # should not raise

    def test_rejects_bad_syntax(self):
        with self.assertRaises(FormulaError):
            validate_expression('a * / b')

    def test_rejects_disallowed_construct_without_namespace(self):
        with self.assertRaises(FormulaError):
            validate_expression('__import__("os")')

    def test_does_not_require_namespace_for_names(self):
        # Unknown *names* are fine at validate time (namespace-dependent);
        # only structurally unsafe constructs are rejected here.
        validate_expression('some_future_variable * 2')
