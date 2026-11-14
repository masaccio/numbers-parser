"""Regression cases for scientific literals in decoded formulas."""

from decimal import Decimal

import pytest

from numbers_parser.formula import number_to_str


@pytest.mark.parametrize(
    "value",
    [1.23456789e-7, 1.23e-9, 1.23e16, 123.4, 0.0],
)
def test_formula_number_preserves_observed_float(value: float) -> None:
    """Formatting scientific notation must preserve its decimal exponent."""
    assert Decimal(number_to_str(value)) == Decimal(repr(value))
