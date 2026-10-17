import math

from numbers_parser import Document, FormattingType, FractionAccuracy

# Corner cases targeting _format_fraction_parts, _float_to_fraction,
# and _float_to_n_digit_fraction
TEST_CASES = [
    # 1. Zeroes & Signed Zeroes (tests `numerator == 0`, `whole == 0`, and `result != "0"`)
    ("Positive zero", 0.0),
    ("Negative IEEE-754 zero", -0.0),
    # 2. Sub-threshold values that round down to "0" (must NOT produce "-0")
    ("Tiny positive (rounds to 0)", 1e-9),
    ("Tiny negative (rounds to 0, no '-' sign)", -1e-9),
    ("Just below 1/200 (rounds to 0 in halves..hundredths)", -0.004),
    # 3. Exact Whole Numbers (tests `whole != 0` and `numerator == 0`)
    ("Exact positive whole", 1.0),
    ("Exact negative whole", -1.0),
    ("Larger whole", 42.0),
    ("Larger negative whole", -42.0),
    # 4. Rounding Overflow (tests `numerator >= denominator` -> `whole += 1, numerator = 0`)
    ("Positive overflow to 1 (0 < val < 1)", 0.9999),
    ("Negative overflow to -1 (-1 < val < 0)", -0.9999),
    ("Positive mixed overflow (2.999 -> 3)", 2.9999),
    ("Negative mixed overflow (-2.999 -> -3)", -2.9999),
    ("Overflow boundary for 1/8 (7.5/8 = 0.9375)", 0.94),
    ("Negative overflow boundary for 1/8", -0.94),
    # 5. Proper Fractions (`whole == 0`, `numerator != 0`)
    ("Proper half", 0.5),
    ("Negative proper half", -0.5),
    ("Proper quarter", 0.25),
    ("Negative proper three-quarters", -0.75),
    ("Proper eighth", 0.125),
    ("Proper sixteenth (15/16)", 0.9375),
    ("Proper third (repeating binary)", 1.0 / 3.0),
    ("Negative proper two-thirds", -2.0 / 3.0),
    # 6. Mixed Fractions (`whole != 0`, `numerator != 0`)
    ("Mixed positive (1 1/2)", 1.5),
    ("Mixed negative (-1 1/2)", -1.5),
    ("Mixed repeating (3 1/3)", 3.0 + 1.0 / 3.0),
    ("Mixed negative repeating (-3 2/3)", -(3.0 + 2.0 / 3.0)),
    # 7. IEEE-754 Float Subtraction Drift (`val - whole` precision traps)
    ("Float representation trap 1.1 + 2.2", 1.1 + 2.2),
    ("Negative float trap -(1.1 + 2.2)", -(1.1 + 2.2)),
    ("High whole with small fraction (precision loss test)", 100000.125),
    ("Negative high whole with small fraction", -100000.125),
    # 8. Halfway / Banker's Rounding Boundaries (`round(x.5)` even/odd ties)
    ("Tie at 1/4 for Halves (0.25 -> even 0)", 0.25),
    ("Tie at 3/4 for Halves (0.75 -> even 1)", 0.75),
    ("Tie at 1.25 for Halves (1.25 -> even 1)", 1.25),
    ("Tie at 1.75 for Halves (1.75 -> even 2)", 1.75),
    ("Negative tie at -0.25", -0.25),
    ("Negative tie at -0.75", -0.75),
    # 9. N-Digit Denominator & Continued Fraction Stress (`limit_denominator(10**n - 1)`)
    ("Pi (classic rational approx: 3 1/7, 3 16/113)", math.pi),
    ("Negative Pi", -math.pi),
    ("Euler's number e", math.e),
    ("Golden ratio phi", (1 + math.sqrt(5)) / 2),
    ("Exact 1-digit max denom (8/9)", 8.0 / 9.0),
    ("Exact 2-digit max denom (98/99)", 98.0 / 99.0),
    ("Exact 3-digit max denom (998/999)", 998.0 / 999.0),
    ("Negative 3-digit max denom (-998/999)", -998.0 / 999.0),
    ("Unreduced fixed vs reduced n-digit (0.5 in 8ths = 4/8 vs 1/2)", 0.5),
]


def build_fraction_stress_doc(filename: str = "fraction_stress_test.numbers") -> None:
    doc = Document()
    sheet = doc.sheets[0]
    sheet.name = "Fraction Stress Tests"
    table = sheet.tables[0]
    table.name = "Corner Cases"

    accuracies = list(FractionAccuracy)

    # Header layout:
    # Col 0: Test Description
    # Col 1: Raw Float
    # For each FractionAccuracy: one formatted live cell + one expected string cell
    table.write(0, 0, "Test Case")
    table.write(0, 1, "Raw Float")

    for idx, acc in enumerate(accuracies):
        val_col = 2 + (idx * 2)
        str_col = val_col + 1
        table.write(0, val_col, f"{acc.name} (Cell)")
        table.write(0, str_col, f"{acc.name} (formatted_value)")

    # Populate rows and apply formatting
    for row_idx, (description, value) in enumerate(TEST_CASES, start=1):
        table.write(row_idx, 0, description)
        table.write(row_idx, 1, value)

        for idx, acc in enumerate(accuracies):
            val_col = 2 + (idx * 2)
            str_col = val_col + 1

            table.write(row_idx, val_col, value)
            table.set_cell_formatting(
                row_idx,
                val_col,
                FormattingType.FRACTION,
                fraction_accuracy=acc,
            )

            # Capture the output of numbers-parser's fraction formatter right now
            formatted = table.cell(row_idx, val_col).formatted_value
            table.write(row_idx, str_col, formatted)

    doc.save(filename)
    print(f"Saved {len(TEST_CASES)} test cases across {len(accuracies)} accuracies to {filename}")


if __name__ == "__main__":
    build_fraction_stress_doc("tests/data/test-extra-formulas.numbers")
