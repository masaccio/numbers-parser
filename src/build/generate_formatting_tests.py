# pragma: noqa: DTZ001

from datetime import datetime
from typing import Any

from numbers_parser import (
    ControlFormattingType,
    Document,
    FractionAccuracy,
    NegativeNumberStyle,
    PaddingType,
    Table,
)


def _get_or_create_table(doc: Document, sheet_name: str, table_name: str = "Corner Cases") -> Table:
    """Return the first table of a new sheet, reusing the default first sheet if untouched."""
    if (
        len(doc.sheets) == 1
        and doc.sheets[0].name == "Sheet 1"
        and doc.sheets[0].tables[0].cell(0, 0).value is None
    ):
        sheet = doc.sheets[0]
        sheet.name = sheet_name
        table = sheet.tables[0]
        table.name = table_name
        return table

    doc.add_sheet(sheet_name, table_name)
    return doc.sheets[sheet_name].tables[table_name]


def _populate_stress_sheet(
    doc: Document,
    sheet_name: str,
    fmt_type: str,
    test_cases: list[tuple[str, Any]],
    subtests: list[tuple[str, dict[str, Any]]],
    ref_generator: callable = None,
) -> None:
    """
    Populate a sheet using the paired-column layout expected by test_fraction_stress:
    Col 0: Test Case Name
    Col 1: Raw Value
    Col 2 + 2*idx: Formatted Cell
    Col 2 + 2*idx + 1: Expected formatted_value string
    """
    table = _get_or_create_table(doc, sheet_name)

    table.write(0, 0, "Test Case")
    table.write(0, 1, "Raw Value")

    for idx, (subtest_name, _) in enumerate(subtests):
        val_col = 2 + (idx * 2)
        str_col = val_col + 1
        table.write(0, val_col, subtest_name)
        table.write(0, str_col, f"{subtest_name} (formatted_value)")

    for row_idx, (test_name, raw_val) in enumerate(test_cases, start=1):
        table.write(row_idx, 0, test_name)
        table.write(row_idx, 1, raw_val)

        canonical_val = table.cell(row_idx, 1).value

        for idx, (subtest_name, kwargs) in enumerate(subtests):
            val_col = 2 + (idx * 2)
            str_col = val_col + 1

            table.write(row_idx, val_col, canonical_val)
            table.set_cell_formatting(row_idx, val_col, fmt_type, **kwargs)

            if ref_generator is not None:
                formatted = ref_generator(subtest_name, canonical_val)
            else:
                formatted = table.cell(row_idx, val_col).formatted_value

            table.write(row_idx, str_col, formatted)


# ---------------------------------------------------------------------------
# 1. "fraction"
# ---------------------------------------------------------------------------
def add_fraction_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Zero", 0.0),
        ("Tiny positive (rounds to 0)", 1e-9),
        ("Tiny negative (rounds to 0, no '-' sign)", -1e-9),
        ("Just below 1/200 (rounds to 0 in halves..hundredths)", -0.004),
        ("Exact positive whole", 1.0),
        ("Exact negative whole", -1.0),
        ("Larger whole", 42.0),
        ("Larger negative whole", -42.0),
        ("Positive overflow to 1 (0 < val < 1)", 0.9999),
        ("Negative overflow to -1 (-1 < val < 0)", -0.9999),
        ("Positive mixed overflow (2.999 -> 3)", 2.9999),
        ("Negative mixed overflow (-2.999 -> -3)", -2.9999),
        ("Overflow boundary for 1/8 (7.5/8 = 0.9375)", 0.94),
        ("Negative overflow boundary for 1/8", -0.94),
        ("Proper half", 0.5),
        ("Negative proper half", -0.5),
        ("Proper quarter", 0.25),
        ("Negative proper three-quarters", -0.75),
        ("Proper eighth", 0.125),
        ("Proper sixteenth (15/16)", 0.9375),
        ("Proper third (12 sig figs)", 0.333333333333),
        ("Negative proper two-thirds (12 sig figs)", -0.666666666667),
        ("Mixed positive (1 1/2)", 1.5),
        ("Mixed negative (-1 1/2)", -1.5),
        ("Mixed repeating (3 1/3, 12 sig figs)", 3.33333333333),
        ("Mixed negative repeating (-3 2/3, 12 sig figs)", -3.66666666667),
        ("Decimal three-tenths (3.3)", 3.3),
        ("Negative decimal three-tenths (-3.3)", -3.3),
        ("High whole with small fraction", 100000.125),
        ("Negative high whole with small fraction", -100000.125),
        ("Tie at 1/4 for Halves", 0.25),
        ("Tie at 3/4 for Halves", 0.75),
        ("Tie at 1.25 for Halves", 1.25),
        ("Tie at 1.75 for Halves", 1.75),
        ("Negative tie at -0.25", -0.25),
        ("Negative tie at -0.75", -0.75),
        ("Pi (12 sig figs)", 3.14159265359),
        ("Negative Pi (12 sig figs)", -3.14159265359),
        ("Euler's number e (12 sig figs)", 2.71828182846),
        ("Golden ratio phi (12 sig figs)", 1.61803398875),
        ("1-digit max denom 8/9 (12 sig figs)", 0.888888888889),
        ("2-digit max denom 98/99 (12 sig figs)", 0.989898989899),
        ("3-digit max denom 998/999 (12 sig figs)", 0.998998998999),
        ("Negative 3-digit max denom -998/999 (12 sig figs)", -0.998998998999),
    ]
    subtests = [(acc.name, {"fraction_accuracy": acc}) for acc in FractionAccuracy]
    _populate_stress_sheet(doc, "fraction", "fraction", test_cases, subtests)


# ---------------------------------------------------------------------------
# 2. "base"
# ---------------------------------------------------------------------------
def add_base_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Zero", 0),
        ("One", 1),
        ("Negative one (two's complement vs minus sign)", -1),
        ("Max 8-bit unsigned (255)", 255),
        ("Negative 255", -255),
        ("Power of 2 boundary (1024)", 1024),
        ("Float rounding up (15.6 -> 16)", 15.6),
        ("Negative float (-15.6 -> -16)", -15.6),
        ("Large integer (65535)", 65535),
    ]
    subtests = [
        ("base_2_0_places", {"base": 2, "base_places": 0, "base_use_minus_sign": True}),
        ("base_2_8_places_twos_comp", {"base": 2, "base_places": 8, "base_use_minus_sign": False}),
        ("base_8_default", {"base": 8, "base_places": 4, "base_use_minus_sign": True}),
        ("base_10_default", {"base": 10, "base_places": 0, "base_use_minus_sign": True}),
        ("base_16_minus", {"base": 16, "base_places": 0, "base_use_minus_sign": True}),
        (
            "base_16_4_places_twos_comp",
            {"base": 16, "base_places": 4, "base_use_minus_sign": False},
        ),
        ("base_36_alpha", {"base": 36, "base_places": 0, "base_use_minus_sign": True}),
    ]
    _populate_stress_sheet(doc, "base", "base", test_cases, subtests)


# ---------------------------------------------------------------------------
# 3. "currency"
# ---------------------------------------------------------------------------
def add_currency_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Zero", 0.0),
        ("Sub-cent positive (rounds to 0.00)", 0.004),
        ("Sub-cent negative (must not be -0.00)", -0.004),
        ("Half-cent tie (0.005)", 0.005),
        ("Negative half-cent tie (-0.005)", -0.005),
        ("Standard positive", 1234.56),
        ("Standard negative", -1234.56),
        ("Millions with fractions (10 sig figs)", 9876543.215),
        ("Negative millions (10 sig figs)", -9876543.215),
    ]
    subtests = [
        (
            "GBP_2dec_minus",
            {
                "currency_code": "GBP",
                "decimal_places": 2,
                "negative_style": NegativeNumberStyle.MINUS,
            },
        ),
        (
            "GBP_2dec_sep_parens",
            {
                "currency_code": "GBP",
                "decimal_places": 2,
                "show_thousands_separator": True,
                "negative_style": NegativeNumberStyle.PARENTHESES,
            },
        ),
        (
            "USD_accounting",
            {
                "currency_code": "USD",
                "decimal_places": 2,
                "show_thousands_separator": True,
                "use_accounting_style": True,
            },
        ),
        (
            "EUR_0dec_red",
            {
                "currency_code": "EUR",
                "decimal_places": 0,
                "show_thousands_separator": True,
                "negative_style": NegativeNumberStyle.RED,
            },
        ),
        (
            "JPY_4dec_red_parens",
            {
                "currency_code": "JPY",
                "decimal_places": 4,
                "negative_style": NegativeNumberStyle.RED_AND_PARENTHESES,
            },
        ),
    ]
    _populate_stress_sheet(doc, "currency", "currency", test_cases, subtests)


# ---------------------------------------------------------------------------
# 4. "datetime" (Using original ICU format strings + Python datetime reference formatting)
# ---------------------------------------------------------------------------
def add_datetime_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Unix epoch", datetime(1970, 1, 1, 0, 0, 0)),
        ("Apple Numbers 2001 epoch", datetime(2001, 1, 1, 0, 0, 0)),
        ("Leap day 2024", datetime(2024, 2, 29, 12, 30, 45)),
        ("End of year midnight", datetime(2026, 12, 31, 23, 59, 59)),
        ("Single digit day and month", datetime(2026, 3, 5, 8, 4, 9)),
    ]
    subtests = [
        ("iso_date", {"date_time_format": "yyyy-MM-dd"}),
        ("uk_long_date", {"date_time_format": "EEEE, d MMMM yyyy"}),
        ("short_date_time", {"date_time_format": "d MMM y HH:mm:ss"}),
        ("12hr_am_pm", {"date_time_format": "h:mm a"}),
        ("year_month_only", {"date_time_format": "MMMM yyyy"}),
    ]

    def datetime_ref_generator(subtest_name: str, dt: datetime) -> str:
        if subtest_name == "iso_date":
            return dt.strftime("%Y-%m-%d")
        if subtest_name == "uk_long_date":
            # EEEE, d MMMM yyyy (no leading zero on day)
            day_str = str(dt.day)
            return dt.strftime(f"%A, {day_str} %B %Y")
        if subtest_name == "short_date_time":
            # d MMM y HH:mm:ss
            day_str = str(dt.day)
            return dt.strftime(f"{day_str} %b %Y %H:%M:%S")
        if subtest_name == "12hr_am_pm":
            # h:mm a (12-hour without leading zero)
            hour = dt.strftime("%I").lstrip("0") or "12"
            return f"{hour}:{dt.strftime('%M %p')}".lower()
        if subtest_name == "year_month_only":
            return dt.strftime("%B %Y")
        return str(dt)

    _populate_stress_sheet(
        doc, "datetime", "datetime", test_cases, subtests, ref_generator=datetime_ref_generator
    )


# ---------------------------------------------------------------------------
# 5. "number"
# ---------------------------------------------------------------------------
def add_number_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Zero", 0.0),
        ("Tiny negative rounding to 0", -0.0001),
        ("Rounding up across whole boundary", 999.999),
        ("Negative rounding across boundary", -999.999),
        ("Banker's rounding tie (2.5)", 2.5),
        ("Banker's rounding tie (3.5)", 3.5),
        ("Large number (13 sig figs)", 12345678.90123),
        ("Large negative number (13 sig figs)", -12345678.90123),
    ]
    subtests = [
        ("default_2dec", {"decimal_places": 2}),
        ("0dec_thousands", {"decimal_places": 0, "show_thousands_separator": True}),
        (
            "4dec_parens",
            {
                "decimal_places": 4,
                "show_thousands_separator": True,
                "negative_style": NegativeNumberStyle.PARENTHESES,
            },
        ),
        ("2dec_red", {"decimal_places": 2, "negative_style": NegativeNumberStyle.RED}),
        ("auto_decimals", {"decimal_places": None, "show_thousands_separator": True}),
    ]
    _populate_stress_sheet(doc, "number", "number", test_cases, subtests)


# ---------------------------------------------------------------------------
# 6. "percentage"
# ---------------------------------------------------------------------------
def add_percentage_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Zero", 0.0),
        ("Tiny negative (rounds to 0%)", -0.00001),
        ("One hundred percent (1.0)", 1.0),
        ("Negative one hundred percent (-1.0)", -1.0),
        ("Fractional percent (12.345%)", 0.12345),
        ("Negative fractional percent (-12.345%)", -0.12345),
        ("Large multiplier (12345.67%)", 123.4567),
    ]
    subtests = [
        ("0dec", {"decimal_places": 0}),
        ("2dec_default", {"decimal_places": 2}),
        (
            "2dec_thousands_parens",
            {
                "decimal_places": 2,
                "show_thousands_separator": True,
                "negative_style": NegativeNumberStyle.PARENTHESES,
            },
        ),
        ("4dec_red", {"decimal_places": 4, "negative_style": NegativeNumberStyle.RED}),
    ]
    _populate_stress_sheet(doc, "percentage", "percentage", test_cases, subtests)


# ---------------------------------------------------------------------------
# 7. "scientific"
# ---------------------------------------------------------------------------
def add_scientific_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Zero", 0.0),
        ("Unit (1.0)", 1.0),
        ("Negative unit (-1.0)", -1.0),
        ("Avogadro scale (9 sig figs)", 6.02214076e23),
        ("Planck scale small (9 sig figs)", 6.62607015e-34),
        ("Negative small (6 sig figs)", -1.23456e-12),
        ("Mantissa rounding overflow (9.999 -> 1.00E+01)", 9.999),
        ("Negative mantissa overflow (-9.999)", -9.999),
    ]
    subtests = [
        ("0dec", {"decimal_places": 0}),
        ("2dec", {"decimal_places": 2}),
        ("4dec", {"decimal_places": 4}),
        ("6dec", {"decimal_places": 6}),
    ]
    _populate_stress_sheet(doc, "scientific", "scientific", test_cases, subtests)


# ---------------------------------------------------------------------------
# 8. "tickbox"
# ---------------------------------------------------------------------------
def add_tickbox_stress_tests(doc: Document) -> None:
    test_cases = [
        ("False boolean", False),
        ("True boolean", True),
    ]
    subtests = [
        ("default_tickbox", {}),
    ]
    _populate_stress_sheet(doc, "tickbox", "tickbox", test_cases, subtests)


# ---------------------------------------------------------------------------
# 9. "rating"
# ---------------------------------------------------------------------------
def add_rating_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Zero stars", 0),
        ("One star", 1),
        ("Three stars", 3),
        ("Five stars (max)", 5),
    ]
    subtests = [
        ("default_rating", {}),
    ]
    _populate_stress_sheet(doc, "rating", "rating", test_cases, subtests)


# ---------------------------------------------------------------------------
# 10. "slider"
# ---------------------------------------------------------------------------
def add_slider_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Minimum bound (0.0)", 0.0),
        ("Midpoint fraction (5.5)", 5.5),
        ("Maximum bound (100.0)", 100.0),
    ]
    subtests = [
        (
            "number_slider",
            {
                "control_format": ControlFormattingType.NUMBER,
                "minimum": 0.0,
                "maximum": 100.0,
                "increment": 0.5,
                "decimal_places": 1,
            },
        ),
        (
            "currency_slider",
            {
                "control_format": ControlFormattingType.CURRENCY,
                "minimum": 0.0,
                "maximum": 100.0,
                "increment": 1.0,
                "currency_code": "GBP",
                "decimal_places": 2,
            },
        ),
        (
            "percent_slider",
            {
                "control_format": ControlFormattingType.PERCENTAGE,
                "minimum": 0.0,
                "maximum": 100.0,
                "increment": 0.1,
                "decimal_places": 1,
            },
        ),
        (
            "fraction_slider",
            {
                "control_format": ControlFormattingType.FRACTION,
                "minimum": 0.0,
                "maximum": 100.0,
                "increment": 0.25,
                "fraction_accuracy": FractionAccuracy.QUARTERS,
            },
        ),
        (
            "scientific_slider",
            {
                "control_format": ControlFormattingType.SCIENTIFIC,
                "minimum": 0.0,
                "maximum": 100.0,
                "increment": 1.0,
                "decimal_places": 2,
            },
        ),
        (
            "base_slider",
            {
                "control_format": ControlFormattingType.BASE,
                "minimum": 0.0,
                "maximum": 100.0,
                "increment": 1.0,
                "base": 16,
                "base_places": 2,
            },
        ),
    ]
    _populate_stress_sheet(doc, "slider", "slider", test_cases, subtests)


# ---------------------------------------------------------------------------
# 11. "stepper"
# ---------------------------------------------------------------------------
def add_stepper_stress_tests(doc: Document) -> None:
    test_cases = [
        ("Negative value (-10.0)", -10.0),
        ("Zero (0.0)", 0.0),
        ("Fractional step (12.75)", 12.75),
        ("Max bound (50.0)", 50.0),
    ]
    subtests = [
        (
            "number_stepper",
            {
                "control_format": ControlFormattingType.NUMBER,
                "minimum": -10.0,
                "maximum": 50.0,
                "increment": 0.25,
                "decimal_places": 2,
            },
        ),
        (
            "currency_stepper",
            {
                "control_format": ControlFormattingType.CURRENCY,
                "minimum": -10.0,
                "maximum": 50.0,
                "increment": 1.0,
                "currency_code": "GBP",
                "decimal_places": 2,
            },
        ),
        (
            "percent_stepper",
            {
                "control_format": ControlFormattingType.PERCENTAGE,
                "minimum": -10.0,
                "maximum": 50.0,
                "increment": 0.05,
                "decimal_places": 1,
            },
        ),
        (
            "fraction_stepper",
            {
                "control_format": ControlFormattingType.FRACTION,
                "minimum": -10.0,
                "maximum": 50.0,
                "increment": 0.25,
                "fraction_accuracy": FractionAccuracy.EIGHTHS,
            },
        ),
        (
            "scientific_stepper",
            {
                "control_format": ControlFormattingType.SCIENTIFIC,
                "minimum": -10.0,
                "maximum": 50.0,
                "increment": 1.0,
                "decimal_places": 2,
            },
        ),
        (
            "base_stepper",
            {
                "control_format": ControlFormattingType.BASE,
                "minimum": -10.0,
                "maximum": 50.0,
                "increment": 1.0,
                "base": 16,
                "base_places": 2,
            },
        ),
    ]
    _populate_stress_sheet(doc, "stepper", "stepper", test_cases, subtests)


# ---------------------------------------------------------------------------
# 12. "popup"
# ---------------------------------------------------------------------------
def add_popup_stress_tests(doc: Document) -> None:
    test_cases = [
        ("First item", "Alpha"),
        ("Middle item", "Beta"),
        ("Numeric string item", "123"),
    ]
    popup_items = ["Alpha", "Beta", "123"]
    subtests = [
        ("allow_none_true", {"popup_values": popup_items, "allow_none": True}),
        ("allow_none_false", {"popup_values": popup_items, "allow_none": False}),
    ]
    _populate_stress_sheet(doc, "popup", "popup", test_cases, subtests)


# ---------------------------------------------------------------------------
# 13. "custom"
# ---------------------------------------------------------------------------
def add_custom_stress_tests(doc: Document) -> None:
    num_pad_zeros = doc.add_custom_format(
        name="NumPadZeros",
        type="number",
        integer_format=PaddingType.ZEROS,
        decimal_format=PaddingType.ZEROS,
        num_integers=6,
        num_decimals=3,
        show_thousands_separator=True,
    )
    num_pad_spaces = doc.add_custom_format(
        name="NumPadSpaces",
        type="number",
        integer_format=PaddingType.SPACES,
        decimal_format=PaddingType.SPACES,
        num_integers=6,
        num_decimals=2,
        show_thousands_separator=False,
    )
    num_cases = [
        ("Zero", 0.0),
        ("Small float", 4.5),
        ("Negative float", -42.125),
        ("Large float (9 sig figs)", 12345.6789),
    ]
    num_subtests = [
        ("custom_num_zeros", {"format": num_pad_zeros}),
        ("custom_num_spaces", {"format": num_pad_spaces}),
    ]
    _populate_stress_sheet(doc, "custom_number", "custom", num_cases, num_subtests)

    dt_custom = doc.add_custom_format(
        name="CustomISO",
        type="datetime",
        format="yyyy/MM/dd HH:mm",
    )
    dt_cases = [
        ("Leap day", datetime(2024, 2, 29, 9, 5)),
        ("Year end", datetime(2026, 12, 31, 23, 45)),
    ]
    dt_subtests = [
        ("custom_datetime_iso", {"format": dt_custom}),
    ]

    def custom_dt_ref(subtest_name: str, dt: datetime) -> str:
        return dt.strftime("%Y/%m/%d %H:%M")

    _populate_stress_sheet(
        doc, "custom_datetime", "custom", dt_cases, dt_subtests, ref_generator=custom_dt_ref
    )

    txt_custom = doc.add_custom_format(
        name="CustomTextWrapper",
        type="text",
        format="Prefix [%s] Suffix",
    )
    txt_cases = [
        ("Simple word", "Token"),
        ("Symbols", "A&B"),
    ]
    txt_subtests = [
        ("custom_text_wrap", {"format": txt_custom}),
    ]

    def custom_txt_ref(subtest_name: str, val: str) -> str:
        return f"Prefix [{val}] Suffix"

    _populate_stress_sheet(
        doc, "custom_text", "custom", txt_cases, txt_subtests, ref_generator=custom_txt_ref
    )


# ---------------------------------------------------------------------------
# Master Document Builder
# ---------------------------------------------------------------------------
ALL_STRESS_GENERATORS = [
    add_fraction_stress_tests,
    add_base_stress_tests,
    add_currency_stress_tests,
    add_datetime_stress_tests,
    add_number_stress_tests,
    add_percentage_stress_tests,
    add_scientific_stress_tests,
    add_tickbox_stress_tests,
    add_rating_stress_tests,
    add_slider_stress_tests,
    add_stepper_stress_tests,
    add_popup_stress_tests,
    add_custom_stress_tests,
]


def build_all_formatting_stress_doc(
    filename: str = "tests/data/standard-format-stress.numbers",
) -> None:
    doc = Document()
    for generator_fn in ALL_STRESS_GENERATORS:
        generator_fn(doc)
    doc.save(filename)
    print(f"Saved {filename}")


if __name__ == "__main__":
    build_all_formatting_stress_doc()
