import pytest
import pytest_check as check
from pendulum import datetime

from numbers_parser import Document, EmptyCell
from numbers_parser.constants import (
    FractionAccuracy,
    NegativeNumberStyle,
)

DATE_FORMAT_REF = [
    ["", "1 pm", "1:25 pm", "1:25:42 pm", "13:25", "13:25:42"],
    [
        "Saturday, 1 April 2023",
        "Saturday, 1 April 2023 1 pm",
        "Saturday, 1 April 2023 1:25 pm",
        "Saturday, 1 April 2023 1:25:42 pm",
        "Saturday, 1 April 2023 13:25",
        "Saturday, 1 April 2023 13:25:42",
    ],
    [
        "Sat, 1 Apr 2023",
        "Sat, 1 Apr 2023 1 pm",
        "Sat, 1 Apr 2023 1:25 pm",
        "Sat, 1 Apr 2023 1:25:42 pm",
        "Sat, 1 Apr 2023 13:25",
        "Sat, 1 Apr 2023 13:25:42",
    ],
    [
        "1 April 2023",
        "1 April 2023 1 pm",
        "1 April 2023 1:25 pm",
        "1 April 2023 1:25:42 pm",
        "1 April 2023 13:25",
        "1 April 2023 13:25:42",
    ],
    [
        "1 Apr 2023",
        "1 Apr 2023 1 pm",
        "1 Apr 2023 1:25 pm",
        "1 Apr 2023 1:25:42 pm",
        "1 Apr 2023 13:25",
        "1 Apr 2023 13:25:42",
    ],
    [
        "1 April",
        "1 April 1 pm",
        "1 April 1:25 pm",
        "1 April 1:25:42 pm",
        "1 April 13:25",
        "1 April 13:25:42",
    ],
    ["Apr 1", "Apr 1 1 pm", "Apr 1 1:25 pm", "Apr 1 1:25:42 pm", "Apr 1 13:25", "Apr 1 13:25:42"],
    [
        "April 2023",
        "April 2023 1 pm",
        "April 2023 1:25 pm",
        "April 2023 1:25:42 pm",
        "April 2023 13:25",
        "April 2023 13:25:42",
    ],
    ["April", "April 1 pm", "April 1:25 pm", "April 1:25:42 pm", "April 13:25", "April 13:25:42"],
    ["2023", "2023 1 pm", "2023 1:25 pm", "2023 1:25:42 pm", "2023 13:25", "2023 13:25:42"],
    [
        "1/4/2023",
        "1/4/2023 1 pm",
        "1/4/2023 1:25 pm",
        "1/4/2023 1:25:42 pm",
        "1/4/2023 13:25",
        "1/4/2023 13:25:42",
    ],
    [
        "01/04/2023",
        "01/04/2023 1 pm",
        "01/04/2023 1:25 pm",
        "01/04/2023 1:25:42 pm",
        "01/04/2023 13:25",
        "01/04/2023 13:25:42",
    ],
    [
        "1/4/23",
        "1/4/23 1 pm",
        "1/4/23 1:25 pm",
        "1/4/23 1:25:42 pm",
        "1/4/23 13:25",
        "1/4/23 13:25:42",
    ],
    [
        "01/04/23",
        "01/04/23 1 pm",
        "01/04/23 1:25 pm",
        "01/04/23 1:25:42 pm",
        "01/04/23 13:25",
        "01/04/23 13:25:42",
    ],
    ["1/4", "1/4 1 pm", "1/4 1:25 pm", "1/4 1:25:42 pm", "1/4 13:25", "1/4 13:25:42"],
    [
        "2023-04-01",
        "2023-04-01 1 pm",
        "2023-04-01 1:25 pm",
        "2023-04-01 1:25:42 pm",
        "2023-04-01 13:25",
        "2023-04-01 13:25:42",
    ],
]

NUMBER_FORMAT_REF = [
    [
        "12345.012346",
        "12345.012346",
        "12345.012346",
        "12345.012346",
        "12,345.012346",
        "12,345.012346",
        "12,345.012346",
        "12,345.012346",
    ],
    ["12345", "12345", "12345", "12345", "12,345", "12,345", "12,345", "12,345"],
    ["12345.0", "12345.0", "12345.0", "12345.0", "12,345.0", "12,345.0", "12,345.0", "12,345.0"],
    [
        "12345.012",
        "12345.012",
        "12345.012",
        "12345.012",
        "12,345.012",
        "12,345.012",
        "12,345.012",
        "12,345.012",
    ],
    [
        "12345.01234600",
        "12345.01234600",
        "12345.01234600",
        "12345.01234600",
        "12,345.01234600",
        "12,345.01234600",
        "12,345.01234600",
        "12,345.01234600",
    ],
    [
        "-12345.012346",
        "12345.012346",
        "(12345.012346)",
        "(12345.012346)",
        "-12,345.012346",
        "12,345.012346",
        "(12,345.012346)",
        "(12,345.012346)",
    ],
    ["-12345", "12345", "(12345)", "(12345)", "-12,345", "12,345", "(12,345)", "(12,345)"],
    [
        "-12345.0",
        "12345.0",
        "(12345.0)",
        "(12345.0)",
        "-12,345.0",
        "12,345.0",
        "(12,345.0)",
        "(12,345.0)",
    ],
    [
        "-12345.012",
        "12345.012",
        "(12345.012)",
        "(12345.012)",
        "-12,345.012",
        "12,345.012",
        "(12,345.012)",
        "(12,345.012)",
    ],
    [
        "-12345.01234600",
        "12345.01234600",
        "(12345.01234600)",
        "(12345.01234600)",
        "-12,345.01234600",
        "12,345.01234600",
        "(12,345.01234600)",
        "(12,345.01234600)",
    ],
]

CURRENCY_FORMAT_REF = [
    [
        "A$12345.01",
        "€12345.01",
        "£12345.01",
        "₹12345.01",
        "₩12345.01",
        "JP¥12345.01",
        "CN¥12345.01",
        "₪12345.01",
    ],
    [
        "₫12345.01",
        "CA$12345.01",
        "SEK 12345.01",
        "£\t12345.01",
        "£\t(12345.01)",
        "£12345.01",
        "£\t12345.01",
        "£\t(12,345.01)",
    ],
]

DATE_FORMATS = [
    "",
    "EEEE, d MMMM yyyy",
    "EEE, d MMM yyyy",
    "d MMMM yyyy",
    "d MMM yyyy",
    "d MMMM",
    "MMM d",
    "MMMM yyyy",
    "MMMM",
    "yyyy",
    "d/M/yyyy",
    "dd/MM/y",
    "d/M/yy",
    "dd/MM/yy",
    "d/M",
    "yyyy-MM-dd",
]

TIME_FORMATS = ["", "h a", "h:mm a", "h:mm:ss a", "HH:mm", "HH:mm:ss"]

OTHER_FORMAT_REF = [
    ["12.34%", "-12.34%", "12.34%", "12.34%", "12.34%", "(12.34%)", "12.34%", "(12.34%)", ""],
    [
        "1,234.56%",
        "-1,234.56%",
        "1,234.56%",
        "1,234.56%",
        "1,234.56%",
        "(1,234.56%)",
        "1,234.56%",
        "(1,234.56%)",
        "",
    ],
    ["12%", "-12%", "12%", "12%", "12%", "(12%)", "12%", "(12%)"],
    ["1,235%", "-1,235%", "1,235%", "1,235%", "1,235%", "(1,235%)", "1,235%", "(1,235%)"],
    ["12.3%", "-12.3%", "12.3%", "12.3%", "12.3%", "(12.3%)", "12.3%", "(12.3%)"],
    [
        "1,234.6%",
        "-1,234.6%",
        "1,234.6%",
        "1,234.6%",
        "1,234.6%",
        "(1,234.6%)",
        "1,234.6%",
        "(1,234.6%)",
    ],
    [
        "12.3400%",
        "-12.3400%",
        "12.3400%",
        "12.3400%",
        "12.3400%",
        "(12.3400%)",
        "12.3400%",
        "(12.3400%)",
    ],
    [
        "1,234.5600%",
        "-1,234.5600%",
        "1,234.5600%",
        "1,234.5600%",
        "1,234.5600%",
        "(1,234.5600%)",
        "1,234.5600%",
        "(1,234.5600%)",
    ],
    ["0", "00000000", "1234", "00001234"],
    [
        "0",
        "00000000",
        "10011010010",
        "10011010010",
        "11111111111111111111101100101110",
        "11111111111111111111101100101110",
        "10000011111111111111111111111110000001",
        "10000011111111111111111111111110000001",
        "10000000000000000000000000000000000000000001",
        "10000000000000000000000000000000000000000001",
    ],
    [
        "0",
        "00000000",
        "2322",
        "00002322",
        "37777775456",
        "37777775456",
        "2037777777601",
        "2037777777601",
        "200000000000001",
        "200000000000001",
    ],
    [
        "0",
        "00000000",
        "4D2",
        "000004D2",
        "FFFFFB2E",
        "FFFFFB2E",
        "20FFFFFF81",
        "20FFFFFF81",
        "80000000001",
        "80000000001",
    ],
    ["0", "00000000", "YA", "000000YA"],
    ["445/553", "70/87", "4/5", "1", "3/4", "6/8", "13/16", "8/10", "80/100", "2"],
    [
        "1E+02",
        "1.0000E+02",
        "1E+03",
        "1.0000E+03",
        "1E+04",
        "1.0000E+04",
        "1E+05",
        "1.0000E+05",
        "1E+06",
        "1.0000E+06",
    ],
]


def test_duration_formatting():
    doc = Document("tests/data/duration_112.numbers")
    for sheet in doc.sheets:
        table = sheet.tables[0]
        for row in table.iter_rows(min_row=1):
            if not isinstance(row[13], EmptyCell):
                duration = row[6].formatted_value
                ref = row[13].value
                check.equal(duration, ref)


def test_date_formatting():
    doc = Document("tests/data/date_formats.numbers")
    for sheet in doc.sheets:
        table = sheet.tables[0]
        for row in table.iter_rows(min_row=1):
            if isinstance(row[6], EmptyCell):
                assert isinstance(row[7], EmptyCell)
            else:
                date = row[6].formatted_value
                ref = row[7].value
                check.equal(date, ref)


def test_custom_formatting(pytestconfig):
    if pytestconfig.getoption("max_check_fails") is not None:
        max_check_fails = pytestconfig.getoption("max_check_fails")
    else:
        max_check_fails = -1
    doc = Document("tests/data/test-custom-formats.numbers")
    fails = 0
    for sheet in doc.sheets:
        table = sheet.tables[0]
        test_col = 1 if table.cell(0, 1).value == "Test" else 2
        for i, row in enumerate(table.iter_rows(min_row=1), start=2):
            value = row[test_col].formatted_value
            ref = row[test_col + 1].value
            check.equal(f"@{i}:{value}", f"@{i}:{ref}")
            if value != ref:
                fails += 1
            if max_check_fails > 0 and fails >= max_check_fails:
                raise AssertionError()


def test_formatting_stress(pytestconfig):
    if pytestconfig.getoption("max_check_fails") is not None:
        max_check_fails = pytestconfig.getoption("max_check_fails")
    else:
        max_check_fails = -1

    doc = Document("tests/data/custom-format-stress.numbers")
    fails = 0
    table = doc.sheets[0].tables[0]
    for i, row in enumerate(table.iter_rows(min_row=2), start=3):
        if row[9].value:
            continue
        value = row[7].formatted_value
        ref = row[8].value
        check.equal(f"@{i}:'{value}'", f"@{i}:'{ref}'", row[0].value)
        if value != ref:
            fails += 1
        if max_check_fails > 0 and fails >= max_check_fails:
            raise AssertionError()


def test_write_date_format(configurable_save_file):
    doc = Document(
        num_header_cols=0, num_header_rows=0, num_cols=len(TIME_FORMATS), num_rows=len(DATE_FORMATS)
    )
    table = doc.sheets[0].tables[0]

    ref_date = datetime(2023, 4, 1, 13, 25, 42)
    table.write("A1", ref_date)
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting("A1", "datetime", date_time_format="XX")
    assert "Invalid format specifier 'XX' in date/time format" in str(e)
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting("A1")
    assert "no type defined for cell format" in str(e)
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting("A1", "invalid")
    assert "unsuported cell format type 'invalid'" in str(e)

    table.write("A1", 0.1)
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting("A1", "datetime", date_time_format="yyyy")
    assert "cannot use date/time formatting for cells of type NumberCell" in str(e)
    table.write("A1", "test")
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting("A1", "number", date_time_format="yyyy")
    assert "cannot set formatting for cells of type TextCell" in str(e)

    for row_num in range(len(DATE_FORMATS)):
        for col_num in range(len(TIME_FORMATS)):
            date_time_format = " ".join([DATE_FORMATS[row_num], TIME_FORMATS[col_num]]).strip()
            table.write(row_num, col_num, ref_date)
            table.set_cell_formatting(
                row_num, col_num, "datetime", date_time_format=date_time_format
            )

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    for row_num, row in enumerate(DATE_FORMAT_REF):
        for col_num, ref_value in enumerate(row):
            check.equal(table.cell(row_num, col_num).value, ref_date)
            check.equal(table.cell(row_num, col_num).formatted_value, ref_value)


def test_write_numbers_format(configurable_save_file):
    doc = Document(num_header_cols=0, num_header_rows=0, num_cols=8, num_rows=10)
    table = doc.sheets[0].tables[0]

    ref_number = 12345.012346
    table.write(0, 0, ref_number)
    with pytest.warns(RuntimeWarning) as record:
        table.set_cell_formatting(
            0,
            0,
            "currency",
            negative_style=NegativeNumberStyle.RED,
            currency_code="GBP",
            use_accounting_style=True,
        )
    assert len(record) == 1
    assert str(record[0].message) == "use_accounting_style overriding negative_style"
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting(0, 0, "currency", currency_code="XYZ")
    assert "Unsupported currency code 'XYZ'" in str(e)

    row_num = 0
    for value in [ref_number, -ref_number]:
        for decimal_places in [None, 0, 1, 3, 8]:
            col_num = 0
            for show_thousands_separator in [False, True]:
                for negative_style in NegativeNumberStyle:
                    table.write(row_num, col_num, value)
                    table.set_cell_formatting(
                        row_num,
                        col_num,
                        "number",
                        decimal_places=decimal_places,
                        negative_style=negative_style,
                        show_thousands_separator=show_thousands_separator,
                    )
                    col_num += 1
            row_num += 1

    for col_num, currency_code in enumerate(
        [
            "AUD",
            "EUR",
            "GBP",
            "INR",
            "KRW",
            "JPY",
            "CNY",
            "ILS",
        ]
    ):
        table.write(10, col_num, ref_number)
        table.set_cell_formatting(
            10,
            col_num,
            "currency",
            decimal_places=2,
            negative_style=NegativeNumberStyle.MINUS,
            show_thousands_separator=False,
            currency_code=currency_code,
        )

    for col_num, currency_code in enumerate(["VND", "CAD", "SEK"]):
        table.write(11, col_num, ref_number)
        table.set_cell_formatting(
            11,
            col_num,
            "currency",
            decimal_places=2,
            negative_style=NegativeNumberStyle.MINUS,
            show_thousands_separator=False,
            currency_code=currency_code,
        )
    table.write(11, 3, ref_number)
    table.set_cell_formatting(
        11,
        3,
        "currency",
        decimal_places=2,
        show_thousands_separator=False,
        currency_code="GBP",
        use_accounting_style=True,
    )
    table.write(11, 4, -ref_number)
    table.set_cell_formatting(
        11,
        4,
        "currency",
        decimal_places=2,
        show_thousands_separator=False,
        currency_code="GBP",
        use_accounting_style=True,
    )
    table.write(11, 5, ref_number)
    table.set_cell_formatting(
        11,
        5,
        "currency",
        decimal_places=2,
        negative_style=NegativeNumberStyle.RED,
        show_thousands_separator=False,
        currency_code="GBP",
    )
    table.write(11, 6, ref_number)
    table.set_cell_formatting(
        11,
        6,
        "currency",
        decimal_places=2,
        show_thousands_separator=False,
        currency_code="GBP",
        use_accounting_style=True,
    )
    table.write(11, 7, -ref_number)
    table.set_cell_formatting(
        11,
        7,
        "currency",
        decimal_places=2,
        show_thousands_separator=True,
        currency_code="GBP",
        use_accounting_style=True,
    )

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    for row_num, row in enumerate(NUMBER_FORMAT_REF):
        for col_num, ref_value in enumerate(row):
            check.equal(abs(table.cell(row_num, col_num).value), ref_number)
            check.equal(table.cell(row_num, col_num).formatted_value, ref_value)

    for row_num, row in enumerate(CURRENCY_FORMAT_REF):
        for col_num, ref_value in enumerate(row):
            check.equal(abs(table.cell(row_num + 10, col_num).value), ref_number)
            check.equal(table.cell(row_num + 10, col_num).formatted_value, ref_value)


def test_write_mixed_number_formats(configurable_save_file):
    doc = Document(
        num_header_cols=0, num_header_rows=0, num_cols=len(TIME_FORMATS), num_rows=len(DATE_FORMATS)
    )
    table = doc.sheets[0].tables[0]

    table.write(0, 0, 0.0)
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting(0, 0, "base", base=1)
    assert "base must be in range 2-36" in str(e)
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting(0, 0, "base", base=37)
    assert "base must be in range 2-36" in str(e)

    row_num = 0
    for decimal_places in [None, 0, 1, 4]:
        for show_thousands_separator in [False, True]:
            if show_thousands_separator:
                values = [12.3456, -12.3456]
            else:
                values = [0.1234, -0.1234]
            col_num = 0
            for negative_style in NegativeNumberStyle:
                for value in values:
                    table.write(row_num, col_num, value)
                    table.set_cell_formatting(
                        row_num,
                        col_num,
                        "percentage",
                        decimal_places=decimal_places,
                        negative_style=negative_style,
                        show_thousands_separator=show_thousands_separator,
                    )
                    col_num += 1
            row_num += 1

    with pytest.raises(TypeError) as e:
        table.write(row_num, 0, -1234)
        table.set_cell_formatting(row_num, 0, "base", base=10, base_use_minus_sign=False)
    assert "base_use_minus_sign must be True for base 10" in str(e)

    for base in [10, 2, 8, 16, 36]:
        col_num = 0
        values = {
            0: True,
            1234: True,
            -1234: True,
            -1234: False,  # False formats as two's complement
            -(0x1F0000007F): False,
            (-0x7FFFFFFFFFF): False,
        }
        for value, base_use_minus_sign in zip(values.keys(), values.values()):
            for base_places in [0, 8]:
                if base not in [2, 8, 16] and not base_use_minus_sign:
                    continue
                table.write(row_num, col_num, value)
                table.set_cell_formatting(
                    row_num,
                    col_num,
                    "base",
                    base=base,
                    base_places=base_places,
                    base_use_minus_sign=base_use_minus_sign,
                )
                col_num += 1
        row_num += 1

    ref_value = 445 / 553
    for col_num, fraction_accuracy in enumerate(FractionAccuracy):
        with pytest.warns(RuntimeWarning):
            # Ignore warning about rounding the fraction
            table.write(row_num, col_num, ref_value)
        table.set_cell_formatting(row_num, col_num, "fraction", fraction_accuracy=fraction_accuracy)

    table.write(row_num, len(FractionAccuracy), 2.0)
    table.set_cell_formatting(
        row_num, len(FractionAccuracy), "fraction", fraction_accuracy=FractionAccuracy.HALVES
    )

    row_num += 1

    col_num = 0
    for value in [100, 1000, 10000, 100000, 1000000]:
        for decimal_places in [0, 4]:
            table.write(row_num, col_num, value)
            table.set_cell_formatting(row_num, col_num, "scientific", decimal_places=decimal_places)
            col_num += 1

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    for row_num, row in enumerate(OTHER_FORMAT_REF):
        for col_num, ref_value in enumerate(row):
            value = table.cell(row_num, col_num).formatted_value
            check.equal(f"[{row_num},{col_num}]:{value}", f"[{row_num},{col_num}]:{ref_value}")


def test_currency_updates():
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.write(0, 0, -12.50)
    table.set_cell_formatting(
        0, 0, "currency", currency_code="EUR", use_accounting_style=True, decimal_places=1
    )
    assert table.cell(0, 0).formatted_value == "€\t(12.5)"
