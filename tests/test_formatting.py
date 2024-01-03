import pytest
import pytest_check as check
from pendulum import datetime

from numbers_parser import Document, EmptyCell
from numbers_parser.constants import NegativeNumberStyle

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
    with pytest.raises(TypeError) as e:
        table.write(0, 0, "test", formatting={"date_time_format": "XX"})
    assert "Cannot set formatting for cells of type TextCell" in str(e)
    with pytest.raises(TypeError) as e:
        table.write(0, 1, ref_date, formatting={"date_time_format": "XX"})
    assert "Invalid format specifier 'XX' in date/time format" in str(e)
    with pytest.raises(TypeError) as e:
        table.write(0, 2, ref_date, formatting=object())
    assert "formatting values must be a dict" in str(e)

    with pytest.raises(TypeError) as e:
        table.write(0, 0, ref_date, formatting={"xx": "XX"})
    assert "No date_time_format specified for DateCell formatting" in str(e)

    for row_num in range(len(DATE_FORMATS)):
        for col_num in range(len(TIME_FORMATS)):
            date_time_format = " ".join([DATE_FORMATS[row_num], TIME_FORMATS[col_num]]).strip()
            table.write(
                row_num, col_num, ref_date, formatting={"date_time_format": date_time_format}
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
    with pytest.raises(TypeError) as e:
        table.write(0, 0, "test", formatting={"decimal_places": "XX"})
    assert "Cannot set formatting for cells of type TextCell" in str(e)
    with pytest.raises(TypeError) as e:
        table.write(0, 1, ref_number, formatting={"XX": "XX"})
    assert "Invalid format specifier 'XX' in number format" in str(e)
    with pytest.raises(TypeError) as e:
        table.write(0, 2, ref_number, formatting=object())
    assert "formatting values must be a dict" in str(e)
    with pytest.warns(RuntimeWarning) as record:
        table.write(
            0,
            3,
            ref_number,
            formatting={
                "negative_style": NegativeNumberStyle.MINUS,
                "currency_code": "GBP",
                "use_accounting_style": True,
            },
        )
    assert len(record) == 1
    assert "Use of use_accounting_style overrules negative_style" in str(record[0])
    with pytest.raises(TypeError) as e:
        table.write(0, 4, ref_number, formatting={"currency_code": "XYZ"})
    assert "Unsupported currency code 'XYZ'" in str(e)

    row_num = 0
    for value in [ref_number, -ref_number]:
        for decimal_places in [None, 0, 1, 3, 8]:
            col_num = 0
            for show_thousands_separator in [False, True]:
                for negative_style in NegativeNumberStyle:
                    table.write(
                        row_num,
                        col_num,
                        value,
                        formatting={
                            "decimal_places": decimal_places,
                            "negative_style": negative_style,
                            "show_thousands_separator": show_thousands_separator,
                        },
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
        table.write(
            10,
            col_num,
            ref_number,
            formatting={
                "decimal_places": 2,
                "negative_style": NegativeNumberStyle.MINUS,
                "show_thousands_separator": False,
                "currency_code": currency_code,
            },
        )

    for col_num, currency_code in enumerate(["VND", "CAD", "SEK"]):
        table.write(
            11,
            col_num,
            ref_number,
            formatting={
                "decimal_places": 2,
                "negative_style": NegativeNumberStyle.MINUS,
                "show_thousands_separator": False,
                "currency_code": currency_code,
            },
        )
    table.write(
        11,
        3,
        ref_number,
        formatting={
            "decimal_places": 2,
            "show_thousands_separator": False,
            "currency_code": "GBP",
            "use_accounting_style": True,
        },
    )
    table.write(
        11,
        4,
        -ref_number,
        formatting={
            "decimal_places": 2,
            "show_thousands_separator": False,
            "currency_code": "GBP",
            "use_accounting_style": True,
        },
    )
    table.write(
        11,
        5,
        ref_number,
        formatting={
            "decimal_places": 2,
            "negative_style": NegativeNumberStyle.RED,
            "show_thousands_separator": False,
            "currency_code": "GBP",
        },
    )
    table.write(
        11,
        6,
        ref_number,
        formatting={
            "decimal_places": 2,
            "show_thousands_separator": False,
            "currency_code": "GBP",
            "use_accounting_style": True,
        },
    )
    table.write(
        11,
        7,
        -ref_number,
        formatting={
            "decimal_places": 2,
            "show_thousands_separator": True,
            "currency_code": "GBP",
            "use_accounting_style": True,
        },
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
