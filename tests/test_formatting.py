import pytest
import pytest_check as check
from pendulum import datetime

from numbers_parser import Document, EmptyCell

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
