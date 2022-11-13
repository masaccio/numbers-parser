import pytest
import pytest_check as check


from numbers_parser import Document
from numbers_parser.cell import EmptyCell


def test_duration_formatting():
    doc = Document("tests/data/duration_112.numbers")
    for sheet in doc.sheets:
        table = sheet.tables[0]
        for row in table.iter_rows(min_row=1):
            if type(row[13]) != EmptyCell:
                duration = row[6].formatted_value
                ref = row[13].value
                check.equal(duration, ref)


def test_date_formatting():
    doc = Document("tests/data/date_formats.numbers")
    for sheet in doc.sheets:
        table = sheet.tables[0]
        for row in table.iter_rows(min_row=1):
            if type(row[6]) == EmptyCell:
                assert type(row[7]) == EmptyCell
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
        if table.cell(0, 1).value == "Test":
            test_col = 1
        else:
            test_col = 2
        for i, row in enumerate(table.iter_rows(min_row=1), start=1):
            value = row[test_col].formatted_value
            ref = row[test_col + 1].value
            check.equal(f"@{i}:{value}", f"@{i}:{ref}")
            if value != ref:
                fails += 1
            if max_check_fails > 0 and fails >= max_check_fails:
                assert False


@pytest.mark.experimental
def test_formatting_stress(pytestconfig):
    if pytestconfig.getoption("max_check_fails") is not None:
        max_check_fails = pytestconfig.getoption("max_check_fails")
    else:
        max_check_fails = -1

    doc = Document("tests/data/custom-format-stress.numbers")
    fails = 0
    for sheet in doc.sheets:
        table = sheet.tables[0]
        for i, row in enumerate(table.iter_rows(min_row=2), start=2):
            value = row[7].formatted_value
            ref = row[8].value
            check.equal(f"@{i}:{value}", f"@{i}:{ref}")
            if value != ref:
                fails += 1
            if max_check_fails > 0 and fails >= max_check_fails:
                assert False
