import pytest

from pytest_check import check

from numbers_parser import Document, MergedCell
from numbers_parser.cell import Border, xl_rowcol_to_cell


def check_border(border, test_value):
    values = test_value.split(",")
    values[0] = float(values[0])
    values[1] = eval(values[1].replace(";", ","))
    if border is None:
        return False
    return check.equal(border, Border(values[0], values[1], values[2]))


def print_borders(table):
    for row_num, row in enumerate(table.iter_rows()):
        for col_num, cell in enumerate(row):
            if isinstance(cell, MergedCell):
                continue

            if all(
                [
                    x is None
                    for x in [
                        cell._border.top,
                        cell._border.left,
                        cell._border.right,
                        cell._border.bottom,
                    ]
                ]
            ):
                continue

            print(f"@[{row_num},{col_num}]:")
            print("   top       :", cell._border.top)
            print("   right     :", cell._border.right)
            print("   bottom    :", cell._border.bottom)
            print("   left      :", cell._border.left)


@pytest.mark.experimental
def test_borders():
    doc = Document("tests/data/test-styles.numbers")

    for sheet_name in ["Borders", "Large Borders"]:
        table = doc.sheets[sheet_name].tables[0]

        with pytest.warns() as record:
            table.cell(0, 0).border = object()
        assert len(record) == 1
        assert "cell border values cannot be set" in str(record[0])

        for row_num, row in enumerate(table.iter_rows()):
            for col_num, cell in enumerate(row):
                if isinstance(cell, MergedCell) or not cell.value:
                    continue

                # Cell test values are of the form:
                #
                # T=1,(0;162;255),dashes
                # R=1,(0;162;255),dashes
                # B=1,(0;162;255),dashes
                # L=1,(0;162;255),dashes
                tests = cell.value.split("\n")
                valid = [
                    check_border(cell.border.top, tests[0][2:]),
                    check_border(cell.border.right, tests[1][2:]),
                    check_border(cell.border.bottom, tests[2][2:]),
                    check_border(cell.border.left, tests[3][2:]),
                ]
                if not all(valid):
                    print(f"{sheet_name}@{xl_rowcol_to_cell(row_num, col_num)}: FAIL {valid}")
                    print("   reference :", ", ".join(tests))
                    print("   top       :", cell.border.top)
                    print("   right     :", cell.border.right)
                    print("   bottom    :", cell.border.bottom)
                    print("   left      :", cell.border.left)
                # assert valid
