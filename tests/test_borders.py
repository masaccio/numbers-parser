import pytest

from collections import defaultdict
from pytest_check import check

from numbers_parser import Document, MergedCell
from numbers_parser.cell import Border, xl_rowcol_to_cell, Cell, BorderType


def check_border(cell: Cell, side: str, test_value: str) -> bool:
    values = test_value.split(",")
    values[0] = float(values[0])
    values[1] = eval(values[1].replace(";", ","))
    border_value = getattr(cell.border, side, None)
    if border_value is None:
        return False
    ref = Border(values[0], values[1], values[2])
    valid = check.equal(border_value, ref)
    if not valid:
        cell_name = xl_rowcol_to_cell(cell.row, cell.col)
        print(f"@{cell_name}[{cell.row},{cell.col}].{side}: {border_value} != {ref}")
    return valid


TAG_TO_BORDER_MAP = {"T": "top", "R": "right", "B": "bottom", "L": "left"}


def unpack_test_string(test_value):
    # Cell test values are of the form:
    #
    # T=1,(0;162;255),dashes
    # R=1,(0;162;255),dashes
    # B=1,(0;162;255),dashes
    # L=1,(0;162;255),dashes
    #
    # Merge cells have multiple values T0, T1, etc.
    tests = test_value.split("\n")
    test_values = {}
    for test in tests:
        tag = TAG_TO_BORDER_MAP[test[0]]
        if test[1] == "=":
            test_values[tag] = test[2:]
        else:
            if tag not in test_values:
                test_values[tag] = defaultdict()
            offset = int(test[1])
            test_values[tag][offset] = test[3:]
    return test_values


def test_exceptions():
    with pytest.raises(TypeError) as e:
        _ = Border(width="invalid")
    assert "width must be a float number" in str(e)
    with pytest.raises(TypeError) as e:
        _ = Border(width="invalid")
    assert "width must be a float number" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Border(color=(0, 0, 0, 0))
    assert "RGB color must be an RGB" in str(e)
    with pytest.raises(TypeError) as e:
        _ = Border(color=(0, 0, 1.0))
    assert "RGB color must be an RGB" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Border(style="invalid")
    assert "invalid border style" in str(e)

    style = Border(style=BorderType(3))
    assert str(style) == "Border(width=0.35, color=RGB(r=0, g=0, b=0), style=none)"


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
                if not cell.value or isinstance(cell, MergedCell):
                    continue
                tests = unpack_test_string(cell.value)
                if cell.is_merged:
                    valid = []
                    row_start = row_num
                    row_end = row_num + cell.size[0] - 1
                    col_start = col_num
                    col_end = col_num + cell.size[1] - 1
                    offset = 0
                    for merge_row_num in range(row_start, row_end + 1):
                        merge_cell = table.cell(merge_row_num, col_num)
                        valid.append(check_border(merge_cell, "left", tests["left"][offset]))
                        merge_cell = table.cell(merge_row_num, col_end)
                        valid.append(check_border(merge_cell, "right", tests["right"][offset]))
                        offset += 1

                    offset = 0
                    for merge_col_num in range(col_start, col_end + 1):
                        merge_cell = table.cell(row_num, merge_col_num)
                        valid.append(check_border(merge_cell, "top", tests["top"][offset]))
                        merge_cell = table.cell(row_end, merge_col_num)
                        valid.append(check_border(merge_cell, "bottom", tests["bottom"][offset]))
                        offset += 1
                else:
                    valid = [
                        check_border(cell, "top", tests["top"]),
                        check_border(cell, "right", tests["right"]),
                        check_border(cell, "bottom", tests["bottom"]),
                        check_border(cell, "left", tests["left"]),
                    ]
                assert valid


def test_empty_borders():
    doc = Document("tests/data/test-styles.numbers")
    sheet = doc.sheets["Large Borders"]
    table = sheet.tables[0]

    assert table.cell("F10").border.right is None
    assert table.cell("F10").border.bottom is None
    assert table.cell("F13").border.top is None
    assert table.cell("F13").border.right is None
    assert table.cell("H10").border.left is None
    assert table.cell("H10").border.bottom is None
    assert table.cell("H13").border.left is None
    assert table.cell("H13").border.top is None


@pytest.mark.experimental
def test_create_borders(configurable_save_file):
    pass
