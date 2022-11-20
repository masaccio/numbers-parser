import pytest

from datetime import datetime as builtin_datetime, timedelta as builtin_timedelta
from pendulum import datetime, duration

from numbers_parser import Document
from numbers_parser.cell import EmptyCell
from numbers_parser.constants import MAX_ROW_COUNT, MAX_COL_COUNT


def test_edit_cell_values(tmp_path, pytestconfig):
    doc = Document("tests/data/test-save-1.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]

    with pytest.raises(IndexError) as e:
        table.write(0)
    assert "invalid cell reference" in str(e.value)
    with pytest.raises(ValueError) as e:
        table.write(0, 0, object())
    assert "determine cell type from type" in str(e.value)

    table.write("B2", "new_b2")
    table.write(1, 2, "new_c2")
    table.write(2, 0, True)
    table.write(2, 1, 7890)
    table.write(2, 2, 78.90)
    table.write(4, 3, datetime(2021, 6, 15))
    table.write(4, 4, duration(minutes=1891))
    table.write(5, 3, builtin_datetime(2020, 12, 25))
    table.write(5, 4, builtin_timedelta(seconds=7890))
    table.write(5, 5, "7890")

    assert isinstance(table.cell(3, 4), EmptyCell)
    assert isinstance(table.cell(4, 5), EmptyCell)

    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "test-save-1-new.numbers"
    doc.save(new_filename)

    doc = Document(new_filename)
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]

    assert table.cell(1, 1).value == "new_b2"
    assert table.cell("C2").value == "new_c2"
    assert table.cell(2, 0).value == True
    assert table.cell(2, 1).value == 7890
    assert table.cell(2, 2).value == 78.90
    assert table.cell(4, 3).value == datetime(2021, 6, 15)
    assert table.cell(4, 4).value == duration(minutes=1891)
    assert table.cell(5, 3).value == datetime(2020, 12, 25)
    assert table.cell(5, 4).value == duration(seconds=7890)
    assert table.cell(5, 5).value == "7890"


def test_large_table(tmp_path, pytestconfig):
    doc = Document()
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]
    for i in range(0, 300):
        table.write(i, i, "wide")

    with pytest.raises(IndexError) as e:
        table.write(MAX_ROW_COUNT, 0, "")
    assert "exceeds maximum row" in str(e.value)

    with pytest.raises(IndexError) as e:
        table.write(0, MAX_COL_COUNT, "")
    assert "exceeds maximum column" in str(e.value)

    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "save-large.numbers"
    doc.save(new_filename)

    doc = Document(new_filename)
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]

    data = table.rows()
    assert len(data) == 300
    assert len(data[299]) == 300
    assert table.cell(0, 0).value == "wide"
    assert table.cell(299, 299).value == "wide"
