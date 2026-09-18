from datetime import datetime, timedelta, timezone

import pytest

from numbers_parser import Document, EmptyCell
from numbers_parser.constants import EPOCH, MAX_COL_COUNT, MAX_ROW_COUNT


def test_edit_cell_values(configurable_save_file):
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
    table.write(4, 4, timedelta(minutes=1891))
    table.write(5, 3, datetime(2020, 12, 25, tzinfo=timezone.utc))
    table.write(5, 4, timedelta(seconds=7890))
    table.write(5, 5, "7890")

    assert isinstance(table.cell(3, 4), EmptyCell)
    assert isinstance(table.cell(4, 5), EmptyCell)

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]

    assert table.cell(1, 1).value == "new_b2"
    assert table.cell("C2").value == "new_c2"
    assert table.cell(2, 0).value
    assert table.cell(2, 1).value == 7890
    assert table.cell(2, 2).value == 78.90
    assert table.cell(4, 3).value == datetime(2021, 6, 15)
    assert table.cell(4, 4).value == timedelta(minutes=1891)
    expected_local_date = EPOCH + (
        datetime(2020, 12, 25, tzinfo=timezone.utc) - EPOCH.astimezone(timezone.utc)
    )
    assert table.cell(5, 3).value == expected_local_date
    assert table.cell(5, 4).value == timedelta(seconds=7890)
    assert table.cell(5, 5).value == "7890"


def test_large_table(configurable_save_file):
    doc = Document()
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]
    for i in range(300):
        table.write(i, i, "wide")

    with pytest.raises(ValueError) as e:  # noqa: PT011
        table.add_row(MAX_ROW_COUNT - table.num_rows + 1)
    assert "rows cannot exceed" in str(e.value)

    with pytest.raises(IndexError) as e:
        table.write(MAX_ROW_COUNT, 0, "")
    assert "exceeds maximum row" in str(e.value)

    with pytest.raises(ValueError) as e:  # noqa: PT011
        table.add_column(MAX_COL_COUNT - table.num_cols + 1)
    assert "columns cannot exceed" in str(e.value)

    with pytest.raises(IndexError) as e:
        table.write(0, MAX_COL_COUNT, "")
    assert "exceeds maximum column" in str(e.value)

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]

    data = table.rows()
    assert len(data) == 300
    assert len(data[299]) == 300
    assert table.cell(0, 0).value == "wide"
    assert table.cell(299, 299).value == "wide"
