import pytest

from datetime import datetime, timedelta

from numbers_parser import Document
from numbers_parser.cell import EmptyCell


# @pytest.mark.experimental
def test_edit_cell_values(tmp_path):
    doc = Document("tests/data/test-save-1.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]
    table.write("B2", "new_b2")
    table.write(1, 2, "new_c2")
    table.write(2, 0, True)
    table.write(2, 1, 7890)
    table.write(2, 2, 78.90)
    table.write(5, 3, datetime(2020, 12, 25))
    table.write(5, 4, timedelta(seconds=7890))
    table.write(5, 5, "7890")

    assert isinstance(table.cell(3, 4), EmptyCell)
    assert isinstance(table.cell(4, 4), EmptyCell)

    new_filename = tmp_path / "test-save-1-new.numbers"
    doc.save(new_filename)
    print(f"\nSAVE {new_filename}")

    doc = Document(new_filename)
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]

    assert table.cell(1, 1).value == "new_b2"
    assert table.cell("C2").value == "new_c2"
    assert table.cell(2, 0).value == True
    assert table.cell(2, 1).value == 7890
    assert table.cell(2, 2).value == 78.90
    assert table.cell(5, 3).value == datetime(2020, 12, 25)
    assert table.cell(5, 4).value == timedelta(seconds=7890)
    assert table.cell(5, 5).value == "7890"
