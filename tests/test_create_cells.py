import pytest
import os
from tempfile import TemporaryDirectory

from numbers_parser import Document
from numbers_parser.cell import EmptyCell


@pytest.mark.experimental
def test_edit_strings():
    doc = Document("tests/data/test-save-1.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]
    table.write("B2", "new_b2")
    table.write(2, 1, "new_c2")
    table.write(5, 4, "new_f7")

    assert isinstance(table.cell(3, 4), EmptyCell)
    assert isinstance(table.cell(4, 4), EmptyCell)

    temp_dir = TemporaryDirectory()
    new_filename = os.path.join(temp_dir.name, "test-save-1-new.numbers")
    doc.save(new_filename)

    doc = Document(new_filename)
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]

    assert table.cell(1, 1).value == "new_b2"
    assert table.cell(2, 1).value == "new_c2"
    assert table.cell(5, 4).value == "new_f7"
