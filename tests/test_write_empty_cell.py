import pytest

from numbers_parser import Document, EmptyCell


@pytest.mark.parametrize("coordinate", [(1, 1), ("B2",)])
def test_write_none_clears_cell_and_survives_save(coordinate, tmp_path):
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.write(*coordinate, "old value")

    table.write(*coordinate, None)

    cell = table.cell(1, 1)
    assert isinstance(cell, EmptyCell)
    assert cell.value is None
    assert cell.formatted_value == ""

    path = tmp_path / "cleared.numbers"
    doc.save(path)
    reloaded_cell = Document(path).sheets[0].tables[0].cell(1, 1)
    assert isinstance(reloaded_cell, EmptyCell)
    assert reloaded_cell.value is None


def test_write_none_on_existing_empty_cell():
    table = Document().sheets[0].tables[0]
    table.write("B2", None)
    assert isinstance(table.cell("B2"), EmptyCell)
