import pytest

from numbers_parser import Document, UnsupportedError, Cell, UnsupportedWarning
from numbers_parser.cell import xl_range, xl_rowcol_to_cell, xl_col_to_name
from numbers_parser.constants import EMPTY_STORAGE_BUFFER
from numbers_parser.cell_storage import CellStorage
from numbers_parser.numbers_uuid import NumbersUUID


def test_containers():
    doc = Document("tests/data/test-1.numbers")
    assert len(doc._model.objects) == 738


def test_ranges():
    assert xl_range(2, 2, 2, 2) == "C3"
    assert xl_col_to_name(25) == "Z"
    assert xl_col_to_name(26) == "AA"


def test_cell_storage(tmp_path):
    doc = Document()
    table = doc.sheets[0].tables[0]

    buffer = bytearray(EMPTY_STORAGE_BUFFER)
    buffer[1] = 255
    with pytest.raises(UnsupportedError) as e:
        storage = CellStorage(doc._model, table._table_id, bytes(buffer), 0, 0)
        _ = Cell.from_storage(storage)
    assert "Cell type ID 255 is not recognised" in str(e)

    buffer = bytearray(EMPTY_STORAGE_BUFFER)
    with pytest.raises(UnsupportedError) as e:
        storage = CellStorage(doc._model, table._table_id, bytes(buffer), 0, 0)
        storage.type = 255
        _ = Cell.from_storage(storage)
    assert "Unsupported cell type 255 @:(0,0)" in str(e)

    buffer = bytearray(EMPTY_STORAGE_BUFFER)
    buffer[0] = 4
    with pytest.raises(UnsupportedError) as e:
        storage = CellStorage(doc._model, table._table_id, bytes(buffer), 0, 0)
    assert "Cell storage version 4 is unsupported" in str(e)

    class DummyCell(Cell):
        pass

    doc = Document()
    doc.sheets[0].tables[0]._data[0][0] = DummyCell(0, 0, None)
    new_filename = tmp_path / "new.numbers"
    with pytest.warns(UnsupportedWarning) as record:
        doc.save(new_filename)
    assert len(record) == 1
    assert "unsupported data type DummyCell" in str(record[0])


def test_formatting_exceptions():
    doc = Document("tests/data/test-custom-formats.numbers")

    cell = doc.sheets[0].tables[0].cell("B4")
    format = doc._model.table_format(cell._table_id, cell._storage.date_format_id)
    format_uuid = NumbersUUID(format.custom_uid).hex
    format_map = doc._model.custom_format_map()
    custom_format = format_map[format_uuid].default_format

    custom_format.format_type = 299
    with pytest.warns(UnsupportedWarning) as record:
        _ = cell.formatted_value
    assert len(record) == 1
    assert "Unexpected custom format type 299" in str(record[0])

    custom_format.format_type = 272
    custom_format.custom_format_string = "ZZ"
    with pytest.warns(UnsupportedWarning) as record:
        _ = cell.formatted_value
    assert len(record) == 1
    assert "Unsupported field code 'ZZ'" in str(record[0])

    cell = doc.sheets["Numbers"].tables[0].cell("C38")
    format = doc._model.table_format(cell._table_id, cell._storage.num_format_id)
    format_uuid = NumbersUUID(format.custom_uid).hex
    format_map = doc._model.custom_format_map()
    custom_format = format_map[format_uuid].default_format
    custom_format.custom_format_string = "XX"
    with pytest.warns(UnsupportedWarning) as record:
        _ = cell.formatted_value
    assert len(record) == 1
    assert "Can't parse format string 'XX'" in str(record[0])


def test_range_exceptions():
    with pytest.raises(IndexError) as e:
        _ = xl_rowcol_to_cell(-2, 0)
    assert "row reference -2 below zero" in str(e)

    with pytest.raises(IndexError) as e:
        _ = xl_rowcol_to_cell(0, -2)
    assert "column reference -2 below zero" in str(e)

    with pytest.raises(IndexError) as e:
        _ = xl_col_to_name(-2)
    assert "column reference -2 below zero" in str(e)


def test_datalists():
    doc = Document("tests/data/test-1.numbers")
    table_id = doc.sheets[0].tables[0]._table_id
    assert doc._model.table_string(table_id, 1) == "YYY_2_1"
    assert doc._model._table_strings.lookup_key(table_id, "TEST") == 19
    assert doc._model._table_strings.lookup_key(table_id, "TEST") == 19
    assert doc._model._table_strings.lookup_value(table_id, 19).string == "TEST"
