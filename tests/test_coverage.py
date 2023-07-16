import pytest

from numbers_parser import Document, UnsupportedError, Cell
from numbers_parser.cell import xl_range, xl_rowcol_to_cell, xl_col_to_name
from numbers_parser.constants import EMPTY_STORAGE_BUFFER
from numbers_parser.cell_storage import CellStorage


def test_containers():
    doc = Document("tests/data/test-1.numbers")
    assert len(doc._model.objects) == 738


def test_ranges():
    assert xl_range(2, 2, 2, 2) == "C3"
    assert xl_col_to_name(25) == "Z"
    assert xl_col_to_name(26) == "AA"


def test_cell_storage():
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
