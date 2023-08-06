import json
import pytest

from numbers_parser import (
    Document,
    UnsupportedError,
    Cell,
    UnsupportedWarning,
    xl_range,
    xl_rowcol_to_cell,
    xl_col_to_name,
)
from numbers_parser.constants import EMPTY_STORAGE_BUFFER, DurationUnits
from numbers_parser.cell_storage import (
    CellStorage,
    float_to_n_digit_fraction,
    auto_units,
    decode_number_format,
)
from numbers_parser.numbers_uuid import NumbersUUID
from numbers_parser._unpack_numbers import prettify_uuids
from numbers_parser.experimental import _enable_experimental_features, _experimental_features

from numbers_parser.generated import TSKArchives_pb2 as TSKArchives


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

    assert float_to_n_digit_fraction(0.0, 1) == "0"

    format = TSKArchives.FormatStructArchive(
        format_type=268,
        duration_style=0,
        duration_unit_largest=1,
        duration_unit_smallest=0,
        use_automatic_duration_units=True,
    )
    assert auto_units(60 * 60 * 24 * 7.0, format) == (DurationUnits.WEEK, DurationUnits.WEEK)

    format = TSKArchives.FormatStructArchive(
        format_type=270,
        show_thousands_separator=False,
        use_accounting_style=False,
        fraction_accuracy=0xFFFFFFFD,
        custom_format_string="0000.##",
        scale_factor=1,
        requires_fraction_replacement=False,
        decimal_width=0,
        min_integer_width=0,
        num_nonspace_integer_digits=0,
        num_nonspace_decimal_digits=4,
        index_from_right_last_integer=4,
        num_hash_decimal_digits=0,
        total_num_decimal_digits=0,
        is_complex=False,
        contains_integer_token=False,
    )
    assert decode_number_format(format, 0.1, "test") == "    .1"

    format.custom_format_string = "0.##"
    assert decode_number_format(format, 1.0, "test") == "1"

    with pytest.raises(UnsupportedError) as e:
        _ = Document("tests/data/pre-bnc.numbers")
    assert str(e.value) == "Pre-BNC storage is unsupported"


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


def test_merge(configurable_save_file):
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.merge_cells("A1:A4")
    doc.save(configurable_save_file)

    new_doc = Document(configurable_save_file)
    table = new_doc.sheets[0].tables[0]
    new_doc._model.merge_cells(table._table_id).rect((3, 0)) == (0, 0, 3, 0)


def test_experimental():
    assert not _experimental_features()
    _enable_experimental_features(True)
    assert _experimental_features()
    _enable_experimental_features(False)
    assert not _experimental_features()


def test_prettify_uuids():
    uuid = {"upper": 0, "lower": 0}
    obj = [[1, 2, 3], ["a", "b", "c"], [uuid, uuid, uuid]]
    prettify_uuids(obj)
    assert obj[2][0] == "00000000-0000-0000-0000-000000000000"
