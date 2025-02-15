from datetime import datetime

import pytest

from numbers_parser import (
    Cell,
    CellBorder,
    CustomFormatting,
    Document,
    FileFormatError,
    Formatting,
    PaddingType,
    UnsupportedError,
    UnsupportedWarning,
)
from numbers_parser._unpack_numbers import NumbersUnpacker
from numbers_parser.cell import (
    DurationUnits,
    _auto_units,
    _decode_number_format,
    _float_to_n_digit_fraction,
    _format_decimal,
)
from numbers_parser.constants import (
    DECIMAL_PLACES_AUTO,
    EMPTY_STORAGE_BUFFER,
    NegativeNumberStyle,
)
from numbers_parser.experimental import _enable_experimental_features, _experimental_features
from numbers_parser.generated import TSKArchives_pb2 as TSKArchives
from numbers_parser.numbers_uuid import NumbersUUID
from numbers_parser.xref_utils import xl_col_to_name, xl_range, xl_rowcol_to_cell


def test_containers():
    doc = Document("tests/data/test-1.numbers")
    assert len(doc._model.objects) == 738


def test_ranges():
    assert xl_range(2, 2, 2, 2) == "C3"
    assert xl_col_to_name(25) == "Z"
    assert xl_col_to_name(26) == "AA"


def test_cell_storage(tmp_path):
    doc = Document()
    doc.sheets[0].tables[0]

    model = doc.default_table._model
    table_id = doc.default_table._table_id
    buffer = bytearray(EMPTY_STORAGE_BUFFER)
    buffer[1] = 255
    with pytest.raises(UnsupportedError) as e:
        _ = Cell._from_storage(table_id, 0, 0, buffer, model)
    assert "Cell type ID 255 is not recognised" in str(e)

    buffer = bytearray(EMPTY_STORAGE_BUFFER)
    buffer[0] = 4
    with pytest.raises(UnsupportedError) as e:
        _ = Cell._from_storage(table_id, 0, 0, buffer, model)
    assert "Cell storage version 4 is unsupported" in str(e)

    class DummyCell(Cell):
        def __init__(self, *args):
            super().__init__(*args)
            self._border = CellBorder()

    doc = Document()
    doc.sheets[0].tables[0]._data[0][0] = DummyCell(0, 0, None)
    doc.sheets[0].tables[0]._data[0][0]._model = doc._model
    doc.sheets[0].tables[0]._data[0][0]._table_id = doc.sheets[0].tables[0]._table_id
    new_filename = tmp_path / "new.numbers"
    with pytest.warns(UnsupportedWarning) as record:
        doc.save(new_filename)
    assert len(record) == 1
    assert "unsupported data type DummyCell" in str(record[0])

    assert _float_to_n_digit_fraction(0.0, 1) == "0"

    format = TSKArchives.FormatStructArchive(
        format_type=268,
        duration_style=0,
        duration_unit_largest=1,
        duration_unit_smallest=0,
        use_automatic_duration_units=True,
    )
    assert _auto_units(60 * 60 * 24 * 7.0, format) == (DurationUnits.WEEK, DurationUnits.WEEK)

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
    assert _decode_number_format(format, 0.1, "test") == "    .1"

    format.custom_format_string = "0.##"
    assert _decode_number_format(format, 1.0, "test") == "1"

    with pytest.raises(UnsupportedError) as e:
        _ = Document("tests/data/pre-bnc.numbers")
    assert str(e.value) == "Pre-BNC storage is unsupported"


def test_formatting_exceptions():
    doc = Document("tests/data/test-custom-formats.numbers")

    cell = doc.sheets[0].tables[0].cell("B4")
    format = doc._model.table_format(cell._table_id, cell._date_format_id)
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
    format = doc._model.table_format(cell._table_id, cell._num_format_id)
    format_uuid = NumbersUUID(format.custom_uid).hex
    format_map = doc._model.custom_format_map()
    custom_format = format_map[format_uuid].default_format
    custom_format.custom_format_string = "XX"
    with pytest.warns(UnsupportedWarning) as record:
        _ = cell.formatted_value
    assert len(record) == 1
    assert "Can't parse format string 'XX'" in str(record[0])


def test_pretty_uuids():
    obj = [[1, 2, 3], [4, 5, 6]]
    unpacker = NumbersUnpacker()
    unpacker.prettify_uuids(obj)
    assert str(obj) == "[[1, 2, 3], [4, 5, 6]]"


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
    assert new_doc._model.merge_cells(table._table_id).rect((3, 0)) == (0, 0, 3, 0)


def test_experimental():
    assert not _experimental_features()
    _enable_experimental_features(True)
    assert _experimental_features()
    _enable_experimental_features(False)
    assert not _experimental_features()


def test_bad_image_filenames():
    doc = Document("tests/data/issue-69b.numbers")
    table = doc.sheets[0].tables[0]
    _ = table._model.objects.file_store.pop("Data/numbers_1-16.png")
    with pytest.warns(RuntimeWarning) as record:
        assert table.cell(0, 0).style.bg_image is None
    assert len(record) == 1
    assert str(record[0].message) == "Cannot find file 'numbers_1.png' in Numbers archive"


def test_set_number_defaults():
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.write(0, 0, 0.0)
    table.set_cell_formatting(0, 0, "number")
    num_format_id = table.cell(0, 0)._num_format_id
    format = doc._model._table_formats.lookup_value(table._table_id, num_format_id).format
    assert not format.show_thousands_separator
    assert format.negative_style == NegativeNumberStyle.MINUS
    assert format.decimal_places == DECIMAL_PLACES_AUTO

    table.set_cell_formatting(0, 0, "base")
    num_format_id = table.cell(0, 0)._num_format_id
    format = doc._model._table_formats.lookup_value(table._table_id, num_format_id).format
    assert not format.show_thousands_separator
    assert format.base_places == 0

    table.set_cell_formatting(0, 0, "currency")
    num_format_id = table.cell(0, 0)._currency_format_id
    format = doc._model._table_formats.lookup_value(table._table_id, num_format_id).format
    assert not format.show_thousands_separator
    assert format.decimal_places == 2

    with pytest.raises(TypeError) as e:
        Formatting(type=object())
    assert "Invalid format type 'object'" in str(e)

    with pytest.raises(TypeError) as e:
        CustomFormatting(type=object())
    assert "Invalid format type 'object'" in str(e)


def test_formatting_empty_cell():
    doc = Document()
    assert doc.default_table.cell(0, 0).formatted_value == ""
    assert _format_decimal(None, object()) == ""


def test_custom_format_from_archive(configurable_save_file):
    doc = Document()
    doc.add_custom_format(type="datetime", name="date_format")
    doc.add_custom_format(type="text", name="text_format")
    doc.add_custom_format(type="number", name="number_format")
    doc.add_custom_format(
        type="number",
        name="number_format_1",
        num_integers=4,
        integer_format=PaddingType.ZEROS,
    )
    doc.add_custom_format(
        type="number",
        name="number_format_2",
        num_integers=7,
        integer_format=PaddingType.ZEROS,
    )
    table = doc.sheets[0].tables[0]
    table.write(0, 0, datetime(2022, 1, 1))
    table.write(0, 1, "test")
    table.write(0, 2, 1.0)
    table.write(0, 3, 1.0)
    table.write(0, 4, 1.0)
    table.set_cell_formatting(0, 3, "custom", format="number_format_1")
    table.set_cell_formatting(0, 4, "custom", format="number_format_2")
    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    table.set_cell_formatting(0, 0, "custom", format="date_format")
    table.set_cell_formatting(0, 1, "custom", format="text_format")
    table.set_cell_formatting(0, 2, "custom", format="number_format")
    assert table.cell(0, 0)._date_format_id == 3
    assert table.cell(0, 1)._text_format_id == 4
    assert table.cell(0, 2)._num_format_id == 5
    assert table.cell(0, 3).formatted_value == "0001"
    assert table.cell(0, 4).formatted_value == "0000001"


def test_cell_repr():
    doc = Document("tests/data/test-1.numbers")
    cell = doc.default_table.cell(1, 1)
    assert str(cell) == (
        "ZZZ_Sheet_1@ZZZ_Table_1[1,1]:table_id=874482, type=TEXT, value=YYY_1_1, "
        "flags=00021008, extras=0000, string_id=4, suggest_id=5, text_format_id=1"
    )


def test_invalid_format():
    doc = Document("tests/data/test-1.numbers")
    cell = doc.default_table.cell(1, 1)
    cell._text_format_id = None
    assert cell._custom_format() == cell.value


def test_invalid_format_2():
    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/invalid-index-zip.numbers")
    assert "invalid Numbers document" in str(e)
