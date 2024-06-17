import pytest

from numbers_parser import Document


def test_row_col_sizes(configurable_save_file):
    doc = Document("tests/data/test-1.numbers")
    table = doc.sheets[0].tables[0]

    # Do first to check cache init happens
    table.row_height(4, 40)
    table.col_width(2, 200)

    assert table.height == 120
    assert table.width == 396
    assert table.row_height(0) == 20
    assert table.row_height(table.num_rows - 1) == 40
    assert table.col_width(0) == 98
    assert table.col_width(table.num_cols - 1) == 200

    table.row_height(2, 40)
    assert table.row_height(2) == 40
    table.row_height(3, 10)
    table.col_width(0, 50)
    assert table.width == 348
    assert table.height == 130

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    assert table.row_height(0) == 20
    assert table.row_height(2) == 40
    assert table.row_height(3) == 10
    assert table.row_height(4) == 40
    assert table.col_width(0) == 50
    assert table.col_width(2) == 200
    assert table.height == 130
    assert table.width == 348


def test_header_size(configurable_save_file):
    doc = Document()
    table = doc.sheets[0].tables[0]

    with pytest.raises(ValueError) as e:
        table.num_header_rows = -10
    assert str(e.value) == "Number of headers cannot be negative"

    with pytest.raises(ValueError) as e:
        table.num_header_rows = 20
    assert str(e.value) == "Number of headers cannot exceed the number of rows"

    with pytest.raises(ValueError) as e:
        table.num_header_rows = 6
    assert str(e.value) == "Number of headers cannot exceed 5 rows"

    with pytest.raises(ValueError) as e:
        table.num_header_cols = -10
    assert str(e.value) == "Number of headers cannot be negative"

    with pytest.raises(ValueError) as e:
        table.num_header_cols = 20
    assert str(e.value) == "Number of headers cannot exceed the number of columns"

    with pytest.raises(ValueError) as e:
        table.num_header_cols = 6
    assert str(e.value) == "Number of headers cannot exceed 5 columns"

    assert table.num_header_rows == 1
    assert table.num_header_cols == 1
    table.num_header_rows += 1
    table.num_header_cols += 1
    assert table.num_header_rows == 2
    assert table.num_header_cols == 2

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    assert table.num_header_rows == 2
    assert table.num_header_cols == 2

    assert int(table.coordinates[0]) == 0
    assert int(table.coordinates[1]) == 0


def test_new_doc(configurable_save_file):
    doc = Document(
        sheet_name="Test Sheet",
        table_name="Test Table",
        num_header_rows=2,
        num_header_cols=2,
        num_rows=10,
        num_cols=5,
    )
    doc.save(configurable_save_file)
    new_doc = Document(configurable_save_file)
    assert new_doc.sheets[0].name == "Test Sheet"
    assert new_doc.sheets[0].tables[0].name == "Test Table"
    assert new_doc.sheets[0].tables[0].num_header_rows == 2
    assert new_doc.sheets[0].tables[0].num_rows == 10
    assert new_doc.sheets[0].tables[0].num_cols == 5
    assert new_doc.sheets[0].tables[0].cell("A1").style.bold
    assert new_doc.sheets[0].tables[0].cell("B1").style.bold
    assert new_doc.sheets[0].tables[0].cell("B2").style.bold
    assert not new_doc.sheets[0].tables[0].cell("C3").style.bold


def test_table_options(configurable_save_file):
    doc = Document()
    doc.sheets[0].tables[0].table_name_enabled = False
    doc.save(configurable_save_file)
    doc = Document(configurable_save_file)
    assert not doc.sheets[0].tables[0].table_name_enabled
