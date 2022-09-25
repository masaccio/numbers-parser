import pytest

from numbers_parser import Document


def test_row_col_sizes(tmp_path, pytestconfig):
    doc = Document("tests/data/test-1.numbers")
    table = doc.sheets[0].tables[0]
    assert table.height == 100
    assert table.width == 294
    assert table.row_height(0) == 20
    assert table.row_height(table.num_rows - 1) == 20
    assert table.col_width(0) == 98
    assert table.col_width(table.num_cols - 1) == 98

    table.row_height(2, 40)
    table.row_height(3, 10)
    table.col_width(0, 50)
    table.col_width(2, 200)
    assert table.width == 348

    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "test-1-new.numbers"
    doc.save(new_filename)

    doc = Document(new_filename)
    table = doc.sheets[0].tables[0]
    assert table.row_height(0) == 20
    assert table.row_height(2) == 40
    assert table.row_height(3) == 10
    assert table.col_width(0) == 50
    assert table.col_width(2) == 200
    assert table.height == 110
    assert table.width == 348


def test_header_size(tmp_path, pytestconfig):
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

    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "test-1-new.numbers"
    doc.save(new_filename)

    doc = Document(new_filename)
    table = doc.sheets[0].tables[0]
    assert table.num_header_rows == 2
    assert table.num_header_cols == 2
