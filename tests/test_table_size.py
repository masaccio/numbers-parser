"""Tests for table dimensions, row heights, column widths, and header counts."""

from pathlib import Path

import pytest

from numbers_parser import Document


def test_row_col_sizes(configurable_save_file: Path) -> None:
    """Test getting and setting row heights and column widths."""
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


def test_header_size(configurable_save_file: Path) -> None:
    """Test getting, setting, and validating header row and column counts."""
    doc = Document()
    table = doc.sheets[0].tables[0]

    with pytest.raises(ValueError, match="Number of headers cannot be negative"):
        table.num_header_rows = -10

    with pytest.raises(ValueError, match="Number of headers cannot exceed the number of rows"):
        table.num_header_rows = 20

    with pytest.raises(ValueError, match="Number of headers cannot exceed 5 rows"):
        table.num_header_rows = 6

    with pytest.raises(ValueError, match="Number of headers cannot be negative"):
        table.num_header_cols = -10

    with pytest.raises(ValueError, match="Number of headers cannot exceed the number of columns"):
        table.num_header_cols = 20

    with pytest.raises(ValueError, match="Number of headers cannot exceed 5 columns"):
        table.num_header_cols = 6

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


def test_new_doc(configurable_save_file: Path) -> None:
    """Test creating a new document with custom dimensions and header counts."""
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
    table = new_doc.sheets[0].tables[0]
    assert new_doc.sheets[0].name == "Test Sheet"
    assert table.name == "Test Table"
    assert table.num_header_rows == 2
    assert table.num_rows == 10
    assert table.num_cols == 5

    for ref in ("A1", "B1", "B2"):
        style = table.cell(ref).style
        assert style is not None
        assert style.bold

    c3_style = table.cell("C3").style
    assert c3_style is not None
    assert not c3_style.bold


def test_table_options(configurable_save_file: Path) -> None:
    """Test disabling table name display."""
    doc = Document()
    doc.sheets[0].tables[0].table_name_enabled = False
    doc.save(configurable_save_file)
    doc = Document(configurable_save_file)
    assert not doc.sheets[0].tables[0].table_name_enabled
