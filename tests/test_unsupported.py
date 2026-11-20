"""Tests for unsupported table features and error cells."""

from numbers_parser import Document, ErrorCell


def test_table_unsupported() -> None:
    """Test reading an ErrorCell from a table."""
    doc = Document("tests/data/test-8.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    assert isinstance(tables[0].cell(3, 2), ErrorCell)
