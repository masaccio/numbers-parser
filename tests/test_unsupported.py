import pytest

from numbers_parser import Document, FormulaCell, ErrorCell

def test_table_unsupported():
    doc = Document("tests/data/test-8.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    assert isinstance(tables[0].cell(1, 5), FormulaCell)
    assert isinstance(tables[0].cell(3, 2), ErrorCell)
