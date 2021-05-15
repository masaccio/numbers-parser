import pytest

from numbers_parser.document import Document

def test_table_unsupported():
    doc = Document("tests/data/test-8.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    assert tables[0].cell(1, 5) == "*FORMULA*"
    assert tables[0].cell(3, 2) == "*ERROR*"
