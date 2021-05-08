import pytest

from numbers_parser.document import Document

def test_many_rows():
    doc = Document("tests/data/test-3.numbers")
    sheets = doc.sheets()
    tables = sheets["Sheet_1"].tables()
    data = tables[0].data
    ref = []
    _ = [ref.append(["ROW" + str(x)]) for x in range(1, 601)]
    assert data == ref


def test_many_columns():
    doc = Document("tests/data/test-3.numbers")
    sheets = doc.sheets()
    tables = sheets["Sheet_2"].tables()
    data = tables[0].data
    ref = [["COLUMN" + str(x) for x in range(1, 601)]]
    assert data == ref
