import pytest

from numbers_parser import Document, NumberCell

ISSUE_3_REF = [("A", "B"), (2.0, 0.0), (3.0, 1.0), (None, None)]


def test_issue_3():
    doc = Document("tests/data/issue-3.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]
    ref = []
    for row in tables[0].iter_rows():
        ref.append(tuple([x.value for x in row]))
    assert ref == ISSUE_3_REF

    ref = []
    for row in tables[0].iter_rows(values_only=True):
        ref.append(row)
    assert ref == ISSUE_3_REF
