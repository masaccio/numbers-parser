import pytest
from datetime import datetime

from numbers_parser.document import Document

ZZZ_TABLE_2_REF = [
    [None, "ZZZ_COL_1", "ZZZ_COL_2"],
    ["TEXT", "ZZZ_1_1", "ZZZ_1_2"],
    [1.000, 2.000, 123.456],
    [datetime(2001, 1, 1), datetime(2001, 1, 20), datetime(2002, 1, 1)],
    [1.001, 86400.0, 864000.0],
    [False, True, False],
    ["STRING1", "STRING2", "STRING3"],
]


def test_table_contents():
    doc = Document("tests/data/test-4.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    data = tables["ZZZ_Table_2"].data
    assert data == ZZZ_TABLE_2_REF
