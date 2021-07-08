import pytest

from numbers_parser import Document

XXX_TABLE_1_REF = [
    [None, "XXX_COL_1", "XXX_COL_2", "XXX_COL_3", "XXX_COL_4", "XXX_COL_5" ],
    ["XXX_ROW_1", "XXX_1_1", "XXX_1_2", "XXX_1_3", "XXX_1_4", "XXX_1_5" ],
    ["XXX_ROW_2", "XXX_2_1", "XXX_2_2", None, "XXX_2_4", "XXX_2_5" ],
    ["XXX_ROW_3", "XXX_3_1", None, "XXX_3_3", "XXX_3_4", "XXX_3_5" ],
]


def test_read_folder():
    doc = Document("tests/data/test-5.numbers")
    sheets = doc.sheets()
    tables = sheets["ZZZ_Sheet_1"].tables()
    data = tables["XXX_Table_1"].rows(values_only=True)
    assert data == XXX_TABLE_1_REF
