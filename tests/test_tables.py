import pytest

from numbers_parser.document import Document

ZZZ_TABLE_1_REF = [
    ["", "YYY_COL_1", "YYY_COL_2"],
    ["YYY_ROW_1", "YYY_1_1", "YYY_1_2"],
    ["YYY_ROW_2", "YYY_2_1", "YYY_2_2"],
    ["YYY_ROW_3", "YYY_3_1", "YYY_3_2"],
    ["YYY_ROW_4", "YYY_4_1", "YYY_4_2"],
]

XXX_TABLE_1_REF = [
    ["", "XXX_COL_1", "XXX_COL_2", "XXX_COL_3", "XXX_COL_4", "XXX_COL_5"],
    ["XXX_ROW_1", "XXX_1_1", "XXX_1_2", "XXX_1_3", "XXX_1_4", "XXX_1_5"],
    ["XXX_ROW_2", "XXX_2_1", "XXX_2_2", "", "XXX_2_4", "XXX_2_5"],
    ["XXX_ROW_3", "XXX_3_1", "", "XXX_3_3", "XXX_3_4", "XXX_3_5"],
]


def test_simple_spreadsheet():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets()
    assert len(sheets) == 2
    assert sheets[0].name == "ZZZ_Sheet_1"
    assert sheets[1].name == "ZZZ_Sheet_2"
    pytest.set_trace()
    tables = sheets[0].tables()
    assert len(tables) == 2
    assert tables[0].name == "ZZZ_Table_1"
    assert tables[1].name == "ZZZ_Table_2"
    assert tables[0].data == ZZZ_TABLE_1_REF

def test_empty_cells():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets()
    tables = sheets["ZZZ_Sheet_2"].tables()
    assert len(tables) == 1
    assert tables["XXX_Table_1"].data == XXX_TABLE_1_REF
