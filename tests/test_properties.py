from numbers_parser import Document
from pendulum import datetime, duration

ZZZ_TABLE_1_REF = [
    ["YYY_1_1", "YYY_1_2"],
    ["YYY_2a_1", None],
    ["YYY_2b_1", None],
    ["YYY_3_1", "YYY_3_2"],
    ["YYY_4_1", "YYY_4_2"],
]


ZZZ_TABLE_2_REF = [
    [None, "ZZZ_COL_1", "ZZZ_COL_2"],
    ["TEXT", "ZZZ_1_1", "ZZZ_1_2"],
    [1.000, 2.000, 123.456],
    [datetime(2001, 1, 1), datetime(2001, 1, 20), datetime(2002, 1, 1)],
    [
        duration(seconds=1, microseconds=1000),
        duration(days=1),
        duration(days=10),
    ],
    [False, True, False],
    ["STRING1", "STRING2", "STRING3"],
]


YYY_TABLE_1_REF = [
    [None, "AAA_COL_1", "AAA_COL_2", "AAA_COL_3", "AAA_COL_4", "AAA_COL_5"],
    ["AAA_ROW_1", "AAA_1_1", "AAA_2_1", "AAA_3_1", "AAA_4_1", "AAA_5_1", "AAA_6_1"],
    ["AAA_ROW_2", "AAA_1_2", None, "AAA_3_2", "AAA_4_2", "AAA_5_2", "AAA_6_2"],
    ["AAA_ROW_3", "AAA_1_3", None, None, "AAA_4_3", "AAA_5_3", "AAA_6_3"],
    ["AAA_ROW_4", "AAA_1_4", None, None, None, "AAA_5_4", "AAA_6_4"],
    [None, "AAA_1_5", None, None, None, None, None],
    ["AAA_ROW_6", "AAA_1_6", "AAA_2_6", "AAA_3_6", "AAA_4_6", "AAA_5_6"],
    ["AAA_ROW_7", "AAA_1_7", "AAA_2_7", "AAA_3_7", "AAA_4_7"],
    [None, "AAA_1_8", "AAA_2_8", "AAA_3_8", None, None, None],
    ["AAA_ROW_8", "AAA_1_9", "AAA_2_9", None, None, None, None],
]


def test_cell_merging():
    doc = Document("tests/data/test-4.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    data = tables["ZZZ_Table_1"].rows(values_only=True)
    assert data == ZZZ_TABLE_1_REF


def test_cell_types():
    doc = Document("tests/data/test-4.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    data = tables["ZZZ_Table_2"].rows(values_only=True)
    assert data == ZZZ_TABLE_2_REF
