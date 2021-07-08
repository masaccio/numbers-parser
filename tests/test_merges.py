import pytest

from numbers_parser import Document, Cell, MergedCell

XXX_TABLE_1_REF = [
    ["XXX_COL_1", "XXX_COL_2", "XXX_COL_3", "XXX_COL_4", "XXX_COL_5"],
    ["XXX_1_1__1_2", None, "XXX_1_3", "XXX_1_4", "XXX_1_5"],
    ["XXX_2_1", "XXX_2_2", "XXX_2_3", "XXX_2_4", "XXX_2_5"],
    ["XXX_3_1", "XXX_3_2", "XXX_3_3__3_5", None, None],
    ["XXX_4_1", "XXX_4_2__4_5", None, None, None],
    ["XXX_5_1", "XXX_5_2__XXX_7_2", "XXX_5_3", "XXX_5_4", "XXX_5_5"],
    ["XXX_6_1", None, "XXX_6_3", "XXX_6_4__XXX_7_5", None],
    ["XXX_7_1", None, "XXX_7_3", None, None],
]


XXX_TABLE_1_CLASSES = [
    ["TextCell", "TextCell", "TextCell", "TextCell", "TextCell"],
    ["TextCell", "MergedCell", "TextCell", "TextCell", "TextCell"],
    ["TextCell", "TextCell", "TextCell", "TextCell", "TextCell"],
    ["TextCell", "TextCell", "TextCell", "MergedCell", "MergedCell"],
    ["TextCell", "TextCell", "MergedCell", "MergedCell", "MergedCell"],
    ["TextCell", "TextCell", "TextCell", "TextCell", "TextCell"],
    ["TextCell", "MergedCell", "TextCell", "TextCell", "MergedCell"],
    ["TextCell", "MergedCell", "TextCell", "MergedCell", "MergedCell"],
]


def test_table_contents():
    doc = Document("tests/data/test-9.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    data = tables[0].rows(values_only=True)
    assert data == XXX_TABLE_1_REF


def test_cell_classes():
    doc = Document("tests/data/test-9.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    data = []
    for row in tables[0].iter_rows():
        data.append([type(c).__name__ for c in row])
    assert data == XXX_TABLE_1_CLASSES


def test_merge_references():
    doc = Document("tests/data/test-9.numbers")
    sheets = doc.sheets()
    table = sheets[0].tables()[0]
    assert table.cell("B2").merge_range == "A2:B2"
    assert table.cell("C5").merge_range == "B5:E5"
    assert table.cell("B2").size == None
    assert table.cell("C5").size == None
    assert table.cell("C5").row_start == 4
    assert table.cell("C5").row_end == 4
    assert table.cell("C5").col_start == 1
    assert table.cell("C5").col_end == 4
    assert table.cell("A2").is_merged
    assert table.cell("C1").is_merged == False
    assert table.cell("D7").size == (2, 2)

    table = sheets[1].tables()[0]
    assert table.cell("A1").is_merged
    assert table.cell("A1").size == (1, 2)
    assert table.cell("B4").is_merged
    assert table.cell("B4").size == (2, 2)


def test_all_merged_ranges():
    doc = Document("tests/data/test-9.numbers")
    sheets = doc.sheets()
    table = sheets[0].tables()[0]
    assert table.merge_ranges == ["A2:B2", "B5:E5", "B6:B8", "C4:E4", "D7:E8"]
    table = sheets[1].tables()[0]
    assert table.merge_ranges == ["A1:B1", "B4:C5"]
