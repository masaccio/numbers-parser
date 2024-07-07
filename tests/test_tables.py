import pytest

from numbers_parser import Document, TextCell
from numbers_parser.exceptions import FileError, FileFormatError

ZZZ_TABLE_1_REF = [
    [None, "YYY_COL_1", "YYY_COL_2"],
    ["YYY_ROW_1", "YYY_1_1", "YYY_1_2"],
    ["YYY_ROW_2", "YYY_2_1", "YYY_2_2"],
    ["YYY_ROW_3", "YYY_3_1", "YYY_3_2"],
    ["YYY_ROW_4", "YYY_4_1", "YYY_4_2"],
]

ZZZ_TABLE_2_REF = [
    [None, "ZZZ_COL_1", "ZZZ_COL_2", "ZZZ_COL_3"],
    ["ZZZ_ROW_1", "ZZZ_1_1", "ZZZ_1_2", "ZZZ_1_3"],
    ["ZZZ_ROW_2", "ZZZ_2_1", "ZZZ_2_2", "ZZZ_2_3"],
    ["ZZZ_ROW_3", "ZZZ_3_1", "ZZZ_3_2", "ZZZ_3_3"],
]

XXX_TABLE_1_REF = [
    [None, "XXX_COL_1", "XXX_COL_2", "XXX_COL_3", "XXX_COL_4", "XXX_COL_5"],
    ["XXX_ROW_1", "XXX_1_1", "XXX_1_2", "XXX_1_3", "XXX_1_4", "XXX_1_5"],
    ["XXX_ROW_2", "XXX_2_1", "XXX_2_2", None, "XXX_2_4", "XXX_2_5"],
    ["XXX_ROW_3", "XXX_3_1", None, "XXX_3_3", "XXX_3_4", "XXX_3_5"],
]

EMPTY_ROWS_REF = [
    ["A", "B", "C", None, None, None],
    ["A_1", "B_1", "C_1", None, None, None],
    ["A_2", "B_2", "C_2", None, None, None],
    [None, None, None, None, None, None],
    [None, None, None, None, None, None],
    ["A_5", "B_5", "C_5", None, None, None],
    ["A_6", "B_6", "C_6", None, None, None],
    [None, None, None, None, None, None],
    ["A_8", None, "C_8", None, None, "F_8"],
    ["A_9", None, "C_9", None, None, None],
]


def test_file_exceptions():
    with pytest.raises(FileError) as e:
        _ = Document("tests/data/test-1.not-found")
    assert str(e.value) == "no such file or directory"
    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/conftest.py")
    assert str(e.value) == "invalid Numbers document (not a .numbers package/file)"


def test_sheet_exceptions():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets
    with pytest.raises(IndexError) as e:
        _ = sheets[3]
    assert "out of range" in str(e.value)
    with pytest.raises(KeyError) as e:
        _ = sheets["invalid"]
    assert "no sheet named" in str(e.value)
    with pytest.raises(LookupError) as e:
        _ = sheets[float(1)]
    assert "invalid index" in str(e.value)


def test_names_refs():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets
    assert len(sheets) == 2
    assert sheets[0].name == "ZZZ_Sheet_1"
    assert sheets[1].name == "ZZZ_Sheet_2"
    tables = sheets["ZZZ_Sheet_1"].tables
    assert len(tables) == 2
    assert tables[0].name == "ZZZ_Table_1"
    assert tables[1].name == "ZZZ_Table_2"
    assert tables["ZZZ_Table_1"].num_cols == 3
    assert tables["ZZZ_Table_1"].num_rows == 5


def test_table_contents():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    data = tables[0].rows(values_only=True)
    assert data == ZZZ_TABLE_1_REF
    data = tables[1].rows(values_only=True)
    assert data == ZZZ_TABLE_2_REF


def test_table_contents_as_cells():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    rows = tables[0].rows()
    for index in range(len(rows)):
        data = [cell.value or None for cell in rows[index]]
        assert data == ZZZ_TABLE_1_REF[index]


def test_cell():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    assert isinstance(tables[0].cell("A2"), TextCell)
    assert tables[0].cell("A2").value == "YYY_ROW_1"


def test_empty_cells():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets
    tables = sheets["ZZZ_Sheet_2"].tables
    assert len(tables) == 1
    table = tables["XXX_Table_1"]
    assert table.num_cols == 6
    assert table.num_rows == 4
    assert table.rows(values_only=True) == XXX_TABLE_1_REF


def test_empty_rows():
    doc = Document("tests/data/test-empty-rows.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]
    assert table.rows(values_only=True) == EMPTY_ROWS_REF


def test_default_table():
    doc = Document()
    assert doc.default_table.name == "Table 1"


def table_titles_test_runner(doc):
    for table in doc.sheets[0].tables:
        test_type = table.cell(0, 0).value
        assert table.name == test_type

        if "no title" in test_type:
            assert not table.table_name_enabled
        else:
            assert table.table_name_enabled

        if "no caption" in test_type:
            assert not table.caption_enabled
            assert table.caption == "caption"
        elif test_type == "new table":
            assert table.caption_enabled
            assert table.caption == "new caption"
        else:
            assert table.caption_enabled
            assert table.caption == "test: caption"


def test_table_titles(configurable_save_file):
    doc = Document("tests/data/test-titles.numbers")
    table_titles_test_runner(doc)

    for table in doc.sheets[0].tables:
        table.table_name_enabled = not table.table_name_enabled
        table.caption_enabled = not table.caption_enabled
        new_test_type = "test: "
        if table.table_name_enabled:
            new_test_type += "title, "
        else:
            new_test_type += "no title, "
        if table.caption_enabled:
            table.caption = "test: caption"
            new_test_type += "caption"
        else:
            new_test_type += "no caption"
            table.caption = "caption"

        table.name = new_test_type
        table.write(0, 0, new_test_type)

    new_table = doc.sheets[0].add_table(table_name="new table", x=0, y=650)
    new_table.caption = "new caption"
    new_table.write(0, 0, "new table")
    new_table.caption_enabled = True

    doc.save(configurable_save_file)

    new_doc = Document(configurable_save_file)
    table_titles_test_runner(new_doc)


def test_stub_captions(configurable_save_file):
    doc = Document("tests/data/test-1.numbers")
    table0 = doc.sheets[0].tables[0]
    assert table0.caption == "Caption"
    assert not table0.caption_enabled
    table0.caption_enabled = True
    table0.caption = "New Caption"
    doc.save(configurable_save_file)

    doc = Document()
    table0 = doc.sheets[0].tables[0]
    assert table0.caption == "Caption"
    assert not table0.caption_enabled
    table0.caption_enabled = True
    table0.caption = "New Caption"
    table1 = doc.sheets[0].add_table()
    assert table1.caption == "Caption"
    assert not table1.caption_enabled
    doc.save(configurable_save_file)
