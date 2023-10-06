import pytest
from pendulum import Duration, datetime

from numbers_parser import Document, EmptyCell, MergedCell, NumberCell, TextCell
from numbers_parser.constants import (
    DEFAULT_COLUMN_COUNT,
    DEFAULT_NUM_HEADERS,
    DEFAULT_ROW_COUNT,
    DEFAULT_ROW_HEIGHT,
    DEFAULT_TABLE_OFFSET,
)


def test_empty_document(configurable_save_file):
    doc = Document()
    data = doc.sheets[0].tables[0].rows()
    assert len(data) == DEFAULT_ROW_COUNT
    assert len(data[0]) == DEFAULT_COLUMN_COUNT
    assert isinstance(data[0][0], EmptyCell)
    assert doc.sheets[0].tables[0].num_header_rows == DEFAULT_NUM_HEADERS
    assert doc.sheets[0].tables[0].num_header_cols == DEFAULT_NUM_HEADERS

    doc.save(configurable_save_file)
    doc = Document(configurable_save_file)
    assert len(data) == DEFAULT_ROW_COUNT
    assert len(data[0]) == DEFAULT_COLUMN_COUNT
    assert isinstance(data[0][0], EmptyCell)
    assert doc.sheets[0].tables[0].num_header_rows == DEFAULT_NUM_HEADERS
    assert doc.sheets[0].tables[0].num_header_cols == DEFAULT_NUM_HEADERS


def test_save_document(tmp_path):
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    cell_values = tables["ZZZ_Table_1"].rows(values_only=True)

    sheets["ZZZ_Sheet_1"].name = "ZZZ_Sheet_1 NEW"
    tables["ZZZ_Table_1"].name = "ZZZ_Table_1 NEW"

    new_filename = tmp_path / "test-1-new.numbers"
    doc.save(new_filename)

    new_doc = Document(new_filename)
    new_sheets = new_doc.sheets
    new_tables = new_sheets["ZZZ_Sheet_1 NEW"].tables
    new_cell_values = new_tables["ZZZ_Table_1 NEW"].rows(values_only=True)

    assert cell_values == new_cell_values


def test_save_merges(configurable_save_file):
    doc = Document("tests/data/test-save-1.numbers")
    sheets = doc.sheets
    table = sheets[0].tables[0]
    table.add_column(2)
    table.add_row(3)
    table.write("B2", "merge_1")
    table.write("B5", "merge_2")
    table.write("D2", "merge_3")
    table.write("F4", "")
    table.merge_cells("B2:C2")
    table.merge_cells(["B5:E5", "D2:F4"])

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    sheets = doc.sheets
    table = sheets[0].tables[0]
    assert table.cell("B2").is_merged
    assert table.cell("B2").size == (1, 2)
    assert table.cell("B5").is_merged
    assert table.cell("B5").size == (1, 4)
    assert table.cell("D2").is_merged
    assert table.cell("D2").size == (3, 3)
    assert table.merge_ranges == ["B2:C2", "B5:E5", "D2:F4"]
    assert all(isinstance(table.cell(row, 3), MergedCell) for row in range(2, 4))
    assert all(isinstance(table.cell(row, 4), MergedCell) for row in range(1, 4))


def test_create_table(configurable_save_file):
    doc = Document()
    sheets = doc.sheets

    with pytest.raises(IndexError) as e:
        _ = sheets[0].add_table("TablE 1")
    assert "table 'TablE 1' already exists" in str(e.value)

    table = sheets[0].add_table()
    assert table.name == "Table 2"
    table.write("B1", "Column B")
    table.write("C1", "Column C")
    table.write("D1", "Column D")
    table.write("B2", "Mary had")
    table.write("C2", "a little")
    table.write("D2", "lamb")

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    sheets = doc.sheets

    table = sheets[0].tables[1]
    assert sheets[0].tables[1].name == "Table 2"
    assert table.cell("B1").value == "Column B"
    assert table.cell("C1").value == "Column C"
    assert table.cell("D1").value == "Column D"
    assert table.cell("B2").value == "Mary had"
    assert table.cell("C2").value == "a little"
    assert table.cell("D2").value == "lamb"


def test_create_sheet(configurable_save_file):
    doc = Document()
    sheets = doc.sheets

    with pytest.raises(IndexError) as e:
        _ = doc.add_sheet("SheeT 1")
    assert "sheet 'SheeT 1' already exists" in str(e.value)

    doc.add_sheet("New Sheet", "New Table")
    sheet = doc.sheets["New Sheet"]
    table = sheet.tables["New Table"]
    table.write(0, 1, "Column 1")
    table.write(0, 2, "Column 2")
    table.write(0, 3, "Column 3")
    table.write(1, 1, 1000)
    table.write(1, 2, 2000)
    table.write(1, 3, 3000)

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    sheets = doc.sheets

    assert sheets[1].name == "New Sheet"

    table = sheets[1].tables[0]
    assert table.name == "New Table"

    assert type(table.cell(0, 1)) == TextCell
    assert table.cell(0, 1).value == "Column 1"
    assert type(table.cell(0, 1)) == TextCell
    assert type(table.cell(1, 3)) == NumberCell
    assert table.cell(1, 3).value == 3000


def test_create_multi(configurable_save_file):
    doc = Document()

    doc.sheets[0].tables[0].write(0, 0, "S0T1 A1")

    doc.add_sheet()
    doc.sheets[1].add_table()
    doc.sheets[1].add_table()
    doc.sheets[1].tables[0].write(0, 0, "S1T0 A1")
    doc.sheets[1].tables[1].write(0, 0, "S1T1 A1")
    doc.sheets[1].tables[2].write(0, 0, "S1T2 A1")

    doc.add_sheet()
    offset = 700
    doc.sheets[2].add_table(x=100.0)
    doc.sheets[2].add_table(x=0.0, y=offset)
    doc.sheets[2].add_table()
    doc.sheets[2].tables[0].write(0, 0, "S2T0 A1")
    doc.sheets[2].tables[1].write(0, 0, "S2T1 A1")
    doc.sheets[2].tables[2].write(0, 0, "S2T2 A1")
    doc.sheets[2].tables[3].write(0, 0, "S2T3 A1")

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    height1 = DEFAULT_ROW_HEIGHT * DEFAULT_ROW_COUNT
    height2 = height1 + offset + DEFAULT_TABLE_OFFSET
    assert doc.sheets[1].tables[1].height == height1
    assert doc.sheets[2].tables[1].coordinates == (100.0, height1 + DEFAULT_TABLE_OFFSET)
    assert doc.sheets[2].tables[2].coordinates == (0.0, offset)
    assert doc.sheets[2].tables[3].coordinates == (0.0, height2)
    assert len(doc.sheets) == 3
    assert len(doc.sheets[2].tables) == 4


def test_duplicate_name():
    doc = Document()
    doc.sheets[0].add_table("Test")
    with pytest.raises(IndexError) as e:
        doc.sheets[0].add_table("Test")
    assert "table 'Test' already exists" in str(e)


def test_save_formats(configurable_save_file):
    doc = Document("tests/data/test-format-save.numbers")
    doc.save(configurable_save_file)
    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    assert table.cell("A1").value
    assert table.cell("A2").value == 123.45
    assert table.cell("A3").value == Duration(days=5)
    assert table.cell("A4").value == datetime(2000, 12, 1)
