from datetime import datetime, timedelta
from typing import Optional

import pytest

from numbers_parser import Document, EmptyCell, MergedCell, NumberCell, TextCell
from numbers_parser.constants import (
    DEFAULT_COLUMN_COUNT,
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
    assert doc.sheets[0].tables[0].num_header_rows == 1
    assert doc.sheets[0].tables[0].num_header_cols == 1

    doc.save(configurable_save_file)
    doc = Document(configurable_save_file)
    assert len(data) == DEFAULT_ROW_COUNT
    assert len(data[0]) == DEFAULT_COLUMN_COUNT
    assert isinstance(data[0][0], EmptyCell)
    assert doc.sheets[0].tables[0].num_header_rows == 1
    assert doc.sheets[0].tables[0].num_header_cols == 1


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
    assert table.cell(table.num_rows - 1, table.num_cols - 1).row == table.num_rows - 1
    assert table.cell(table.num_rows - 1, table.num_cols - 1).col == table.num_cols - 1
    table.write("B2", "merge_1")
    table.write("B5", "merge_2")
    table.write("D2", "merge_3")
    table.write("F4", "")
    table.merge_cells("B2:C2")
    table.merge_cells(["B5:E5", "D2:F4"])
    assert table.cell("B2").is_merged
    assert table.cell("B2").size == (1, 2)
    assert table.cell("B5").is_merged
    assert table.cell("B5").size == (1, 4)
    assert table.cell("D2").is_merged
    assert table.cell("D2").size == (3, 3)
    assert table.merge_ranges == ["B2:C2", "B5:E5", "D2:F4"]

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

    assert isinstance(table.cell(0, 1), TextCell)
    assert table.cell(0, 1).value == "Column 1"
    assert isinstance(table.cell(0, 1), TextCell)
    assert isinstance(table.cell(1, 3), NumberCell)
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
    assert table.cell("A3").value == timedelta(days=5)
    assert table.cell("A4").value == datetime(2000, 12, 1)


NUM_ROW_COLS = 12


def test_edit_table_rows_columns(configurable_save_file):
    doc = Document(num_cols=12, num_rows=NUM_ROW_COLS, num_header_cols=0, num_header_rows=0)
    table = doc.sheets[0].tables[0]

    with pytest.raises(IndexError) as e:
        table.delete_row(start_row=NUM_ROW_COLS)
    assert "Row number not in range for table" in str(e)
    with pytest.raises(IndexError) as e:
        table.delete_row(start_row=-1)
    assert "Row number not in range for table" in str(e)
    with pytest.raises(IndexError) as e:
        table.delete_column(start_col=NUM_ROW_COLS)
    assert "Column number not in range for table" in str(e)
    with pytest.raises(IndexError) as e:
        table.delete_column(start_col=-1)
    assert "Column number not in range for table" in str(e)

    with pytest.raises(IndexError) as e:
        table.add_row(start_row=NUM_ROW_COLS)
    assert "Row number not in range for table" in str(e)
    with pytest.raises(IndexError) as e:
        table.add_row(start_row=-1)
    assert "Row number not in range for table" in str(e)
    with pytest.raises(IndexError) as e:
        table.add_column(start_col=NUM_ROW_COLS)
    assert "Column number not in range for table" in str(e)
    with pytest.raises(IndexError) as e:
        table.add_column(start_col=-1)
    assert "Column number not in range for table" in str(e)

    for row, cells in enumerate(table.iter_rows()):
        for col, _ in enumerate(cells):
            table.write(row, col, f"cell[{row},{col}]")

    table_num_rows = table.num_rows
    table_num_cols = table.num_cols

    def check_cell_cords(table: object):
        for row in range(table.num_rows):
            for col in range(table.num_cols):
                if table.cell(row, col).row != row or table.cell(row, col).col != col:
                    pass
                assert table.cell(row, col).row == row
                assert table.cell(row, col).col == col

    def add_row(
        table: object,
        table_num_rows: int,
        num_rows: int = 1,
        start_row: Optional[int] = None,
        default: object = None,
    ) -> int:
        table.add_row(num_rows, start_row, default)
        table_num_rows += num_rows
        assert table.num_rows == table_num_rows
        check_cell_cords(table)
        return table_num_rows

    def add_column(
        table: object,
        table_num_cols: int,
        num_cols: int = 1,
        start_col: Optional[int] = None,
        default: object = None,
    ) -> int:
        table.add_column(num_cols, start_col, default)
        table_num_cols += num_cols
        assert table.num_cols == table_num_cols
        check_cell_cords(table)
        return table_num_cols

    def delete_row(
        table: object, table_num_rows: int, num_rows: int = 1, start_row: Optional[int] = None,
    ) -> int:
        table.delete_row(num_rows, start_row)
        table_num_rows -= num_rows
        assert table.num_rows == table_num_rows
        check_cell_cords(table)
        return table_num_rows

    def delete_column(
        table: object, table_num_cols: int, num_cols: int = 1, start_col: Optional[int] = None,
    ) -> int:
        table.delete_column(num_cols, start_col)
        table_num_cols -= num_cols
        assert table.num_cols == table_num_cols
        check_cell_cords(table)
        return table_num_cols

    table_num_rows = add_row(table, table_num_rows, start_row=1, num_rows=3, default="new_row")
    table_num_cols = add_column(table, table_num_cols, start_col=1, num_cols=3, default="new_col")
    table_num_rows = add_row(table, table_num_rows, start_row=10, num_rows=2)
    table_num_cols = add_column(table, table_num_cols, start_col=10, num_cols=2)
    table_num_rows = delete_row(table, table_num_rows, start_row=6, num_rows=2)
    table_num_cols = delete_column(table, table_num_cols, start_col=6, num_cols=2)
    table_num_rows = delete_row(table, table_num_rows, num_rows=2)
    table_num_cols = delete_column(table, table_num_cols, num_cols=2)
    table_num_rows = add_row(table, table_num_rows, default=1.234)
    table_num_cols = add_column(table, table_num_cols, default=5.678)

    doc.save(configurable_save_file)

    ref_values = []
    for row in range(NUM_ROW_COLS):
        ref_values.append([])
        for col in range(NUM_ROW_COLS):
            ref_values[row].append(f"cell[{row},{col}]")

    ref_values[1:1] = [["new_row" for _ in range(NUM_ROW_COLS)] for _ in range(3)]
    for row, _ in enumerate(ref_values):
        ref_values[row][1:1] = ["new_col"] * 3
    ref_values[10:10] = [["" for _ in range(NUM_ROW_COLS + 3)] for _ in range(2)]
    for row, _ in enumerate(ref_values):
        ref_values[row][10:10] = ["", ""]
    del ref_values[6:8]
    del ref_values[-2:]
    for row, _ in enumerate(ref_values):
        del ref_values[row][6:8]
        del ref_values[row][-2:]
    ref_values.append(["1.234" for _ in ref_values[0]])
    for row, _ in enumerate(ref_values):
        ref_values[row].append("5.678")

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    for row, cells in enumerate(table.iter_rows()):
        for col, _ in enumerate(cells):
            assert table.cell(row, col).formatted_value == ref_values[row][col]
