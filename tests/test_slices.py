import pytest

from numbers_parser import Document, NumberCell


def test_exceptions():
    doc = Document("tests/data/test-7.numbers")
    sheets = doc.sheets()
    tables = sheets["ZZZ_Sheet_1"].tables()
    table = tables["XXX_Table_1"]
    with pytest.raises(IndexError) as e:
        _ = table.cell("XYZ")
    assert "invalid cell" in str(e.value)
    with pytest.raises(IndexError) as e:
        _ = table.cell(4, 0)
    assert "out of range" in str(e.value)
    with pytest.raises(IndexError) as e:
        _ = table.cell(0, 5)
    assert "out of range" in str(e.value)
    with pytest.raises(IndexError) as e:
        _ = table.cell(0, 5, 2)
    assert "invalid cell reference" in str(e.value)


def test_cell_lookup():
    doc = Document("tests/data/test-7.numbers")
    sheets = doc.sheets()
    tables = sheets["ZZZ_Sheet_1"].tables()
    table = tables["XXX_Table_1"]
    assert table.num_cols == 5
    assert table.num_rows == 4
    assert table.cell(0, 1).value == "XXX_COL_2"
    assert table.cell(2, 2).value is None
    assert table.cell(3, 4).value == "XXX_3_5"


def test_cell_ref_lookup():
    doc = Document("tests/data/test-7.numbers")
    sheets = doc.sheets()
    tables = sheets["ZZZ_Sheet_1"].tables()
    table = tables["XXX_Table_1"]
    assert table.cell("A1").value == "XXX_COL_1"
    assert table.cell("C3").value is None
    assert table.cell("E4").value == "XXX_3_5"
    with pytest.raises(IndexError) as e:
        _ = table.cell("E5")
    assert "out of range" in str(e.value)
    with pytest.raises(IndexError) as e:
        _ = table.cell("A5")
    assert "out of range" in str(e.value)


def test_cell_wide_ref():
    doc = Document("tests/data/test-7.numbers")
    sheets = doc.sheets()
    tables = sheets["ZZZ_Sheet_2"].tables()
    table = tables["XXX_Table_1"]
    assert table.cell("").value == "excepteur" # defaults to [0, 0]
    assert table.cell("A1").value == "excepteur"
    assert table.cell("Z1").value == "ea"
    assert table.cell("BA1").value == "veniam"
    assert table.cell("CZ1").value == "dolore"


def test_row_iterator():
    doc = Document("tests/data/test-7.numbers")
    sheets = doc.sheets()
    tables = sheets["ZZZ_Sheet_1"].tables()
    table = tables["XXX_Table_2"]
    val = 0
    for row in table.iterrows():
        val += row[0].value if row[0] is not None else 0.0
    assert val == 252
    val = 0.0
    for row in table.iterrows(min_row=2, max_row=7):
        if isinstance(row[2], NumberCell):
            val += row[2].value if row[2] is not None else 0.0
    assert val == 672.0


def test_col_iterator():
    doc = Document("tests/data/test-7.numbers")
    sheets = doc.sheets()
    tables = sheets["ZZZ_Sheet_1"].tables()
    table = tables["XXX_Table_2"]
    val = 0
    for col in table.itercols():
        val += col[2].value if col[2] is not None else 0.0
    assert val == 164.224
    val = 0.0
    for col in table.itercols(min_col=1, max_col=3):
        val += col[8].value if col[8] is not None else 0.0
    assert val == 336.072
