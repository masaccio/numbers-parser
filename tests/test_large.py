from numbers_parser import Document


def test_many_rows():
    doc = Document("tests/data/test-3.numbers")
    sheets = doc.sheets
    tables = sheets["Sheet_1"].tables
    data = tables[0].rows(values_only=True)
    ref = []
    _ = [ref.append(["ROW" + str(x)]) for x in range(1, 601)]
    assert data == ref


def test_many_columns():
    doc = Document("tests/data/test-3.numbers")
    sheets = doc.sheets
    tables = sheets["Sheet_2"].tables
    data = tables[0].rows(values_only=True)
    ref = [["COLUMN" + str(x) for x in range(1, 601)]]
    assert data == ref


def ref_cell_text(row, col):
    return "CELL [" + str(row + 1) + "," + str(col + 1) + "]"


def test_large_table():
    row = [None for i in range(270)]
    ref = [row.copy() for i in range(90)]
    for i in range(90):
        ref[i][i] = ref_cell_text(i, i)
        ref[i][90 + i] = ref_cell_text(i, 90 + i)
        ref[i][180 + i] = ref_cell_text(i, 180 + i)

    doc = Document("tests/data/test-6.numbers")
    sheets = doc.sheets
    tables = sheets["Sheet"].tables
    data = tables[0].rows(values_only=True)
    assert data == ref
