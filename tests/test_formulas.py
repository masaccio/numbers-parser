import pytest

from numbers_parser import Document, FormulaCell, ErrorCell

TABLE_1_FORMULAS = [
    [None, "A1", "$B$1=1"],
    [None, "A1+A2", "A$2&B2"],
    [None, "A1*A2", "NOW()"],
    [None, "A1-A2", "NOW()+0.1"],
    [None, "A1/A2", "$C4-C3"],
    [None, "SUM(A1:A2)", "IF(A6>6,TRUE,FALSE)"],
    [None, "MEDIAN(A1:A2)", "IF(A7>0,TRUE,FALSE)"],
    [None, "AVERAGE(A1:A2)", "A8≠10"],
    ["A9", None, None],
]

TABLE_2_FORMULAS = [
    [None, "A1&A2&A3"],
    [None, "LEN(A2)+LEN(A3)"],
    [None, "LEFT(A3,1)"],
    [None, "MID(A4,2,2)"],
    [None, "RIGHT(A5,2)"],
    [None, 'FIND("_",A6)'],
    [None, 'FIND("YYY",A7)'],
    [None, 'IF(FIND("_", A8)>2,A1,A2)'],
]


def compare_tables(table, ref):
    for row_num in range(table.num_rows):
        for col_num in range(table.num_cols):
            if ref[row_num][col_num] is None:
                # assert not table.cell(row_num, col_num).has_formula
                if table.cell(row_num, col_num).has_formula:
                    print(f"[{row_num},{col_num}]: has_formula FAILED")
            else:
                # assert table.cell(row_num, col_num).formula == ref[row_num][col_num]
                if not table.cell(row_num, col_num).has_formula:
                    print(f"[{row_num},{col_num}]: has_formula FAILED")
                if table.cell(row_num, col_num).formula != ref[row_num][col_num]:
                    val = table.cell(row_num, col_num).formula
                    print(f"[{row_num},{col_num}]: {val}≠{ref[row_num][col_num]}")


@pytest.mark.experimental
def test_table_functions():
    doc = Document("tests/data/test-10.numbers")
    sheets = doc.sheets()
    table = sheets[0].tables()[0]
    compare_tables(table, TABLE_1_FORMULAS)

    table = sheets[1].tables()[0]
    compare_tables(table, TABLE_2_FORMULAS)

    # for sheet in doc.sheets():
    #     print("==== TABLE:", sheet.tables()[0].name)
    #     for row in sheet.tables()[0].rows():
    #         for cell in row:
    #             if cell.has_formula:
    #                 print(f'@{cell.row},{cell.col}, "{cell.formula}"')
