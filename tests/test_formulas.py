import pytest
import pytest_check as check

from numbers_parser import Document, ErrorCell

TABLE_1_FORMULAS = [
    [None, "A1", "$B$1=1"],
    [None, "A1+A2", "A$2&B2"],
    [None, "A1×A2", "NOW()"],
    [None, "A1-A2", "NOW()+0.1"],
    [None, "A1÷A2", "$C4-C3"],
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
    [None, 'IF(FIND("_",A8)>2,A1,A2)'],
    [None, "100×(A9×2)%"],
    [None, 'IF(A10<5,"smaller","larger")'],
    [None, 'IF(A11≤5,"smaller","larger")'],
]


def compare_tables(table, ref):
    for row_num in range(table.num_rows):
        for col_num in range(table.num_cols):
            if ref[row_num][col_num] is None:
                check.is_none(
                    table.cell(row_num, col_num).formula,
                    f"!existsy@[{row_num},{col_num}]",
                )
            else:
                check.is_not_none(
                    table.cell(row_num, col_num).formula,
                    f"exists@[{row_num},{col_num}]",
                )
                check.equal(
                    table.cell(row_num, col_num).formula,
                    ref[row_num][col_num],
                    f"formula@[{row_num},{col_num}]",
                )


def test_table_functions():
    doc = Document("tests/data/test-10.numbers")
    sheets = doc.sheets
    table = sheets[0].tables[0]
    compare_tables(table, TABLE_1_FORMULAS)

    table = sheets[1].tables[0]
    compare_tables(table, TABLE_2_FORMULAS)
