from numbers_parser import Document, MergedCell
from numbers_parser.cell import Border

doc = Document("tests/data/test-styles.numbers")
table = doc.sheets["Large Borders"].tables[0]
strokes = doc._model.extract_strokes(table._table_id)


def check_border(border, test_value):
    values = test_value.split(",")
    values[0] = float(values[0])
    values[1] = eval(values[1].replace(";", ","))
    if border is None:
        return False
    return border == Border(values[0], values[1], values[2])


for row_num, row in enumerate(table.iter_rows()):
    for col_num, cell in enumerate(row):
        if isinstance(cell, MergedCell):
            continue

        tests = cell.value.split("\n")
        if all(
            [
                check_border(cell._border.top, tests[0][2:]),
                check_border(cell._border.left, tests[1][2:]),
                check_border(cell._border.right, tests[2][2:]),
                check_border(cell._border.bottom, tests[3][2:]),
            ]
        ):
            print(f"@[{row_num},{col_num}]: OK")
        else:
            print(f"@[{row_num},{col_num}]: FAIL")
            print("   reference :", ", ".join(tests))
            print("   top       :", cell._border.top)
            print("   left      :", cell._border.left)
            print("   right     :", cell._border.right)
            print("   bottom    :", cell._border.bottom)
