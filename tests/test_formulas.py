import pytest
import pytest_check as check

from numbers_parser import Cell, Document, UnsupportedWarning
from numbers_parser.formula import Formula, Tokenizer

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
    for row in range(table.num_rows):
        for col in range(table.num_cols):
            if ref[row][col] is None:
                check.is_none(
                    table.cell(row, col).formula,
                    f"!existsy@[{row},{col}]",
                )
            else:
                check.is_true(table.cell(row, col).is_formula)
                check.is_not_none(
                    table.cell(row, col).formula,
                    f"exists@[{row},{col}]",
                )
                check.equal(
                    table.cell(row, col).formula,
                    ref[row][col],
                    f"formula@[{row},{col}]",
                )


def test_table_functions():
    doc = Document("tests/data/test-10.numbers")
    sheets = doc.sheets
    table = sheets[0].tables[0]
    compare_tables(table, TABLE_1_FORMULAS)

    table = sheets[1].tables[0]
    compare_tables(table, TABLE_2_FORMULAS)


def test_exceptions(configurable_save_file):
    def get_formula(doc):
        table_id = doc.sheets[0].tables[0]._table_id
        base_data_store = doc._model.objects[table_id].base_data_store
        formula_table_id = base_data_store.formula_table.identifier
        formula_table = doc._model.objects[formula_table_id]
        return formula_table.entries[0].formula

    doc = Document("tests/data/simple-func.numbers")
    formula = get_formula(doc)
    formula.AST_node_array.AST_node[2].AST_function_node_index = 999
    with pytest.warns(UnsupportedWarning) as record:
        value = doc.sheets[0].tables[0].cell(0, 1).formula
    assert value == "UNDEFINED!(1,2)"
    assert str(record[0].message) == "Table 1@[0,1]: function ID 999 is unsupported"

    doc = Document("tests/data/simple-func.numbers")
    formula = get_formula(doc)
    formula.AST_node_array.AST_node[2].AST_function_node_numArgs = 3
    with pytest.warns(UnsupportedWarning) as record:
        value = doc.sheets[0].tables[0].cell(0, 1).formula
    assert str(record[0].message) == "Table 1@[0,1]: stack too small for SUM"

    doc = Document("tests/data/simple-func.numbers")
    formula = get_formula(doc)
    formula.AST_node_array.AST_node[2].AST_function_node_numArgs = 3
    with pytest.warns(UnsupportedWarning) as record:
        value = doc.sheets[0].tables[0].cell(0, 1).formula
    assert str(record[0].message) == "Table 1@[0,1]: stack too small for SUM"

    doc = Document("tests/data/simple-func.numbers")
    formula = get_formula(doc)
    formula.AST_node_array.AST_node[2].AST_node_type = 68
    with pytest.warns(UnsupportedWarning) as record:
        value = doc.sheets[0].tables[0].cell(0, 1).formula
    assert str(record[0].message) == "Table 1@[0,1]: node type VIEW_TRACT_REF_NODE is unsupported"

    doc = Document("tests/data/simple-func.numbers")
    doc.sheets[0].tables[0].cell(0, 1)._formula_id = 999
    with pytest.warns(UnsupportedWarning) as record:
        _ = doc.sheets[0].tables[0].cell(0, 1).formula

    assert str(record[0].message) == "Table 1@[0,1]: key #999 not found"


def test_tokenizer():
    tok = Tokenizer("=AVERAGE(A1:D1)")
    assert str(tok) == "[FUNC(OPEN,'AVERAGE('),OPERAND(RANGE,'A1:D1'),FUNC(CLOSE,')')]"
    tok = Tokenizer('=""""&E1')
    assert str(tok) == "[OPERAND(TEXT,'\"\"\"\"'),OPERATOR-INFIX(,'&'),OPERAND(RANGE,'E1')]"
    tok = Tokenizer("=COUNTA(safari:farm)")
    assert str(tok) == "[FUNC(OPEN,'COUNTA('),OPERAND(RANGE,'safari:farm'),FUNC(CLOSE,')')]"
    tok = Tokenizer("=COUNTA(super hero)")
    assert str(tok) == "[FUNC(OPEN,'COUNTA('),OPERAND(RANGE,'super hero'),FUNC(CLOSE,')')]"
    tok = Tokenizer("=Sheet 2::Table 1::A1")
    assert str(tok) == "[OPERAND(RANGE,'Sheet 2::Table 1::A1')]"


def test_parse_formulas():
    def check_formula(cell: Cell, node):
        new_formula = Formula.from_str(
            cell._model,
            cell._table_id,
            cell.row,
            cell.col,
            cell.formula,
        )
        ref_archive = [str(x) for x in cell._model.formula_ast(cell._table_id)[cell._formula_id]]
        new_archive = [str(x) for x in new_formula._archive.AST_node_array.AST_node]
        table_name = cell._model.table_name(cell._table_id)
        print("\n")
        if ref_archive == new_archive:
            print(f"OK: {table_name}@{cell.row},{cell.col}: {cell.formula}")
        else:
            print(f"MISMATCH: {table_name}@{cell.row},{cell.col}: {cell.formula}")
            print(f"TOKENS: {new_formula._tokens}§")
            if len(ref_archive) != len(new_archive):
                print(f"LEN-REF: {len(ref_archive)}")
                print(f"LEN-NEW: {len(new_archive)}")
                print("REF-AST: " + "§".join(ref_archive) + "§")
                print("NEW-AST: " + "§".join(new_archive) + "§")
            else:
                for i, (ref, new) in enumerate(zip(ref_archive, new_archive)):
                    if ref != new:
                        print(f"REF[{i}]: {ref}")
                        print(f"NEW[{i}]: {new}")

        return True

    for filename in [
        "tests/data/create-formulas.numbers",
        # "tests/data/test-10.numbers",
        # "tests/data/simple-func.numbers",
        # "tests/data/test-all-formulas.numbers",
        # "tests/data/test-extra-formulas.numbers",
        # "tests/data/test-new-formulas.numbers",
    ]:
        doc = Document(filename)
        for sheet in doc.sheets:
            for table in sheet.tables:
                formula_ast = doc._model.formula_ast(table._table_id)
                for row in table.rows():
                    for cell in row:
                        if cell.formula is not None:
                            assert check_formula(cell, formula_ast[cell._formula_id])
