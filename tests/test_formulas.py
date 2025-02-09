import pytest
import pytest_check as check

from numbers_parser import Cell, Document, UnsupportedWarning
from numbers_parser.formula import Formula, Tokenizer
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives

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


def test_named_ranges():
    doc = Document("tests/data/create-formulas.numbers")
    table = doc.sheets["Main Sheet"].tables["Reference Tests"]
    for row_num, row in enumerate(table.iter_rows(min_row=1), start=1):
        if len(row) == 2 or row[2].value:
            assert row[0].formula == row[1].value, f"Reference Tests: row {row_num + 1}"


TOKEN_TESTS = {
    "A1": [
        "OPERAND(RANGE,'A1')",
    ],
    "10+10×20+20": [
        "OPERAND(NUMBER,'10')",
        "OPERAND(NUMBER,'10')",
        "OPERAND(NUMBER,'20')",
        "OPERATOR-INFIX(,'×')",
        "OPERATOR-INFIX(,'+')",
        "OPERAND(NUMBER,'20')",
        "OPERATOR-INFIX(,'+')",
    ],
    "((10+10)÷(20+20))%": [
        "OPERAND(NUMBER,'10')",
        "OPERAND(NUMBER,'10')",
        "OPERATOR-INFIX(,'+')",
        "OPERAND(NUMBER,'20')",
        "OPERAND(NUMBER,'20')",
        "OPERATOR-INFIX(,'+')",
        "OPERATOR-INFIX(,'÷')",
        "OPERATOR-POSTFIX(,'%')",
    ],
    "SUM(A1:B1)": [
        "OPERAND(RANGE,'A1:B1')",
        "FUNC(OPEN,'SUM')",
    ],
    'E1&" "&F1': [
        "OPERAND(RANGE,'E1')",
        "OPERAND(TEXT,'\" \"')",
        "OPERATOR-INFIX(,'&')",
        "OPERAND(RANGE,'F1')",
        "OPERATOR-INFIX(,'&')",
    ],
    '""""&E1&""" """&F1&""""': [
        'OPERAND(TEXT,\'""""\')',
        "OPERAND(RANGE,'E1')",
        "OPERATOR-INFIX(,'&')",
        'OPERAND(TEXT,\'""" """\')',
        "OPERATOR-INFIX(,'&')",
        "OPERAND(RANGE,'F1')",
        "OPERATOR-INFIX(,'&')",
        'OPERAND(TEXT,\'""""\')',
        "OPERATOR-INFIX(,'&')",
    ],
    "A1×B1": [
        "OPERAND(RANGE,'A1')",
        "OPERAND(RANGE,'B1')",
        "OPERATOR-INFIX(,'×')",
    ],
    "A1÷B1": [
        "OPERAND(RANGE,'A1')",
        "OPERAND(RANGE,'B1')",
        "OPERATOR-INFIX(,'÷')",
    ],
    "IF(1≤2,TRUE,FALSE)": [
        "OPERAND(NUMBER,'1')",
        "OPERAND(NUMBER,'2')",
        "OPERATOR-INFIX(,'≤')",
        "OPERAND(LOGICAL,'TRUE')",
        "OPERAND(LOGICAL,'FALSE')",
        "FUNC(OPEN,'IF')",
    ],
    "IF(1≥2,TRUE,FALSE)": [
        "OPERAND(NUMBER,'1')",
        "OPERAND(NUMBER,'2')",
        "OPERATOR-INFIX(,'≥')",
        "OPERAND(LOGICAL,'TRUE')",
        "OPERAND(LOGICAL,'FALSE')",
        "FUNC(OPEN,'IF')",
    ],
    "IF(1≠2,TRUE,FALSE)": [
        "OPERAND(NUMBER,'1')",
        "OPERAND(NUMBER,'2')",
        "OPERATOR-INFIX(,'≠')",
        "OPERAND(LOGICAL,'TRUE')",
        "OPERAND(LOGICAL,'FALSE')",
        "FUNC(OPEN,'IF')",
    ],
    "A1+B1-C1": [
        "OPERAND(RANGE,'A1')",
        "OPERAND(RANGE,'B1')",
        "OPERATOR-INFIX(,'+')",
        "OPERAND(RANGE,'C1')",
        "OPERATOR-INFIX(,'-')",
    ],
    "A1+SUM(A1:B1)-SUM(A2:B2)": [
        "OPERAND(RANGE,'A1')",
        "OPERAND(RANGE,'A1:B1')",
        "FUNC(OPEN,'SUM')",
        "OPERATOR-INFIX(,'+')",
        "OPERAND(RANGE,'A2:B2')",
        "FUNC(OPEN,'SUM')",
        "OPERATOR-INFIX(,'-')",
    ],
    "IF(B1>0,COUNTA(E1:F1),COUNTA(E1:H1))": [
        "OPERAND(RANGE,'B1')",
        "OPERAND(NUMBER,'0')",
        "OPERATOR-INFIX(,'>')",
        "OPERAND(RANGE,'E1:F1')",
        "FUNC(OPEN,'COUNTA')",
        "OPERAND(RANGE,'E1:H1')",
        "FUNC(OPEN,'COUNTA')",
        "FUNC(OPEN,'IF')",
    ],
    'IF(OR(SUM(A2:B2)>1,SUM(A1:B1)>1),"yay","nay")': [
        "OPERAND(RANGE,'A2:B2')",
        "FUNC(OPEN,'SUM')",
        "OPERAND(NUMBER,'1')",
        "OPERATOR-INFIX(,'>')",
        "OPERAND(RANGE,'A1:B1')",
        "FUNC(OPEN,'SUM')",
        "OPERAND(NUMBER,'1')",
        "OPERATOR-INFIX(,'>')",
        "FUNC(OPEN,'OR')",
        "OPERAND(TEXT,'\"yay\"')",
        "OPERAND(TEXT,'\"nay\"')",
        "FUNC(OPEN,'IF')",
    ],
    "IF(TRUE,1,0)": [
        "OPERAND(LOGICAL,'TRUE')",
        "OPERAND(NUMBER,'1')",
        "OPERAND(NUMBER,'0')",
        "FUNC(OPEN,'IF')",
    ],
    "POWER(2,5)": [
        "OPERAND(NUMBER,'2')",
        "OPERAND(NUMBER,'5')",
        "FUNC(OPEN,'POWER')",
    ],
}


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

    for formula, ref_tokens in TOKEN_TESTS.items():
        tok = Tokenizer("=" + formula)
        tokens = Formula.rpn_tokens(tok.items)
        assert [str(x) for x in tokens] == ref_tokens, formula


def check_generated_formula(cell: Cell) -> bool:
    new_formula = Formula.from_str(
        cell._model,
        cell._table_id,
        cell.row,
        cell.col,
        cell.formula,
    )

    ref_archive = [str(x) for x in cell._model.formula_ast(cell._table_id)[cell._formula_id]]
    new_archive = [str(x) for x in new_formula._archive.AST_node_array.AST_node]

    if ref_archive != new_archive:
        node_name_map = {
            k: v.name
            for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()
        }

        table_name = cell._model.table_name(cell._table_id)
        ref_node_types = [
            node_name_map[x.AST_node_type]
            for x in cell._model.formula_ast(cell._table_id)[cell._formula_id]
        ]
        new_node_types = [
            node_name_map[x.AST_node_type] for x in new_formula._archive.AST_node_array.AST_node
        ]

        print("\n")
        print(f"MISMATCH: {table_name}@{cell.row},{cell.col}: {cell.formula}")
        print(f"TOKENS: {new_formula._tokens}")

        max_len = max([len(x) for x in ref_node_types + new_node_types]) + 2
        num_dashes = int(max_len / 2) - 2
        ref_header = "-" * num_dashes + " REF " + "-" * num_dashes
        new_header = "-" * num_dashes + " NEW " + "-" * num_dashes

        if len(ref_archive) != len(new_archive):
            print(f"{ref_header} | {new_header}")
            for i in range(max([len(ref_node_types), len(new_node_types)])):
                ref = ref_node_types[i] if i < len(ref_node_types) else ""
                new = new_node_types[i] if i < len(new_node_types) else ""
                print(f"{ref:{max_len}s} | {new:{max_len}s}")
        elif ref_node_types != new_node_types:
            print(f"{ref_header} | {new_header}")
            for i in range(len(ref_node_types)):
                print(f"{ref_node_types[i]:{max_len}s} | {new_node_types[i]:{max_len}s}")
        else:
            for i, (ref, new) in enumerate(zip(ref_archive, new_archive)):
                if ref != new:
                    print(f"REF[{i}]: {ref}")
                    print(f"NEW[{i}]: {new}")
        return False

    return True


def test_parse_formulas():
    doc = Document("tests/data/create-formulas.numbers")

    cell = doc.sheets["Main Sheet"].tables["Formula Tests"].cell(1, 0)
    with pytest.warns(UnsupportedWarning) as record:
        _ = Formula.from_str(cell._model, cell._table_id, cell.row, cell.col, "=XXX()")
    assert record[0].message.args[0] == "Formula Tests@[1,0]: function XXX is not supported."

    sheet = doc.sheets["Main Sheet"]
    for table_name in ["Formula Tests", "Reference Tests"]:
        table = sheet.tables[table_name]
        for row_num, row in enumerate(table.iter_rows(min_row=1), start=1):
            if len(row) == 2 or row[2].value:
                pass
                # assert row[0].formula == row[1].value, f"{table_name}: row {row_num + 1}"
                # assert check_generated_formula(row[0]), f"{table_name}: row {row_num + 1}"
