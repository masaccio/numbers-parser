import pytest
import pytest_check as check

from numbers_parser import Cell, Document, UnsupportedWarning
from numbers_parser.formula import Formula
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.tokenizer import Tokenizer, TokenizerError

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
                check.is_true(table.cell(row, col).is_formula, f"formula@[{row},{col}]")
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
    test_cases = {
        "AVERAGE(A1:D1)": "[FUNC(OPEN,'AVERAGE('),OPERAND(RANGE,'A1:D1'),FUNC(CLOSE,')')]",
        '""""&E1': "[OPERAND(TEXT,'\"\"\"\"'),OPERATOR-INFIX(,'&'),OPERAND(RANGE,'E1')]",
        "COUNTA(safari:farm)": "[FUNC(OPEN,'COUNTA('),OPERAND(RANGE,'safari:farm'),FUNC(CLOSE,')')]",
        "COUNTA(super hero)": "[FUNC(OPEN,'COUNTA('),OPERAND(RANGE,'super hero'),FUNC(CLOSE,')')]",
        "Sheet 2::Table 1::A1": "[OPERAND(RANGE,'Sheet 2::Table 1::A1')]",
        "COUNTA('hyphen-name')": "[FUNC(OPEN,'COUNTA('),OPERAND(RANGE,''hyphen-name''),FUNC(CLOSE,')')]",
        "COUNTA('10%':'20%')": "[FUNC(OPEN,'COUNTA('),OPERAND(RANGE,''10%':'20%''),FUNC(CLOSE,')')]",
        "SUM(#REF!)": "[FUNC(OPEN,'SUM('),OPERAND(ERROR,'#REF!'),FUNC(CLOSE,')')]",
        "(1)": "[PAREN(OPEN,'('),OPERAND(NUMBER,'1'),PAREN(CLOSE,')')]",
        "-A1": "[OPERATOR-PREFIX(,'-'),OPERAND(RANGE,'A1')]",
        "A1%": "[OPERAND(RANGE,'A1'),OPERATOR-POSTFIX(,'%')]",
    }

    for formula, expected in test_cases.items():
        tok = Tokenizer(formula)
        assert str(tok) == expected, formula

    for formula, ref_tokens in TOKEN_TESTS.items():
        tok = Tokenizer(formula)
        tokens = Formula.rpn_tokens(tok.items)
        assert [str(x) for x in tokens] == ref_tokens, formula

    with pytest.raises(
        TokenizerError,
        match=r"Reached end of formula while parsing link in LEFT\('",
    ):
        _ = Tokenizer("LEFT('")


def check_generated_formula(cell: Cell) -> bool:
    new_formula_id = Formula.from_str(
        cell._model,
        cell._table_id,
        cell.row,
        cell.col,
        cell.formula,
    )

    ref_archive = cell._model._formulas.lookup_value(cell._table_id, cell._formula_id)
    new_archive = cell._model._formulas.lookup_value(cell._table_id, new_formula_id)

    if ref_archive != new_archive:
        node_name_map = {
            k: v.name
            for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()
        }

        table_name = cell._model.table_name(cell._table_id)
        ref_node_types = [
            node_name_map[x.AST_node_type] for x in ref_archive.formula.AST_node_array.AST_node
        ]
        new_node_types = [
            node_name_map[x.AST_node_type] for x in new_archive.formula.AST_node_array.AST_node
        ]

        print("\n")
        print(f"MISMATCH: {table_name}@{cell.row},{cell.col}: {cell.formula}")
        print("TOKENS:", str(Formula.formula_tokens(cell.formula)))

        max_len = max([len(x) for x in ref_node_types + new_node_types]) + 2
        num_dashes = int(max_len / 2) - 2
        ref_header = "-" * num_dashes + " REF " + "-" * num_dashes
        new_header = "-" * num_dashes + " NEW " + "-" * num_dashes

        if len(ref_node_types) != len(new_node_types):
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
            for i, (ref, new) in enumerate(
                zip(
                    ref_archive.formula.AST_node_array.AST_node,
                    new_archive.formula.AST_node_array.AST_node,
                ),
            ):
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
                assert row[0].formula == row[1].value, f"{table_name}: row {row_num + 1}"
                _ = check_generated_formula(row[0]), f"{table_name}: row {row_num + 1}"


@pytest.mark.experimental
def test_create_formula(configurable_save_file):
    def copy_table_data(src, dest):
        for table in src.tables:
            if table.name not in dest.tables:
                new_table = dest.add_table(
                    table_name=table.name,
                    num_rows=table.num_rows,
                    num_cols=table.num_cols,
                    num_header_rows=table.num_header_rows,
                    num_header_cols=table.num_header_cols,
                )
            else:
                new_table = dest.tables[table.name]

            for row_num, row in enumerate(table.iter_rows()):
                for col_num, cell in enumerate(row):
                    if cell.value is not None:
                        new_table.write(row_num, col_num, cell.value)

    def copy_table_formulas(src, dest):
        for table in src.tables:
            for row_num, row in enumerate(table.iter_rows()):
                for col_num, cell in enumerate(row):
                    if cell.formula is not None:
                        new_table = dest.tables[table.name]
                        new_table.cell(row_num, col_num).formula = cell.formula

    ref_doc = doc = Document("tests/data/create-formulas.numbers")
    doc = Document(
        sheet_name="Main Sheet",
        table_name="Formula Tests",
        num_header_cols=0,
        num_cols=3,
    )
    doc.add_sheet("Powers Sheet", table_name="Powers of Two", num_rows=5, num_cols=3)
    copy_table_data(ref_doc.sheets["Main Sheet"], doc.sheets["Main Sheet"])
    copy_table_data(ref_doc.sheets["Powers Sheet"], doc.sheets["Powers Sheet"])
    # Copy formulas only after all data is copied as rows and column references
    # will otherwise fail
    copy_table_formulas(ref_doc.sheets["Main Sheet"], doc.sheets["Main Sheet"])
    copy_table_formulas(ref_doc.sheets["Powers Sheet"], doc.sheets["Powers Sheet"])

    doc.save(configurable_save_file)
