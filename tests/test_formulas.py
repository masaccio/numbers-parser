import pytest

from numbers_parser import Document, FormulaCell, ErrorCell

@pytest.mark.experimental
def test_table_functions():
    doc = Document("tests/data/test-10.numbers")
    sheets = doc.sheets()
    table = sheets[1].tables()[0]
    formulas = table.formulas
    import json
    print(json.dumps(table._formulas, indent=4, sort_keys=True))
    assert formulas[-1]["ast"][0]["type"] == "STRING_NODE"
    assert formulas[-1]["ast"][0]["string"] == "YYY"
    assert formulas[-1]["ast"][1]["type"] == "CELL_REFERENCE_NODE"
    assert formulas[-1]["ast"][1]["ref"] == "A7"
