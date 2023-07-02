from numbers_parser import Document
from numbers_parser.cell import xl_range


def test_containers():
    doc = Document("tests/data/test-1.numbers")
    assert len(doc._model.objects) == 738


def test_ranges():
    assert xl_range(2, 2, 2, 2) == "C3"


def test_datalists():
    doc = Document("tests/data/test-1.numbers")
    table_id = doc.sheets[0].tables[0]._table_id
    assert doc._model.table_string(table_id, 1) == "YYY_2_1"
    assert doc._model._table_strings.lookup_key(table_id, "TEST") == 19
    assert doc._model._table_strings.lookup_key(table_id, "TEST") == 19
    assert doc._model._table_strings.lookup_value(table_id, 19).string == "TEST"
