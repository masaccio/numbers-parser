from numbers_parser import Document
from numbers_parser.cell import xl_range


def test_containers():
    doc = Document("tests/data/test-1.numbers")
    assert len(doc._model.objects) == 738


def test_ranges():
    assert xl_range(2, 2, 2, 2) == "C3"
