import pytest
from psutil import Process

from numbers_parser import Document

NUM_DOCS = 10


@pytest.mark.experimental
def test_mempoy_leaks():
    """Memory leak test (see issue-67)."""
    process = Process()

    rss_base = process.memory_info().rss / 1000000.0

    print("")
    print(f"LEAK TEST: Creating {NUM_DOCS} large documents")
    print(f"LEAK TEST: RSS={rss_base:.1f}MB")

    for i in range(NUM_DOCS):
        doc = Document("tests/data/issue-67.numbers")
        assert doc.sheets[0].tables[0].cell(0, 0).value == "A"

        rss = process.memory_info().rss / 1000000.0
        if i == int(NUM_DOCS / 2):
            pass
        print(f"LEAK TEST: iteration #{i}: RSS={rss:.1f}MB")

    rss_end = process.memory_info().rss / 1000000.0

    print(f"LEAK TEST: Memory in use: base={rss_base:.1f}MB, now={rss_end:.1f}MB")

    assert rss_end < rss_base * 5
