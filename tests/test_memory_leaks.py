import gc

import pytest
from pympler import muppy, summary

from numbers_parser import Document


@pytest.mark.experimental
def test_memory_leaks():
    """Memory leak test (see issue-67)."""
    sum1 = summary.summarize(muppy.get_objects())

    doc = Document("tests/data/issue-67.numbers")
    del doc
    gc.collect()

    sum2 = summary.summarize(muppy.get_objects())
    diff = summary.get_diff(sum1, sum2)

    summary.print_(diff)

    total_size = 0
    for func in filter(lambda x: x[1] > 0, diff):
        total_size += func[2]

    total_size /= 1024 * 1024

    assert total_size < 2.0
