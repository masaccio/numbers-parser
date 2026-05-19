import gc
from sys import version_info

from pympler import muppy, summary

from numbers_parser import Document


def test_memory_leaks():
    """Memory leak test (see issue-67)."""
    if version_info < (3, 11):
        return

    last_num_objects = None
    last_num_bytes = None
    iterations = 10
    for ii in range(iterations):
        doc = Document("tests/data/test-1.numbers")
        del doc
        gc.collect()

        interim_summary = summary.summarize(muppy.get_objects())
        num_objects = sum(r[1] for r in interim_summary)
        num_bytes = sum(r[2] for r in interim_summary)

        if ii > 1:
            assert num_objects - last_num_objects <= 0
            assert num_bytes - last_num_bytes <= 0

        last_num_objects = num_objects
        last_num_bytes = num_bytes
