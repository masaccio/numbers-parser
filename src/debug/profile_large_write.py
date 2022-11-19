from cProfile import Profile
from pstats import Stats
from tempfile import TemporaryDirectory

from numbers_parser import Document
from numbers_parser.constants import MAX_COL_COUNT

with Profile() as pr:
    doc = Document()
    for row_num in range(0, 1000):
        for col_num in range(0, MAX_COL_COUNT):
            doc.sheets[0].tables[0].write(
                row_num, col_num, f"row={row_num}, col={col_num}"
            )
    with TemporaryDirectory() as tmpdirname:
        filename = f"{tmpdirname}/large.numbers"
        doc.save(filename)

pr.create_stats()
stats = Stats(pr)
stats = stats.strip_dirs().sort_stats("time")
stats.print_stats()
