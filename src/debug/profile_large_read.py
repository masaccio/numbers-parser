from cProfile import Profile
from pstats import Stats
from tempfile import TemporaryDirectory

from numbers_parser import Document
from numbers_parser.constants import MAX_COL_COUNT

doc = Document()
for row in range(1000):
    for col in range(MAX_COL_COUNT):
        doc.sheets[0].tables[0].write(row, col, f"row={row}, col={col}")

with TemporaryDirectory() as tmpdirname:
    filename = f"{tmpdirname}/large.numbers"
    doc.save(filename)

    with Profile() as pr:
        Document(filename)

pr.create_stats()
stats = Stats(pr)
stats = stats.strip_dirs().sort_stats("cumtime")
stats.print_stats()
