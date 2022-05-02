import pytest
import os
from tempfile import TemporaryDirectory

from numbers_parser import Document


def test_save_document():
    doc = Document("tests/data/test-1.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    cell_values = tables["ZZZ_Table_1"].rows(values_only=True)

    sheets["ZZZ_Sheet_1"].name = "ZZZ_Sheet_1 NEW"
    tables["ZZZ_Table_1"].name = "ZZZ_Table_1 NEW"

    temp_dir = TemporaryDirectory()
    new_filename = os.path.join(temp_dir.name, "test-1-new.numbers")
    doc.save(new_filename)

    new_doc = Document(new_filename)
    new_sheets = new_doc.sheets()
    new_tables = new_sheets["ZZZ_Sheet_1 NEW"].tables()
    new_cell_values = new_tables["ZZZ_Table_1 NEW"].rows(values_only=True)

    assert cell_values == new_cell_values
