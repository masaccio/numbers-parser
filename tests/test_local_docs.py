import os
import traceback
import warnings

import pytest

from numbers_parser import Document
from numbers_parser.constants import DEFAULT_FONT

ALL_DOCS_DIRS = [
    os.path.join(os.environ.get("HOME"), x)
    for x in [
        "Documents",
        "Library/Mobile Documents/com~apple~Numbers/Documents",
    ]
]


@pytest.mark.experimental
def test_local_docs():
    def all_doc_cells(doc: Document):
        for sheet in doc.sheets:
            for table in sheet.tables:
                for row in table.rows():
                    yield from row

    def all_docs_paths():
        for docs_dir in ALL_DOCS_DIRS:
            for root, dirs, files in os.walk(docs_dir):
                for file in files + dirs:
                    if not file.endswith(".numbers"):
                        continue
                    path = os.path.join(root, file)
                    yield path

    def check_cell_values(doc: Document):
        for cell in all_doc_cells(doc):
            _ = cell.value

    def check_cell_formatted_values(doc: Document):
        for cell in all_doc_cells(doc):
            _ = cell.formatted_value

    def check_cell_formulas(doc: Document):
        for cell in all_doc_cells(doc):
            if cell.is_formula:
                _ = cell.formula

    def check_cell_style(doc: Document):
        for cell in all_doc_cells(doc):
            _ = cell.style

    def check_cell_borders(doc: Document):
        for cell in all_doc_cells(doc):
            _ = cell.border

    def allowed_warning(w):
        return f"falling back to {DEFAULT_FONT}" in str(w.message)

    tests = [
        check_cell_values,
        check_cell_formatted_values,
        check_cell_formulas,
        check_cell_style,
        check_cell_borders,
    ]

    print("\n*** Testing local docs")
    for doc_path in all_docs_paths():
        test_ok_count = 0
        captured_warnings = []
        exc = None
        try:
            with warnings.catch_warnings(record=True) as captured_warnings:
                warnings.simplefilter("always")
                print(doc_path)
                doc = Document(doc_path)
                for test in tests:
                    test(doc)
                    test_ok_count += 1
        except Exception as e:
            exc = e
        else:
            captured_warnings = [w for w in captured_warnings if not allowed_warning(w)]

        assert_msg = ""
        if len(captured_warnings) > 1:
            assert_msg += "\n".join(
                [f"\tWarning: {w.category.__name__}: {w.message}" for w in captured_warnings],
            )
        if exc is not None:
            exc_type = type(exc).__name__
            exc_message = str(exc)
            extracted_tb = traceback.extract_tb(exc.__traceback__)
            last_frame = extracted_tb[-1]
            filename = last_frame.filename
            line_number = last_frame.lineno
            test_name = tests[test_ok_count].__name__
            assert_msg += f"{doc_path}: FAILED with {exc_type} exception"
            assert_msg += f"\tFailing test: {test_name}"
            assert_msg += f"\tMessage: {exc_message}"
            assert_msg += f"\tLocation: File '{filename}', line {line_number}"

        assert assert_msg == "", assert_msg
