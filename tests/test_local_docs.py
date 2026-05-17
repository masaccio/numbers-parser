import os
import traceback
import warnings

import pytest

from numbers_parser import Document, FileFormatError

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
            _ = cell.formula

    def check_cell_style(doc: Document):
        for cell in all_doc_cells(doc):
            _ = cell.style

    def check_cell_borders(doc: Document):
        for cell in all_doc_cells(doc):
            _ = cell.border

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
        try:
            with warnings.catch_warnings(record=True) as captured_warnings:
                warnings.simplefilter("always")
                doc = Document(doc_path)
                for test in tests:
                    test(doc)
                    test_ok_count += 1
        except FileFormatError:
            print(f"{doc_path}: invalid Numbers document")
        except Exception as e:
            exc_type = type(e).__name__
            exc_message = str(e)
            extracted_tb = traceback.extract_tb(e.__traceback__)
            last_frame = extracted_tb[-1]
            filename = last_frame.filename
            line_number = last_frame.lineno
            test_name = tests[test_ok_count].__name__
            print(f"{doc_path}: FAILED with {exc_type} exception")
            print(f"\tFailing test: {test_name}")
            print(f"\tMessage:  {exc_message}")
            print(f"\tLocation: File '{filename}', line {line_number}")
        else:
            if len(captured_warnings) > 0:
                print(f"{doc_path}: OK with {len(captured_warnings)} warnings")
            else:
                print(f"{doc_path}: OK")
