from numbers_parser import Document, EmptyCell


def test_issue_255(configurable_save_file):
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.write(1, 1, "old value")
    table.write("C3", "another value")

    table.write(1, 1, None)
    table.write("C3", None)
    table.write("D4", None)

    for coordinate in ("B2", "C3", "D4"):
        cell = table.cell(coordinate)
        assert isinstance(cell, EmptyCell)
        assert cell.value is None
        assert cell.formatted_value == ""

    doc.save(configurable_save_file)
    reloaded = Document(configurable_save_file).sheets[0].tables[0]
    for coordinate in ("B2", "C3", "D4"):
        cell = reloaded.cell(coordinate)
        assert isinstance(cell, EmptyCell)
        assert cell.value is None
