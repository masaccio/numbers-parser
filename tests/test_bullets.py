"""Tests for bulleted text and hyperlinks."""

from pathlib import Path

from numbers_parser import Document, RichTextCell

TEST_REF_0 = "some\nnewline\ntext"
TEST_REF_1 = "this\nis some\nbulleted\ntext"
TEST_REF_2 = "this\nis some\nnumbered\ntext"
TEST_REF_3 = "this\nis some\nlettered\ntext"


def test_bullets() -> None:
    """Test bulleted text extraction."""
    doc = Document("tests/data/test-bullets.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]
    assert table.cell(0, 0).value == TEST_REF_0
    assert not table.cell(0, 0).is_bulleted
    assert table.cell(0, 1).value == TEST_REF_1
    assert table.cell(1, 1).value == TEST_REF_2
    assert table.cell(2, 1).value == TEST_REF_3

    bullets_0_1 = table.cell(0, 1).bullets
    assert bullets_0_1 is not None
    assert bullets_0_1[0] == TEST_REF_1.split("\n", maxsplit=1)[0]
    assert bullets_0_1[2] == TEST_REF_1.split("\n")[2]
    assert bullets_0_1[3] == TEST_REF_1.split("\n")[3]

    assert table.cell(2, 0).value == "t"
    assert not table.cell(0, 0).is_bulleted
    assert table.cell(1, 1).is_bulleted
    assert table.cell(0, 0).bullets is None


def test_hyperlinks(configurable_save_file: Path) -> None:
    """Test hyperlinks in cells."""

    def check_doc_links(doc: Document) -> None:
        sheets = doc.sheets
        tables = sheets[0].tables
        table = tables[0]

        bullets_2_0 = table.cell(2, 0).bullets
        assert bullets_2_0 is not None
        assert len(bullets_2_0) == 4
        assert table.cell(2, 0).is_bulleted

        bullets_3_0 = table.cell(3, 0).bullets
        assert bullets_3_0 is not None
        assert len(bullets_3_0) == 4

        cell = table.cell(0, 0)
        assert isinstance(cell, RichTextCell)
        assert cell.hyperlinks is not None
        assert len(cell.hyperlinks) == 2
        assert not cell.is_bulleted
        assert cell.hyperlinks[0] == ("BBC News", "http://news.bbc.co.uk/")
        assert cell.value == "visit BBC News for the news and Google to find stuff"

        cell_1_0 = table.cell(1, 0)
        assert isinstance(cell_1_0, RichTextCell)
        assert cell_1_0.hyperlinks is not None
        assert cell_1_0.hyperlinks[0][0] == "Google"

    doc = Document("tests/data/test-hlinks.numbers")
    check_doc_links(doc)

    doc.save(configurable_save_file)
    new_doc = Document(configurable_save_file)
    check_doc_links(new_doc)
