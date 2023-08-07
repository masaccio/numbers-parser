from numbers_parser import Document

TEST_REF_0 = "some\nnewline\ntext"
TEST_REF_1 = "this\nis some\nbulleted\ntext"
TEST_REF_2 = "this\nis some\nnumbered\ntext"
TEST_REF_3 = "this\nis some\nlettered\ntext"


def test_bullets():
    doc = Document("tests/data/test-bullets.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]
    assert table.cell(0, 0).value == TEST_REF_0
    assert not table.cell(0, 0).is_bulleted
    assert table.cell(0, 1).value == TEST_REF_1
    assert table.cell(1, 1).value == TEST_REF_2
    assert table.cell(2, 1).value == TEST_REF_3
    assert table.cell(0, 1).bullets[0] == TEST_REF_1.split("\n")[0]
    assert table.cell(0, 1).bullets[2] == TEST_REF_1.split("\n")[2]
    assert table.cell(0, 1).bullets[3] == TEST_REF_1.split("\n")[3]
    assert table.cell(2, 0).value == "t"
    assert not table.cell(0, 0).is_bulleted
    assert table.cell(1, 1).is_bulleted
    assert table.cell(0, 0).bullets is None


def test_hyperlinks(configurable_save_file):
    def check_doc_links(doc):
        sheets = doc.sheets
        tables = sheets[0].tables
        table = tables[0]

        assert len(table.cell(2, 0).bullets) == 4
        assert table.cell(2, 0).is_bulleted
        assert len(table.cell(3, 0).bullets) == 4

        cell = table.cell(0, 0)
        assert len(cell.hyperlinks) == 2
        assert not cell.is_bulleted
        assert cell.hyperlinks[0] == ("BBC News", "http://news.bbc.co.uk/")
        assert cell.value == "visit BBC News for the news and Google to find stuff"

        assert table.cell(1, 0).hyperlinks[0][0] == "Google"

    doc = Document("tests/data/test-hlinks.numbers")
    check_doc_links(doc)

    doc.save(configurable_save_file)
    new_doc = Document(configurable_save_file)
    check_doc_links(new_doc)
