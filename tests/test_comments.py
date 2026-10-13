from datetime import datetime

import pytest

from numbers_parser import Comment, Document


def test_comments():
    doc = Document("tests/data/test-comments.numbers")
    table = doc.sheets[0].tables[0]

    comments = []
    for row in table.rows():
        for cell in row:
            if cell.comments:
                comments.extend(cell.comments)

    # Sort by root comment text to ensure deterministic assertion order
    comments.sort(key=lambda c: c.text)

    assert len(comments) == 4

    c1 = comments[0]
    assert c1.text == "Comment #1\n"
    assert c1.author == "Jon Connell"
    assert isinstance(c1.creation_date, datetime)
    assert len(c1.replies) == 2

    assert c1.replies[0].text == "Reply to comment #1"
    assert c1.replies[0].author == "Jon Connell"
    assert isinstance(c1.replies[0].creation_date, datetime)

    assert c1.replies[1].text == "Second reply"
    assert c1.replies[1].author == "Jon Connell"
    assert isinstance(c1.replies[1].creation_date, datetime)

    c2 = comments[1]
    assert c2.text == "Comment #2"
    assert c2.author == "Jon Connell"
    assert isinstance(c2.creation_date, datetime)
    assert len(c2.replies) == 0

    c3 = comments[2]
    assert c3.text == "Comment #3"
    assert c3.author == "Jon Connell"
    assert isinstance(c3.creation_date, datetime)
    assert len(c3.replies) == 2

    c4 = comments[3]
    assert c4.text == "Comment on merged cells"


def test_passthrough_comments(configurable_save_file):
    doc = Document("tests/data/test-comments.numbers")
    doc.save(configurable_save_file)

    new_doc = Document(configurable_save_file)
    comments = new_doc.default_table.cell(0, 0).comments
    assert isinstance(comments[0], Comment)
    assert comments[0].author == "Jon Connell"
    assert comments[0].replies[0].text == "Reply to comment #1"


def test_write_comments(configurable_save_file):
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.write(0, 0, "Test #1")
    table.write(1, 1, "Test #2")
    author = "Test User"
    table.add_comment("A1", "Root comment on A1", author=author)
    table.add_comment("B2", "Root comment on B2")
    table.add_comment("A1", "First reply to A1", author=author, reply=True)
    table.add_comment("A1", "Second reply to A1", author=author, reply=True)

    doc.save(configurable_save_file)

    new_doc = Document(configurable_save_file)
    new_table = new_doc.sheets[0].tables[0]

    a1_comments = new_table.cell("A1").comments
    b2_comments = new_table.cell(1, 1).comments

    assert len(a1_comments) == 1
    assert a1_comments[0].text == "Root comment on A1"
    assert a1_comments[0].author == author
    assert len(a1_comments[0].replies) == 2
    assert a1_comments[0].replies[0].text == "First reply to A1"
    assert a1_comments[0].replies[1].text == "Second reply to A1"

    assert len(b2_comments) == 1
    assert b2_comments[0].text == "Root comment on B2"
    assert b2_comments[0].author == "numbers_parser"
    assert len(b2_comments[0].replies) == 0


def test_comment_exceptions():
    doc = Document()
    table = doc.sheets[0].tables[0]

    with pytest.raises(ValueError, match="Cannot reply to a comment in a cell with no comments"):
        table.add_comment("C3", "Invalid reply", reply=True)

    table.add_comment("D4", "Valid root comment")

    with pytest.raises(TypeError, match="no text defined for comment"):
        table.add_comment("D4")

    with pytest.raises(ValueError, match="Cannot reply to existing comment without reply=True"):
        table.add_comment("D4", "Invalid second root comment")

    with pytest.raises(TypeError, match="too many positional arguments to add_comment"):
        table.add_comment("D4", "Too", "Many", "Arguments")


def test_comment_auto_expand(configurable_save_file):
    doc = Document()
    table = doc.sheets[0].tables[0]

    # Force table expansion (default is 12x8)
    table.add_comment(20, 10, "Out of bounds comment")
    assert table.num_rows == 21
    assert table.num_cols == 11

    doc.save(configurable_save_file)
    new_doc = Document(configurable_save_file)
    new_table = new_doc.sheets[0].tables[0]

    comments = new_table.cell(20, 10).comments
    assert len(comments) == 1
    assert comments[0].text == "Out of bounds comment"


def test_comment_on_merged_cell():
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.merge_cells("A1:B2")

    # Adding to the anchor is fine
    table.add_comment("A1", "Anchor comment")

    # Adding to a hidden merged cell should fail
    with pytest.raises(
        TypeError,
        match="Cannot add a comment to a merged cell; add it to the merge anchor instead",
    ):
        table.add_comment("B2", "Merged cell comment")
