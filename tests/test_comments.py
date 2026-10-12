from datetime import datetime

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

    assert len(comments) == 3

    # Validate Comment 1 and its replies
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

    assert c3.replies[0].text == "Comment #3, Reply #1"
    assert c3.replies[0].author == "Jon Connell"
    assert isinstance(c3.replies[0].creation_date, datetime)

    assert c3.replies[1].text == "Comment #3, Reply #2"
    assert c3.replies[1].author == "Jon Connell"
    assert isinstance(c3.replies[1].creation_date, datetime)


def test_passthrough_comments(configurable_save_file):
    doc = Document("tests/data/test-comments.numbers")
    doc.save(configurable_save_file)

    new_doc = Document(configurable_save_file)
    comments = new_doc.default_table.cell(0, 0).comments
    assert isinstance(comments[0], Comment)
    assert comments[0].author == "Jon Connell"
    assert comments[0].replies[0].text == "Reply to comment #1"
