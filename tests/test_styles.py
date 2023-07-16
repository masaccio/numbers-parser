import pytest

from numbers_parser import Document, EmptyCell
from numbers_parser.cell import RGB, Alignment, Style

TEST_NUMBERED_REF = [
    "(1) double-paren-1",
    "2) paren-2",
    "III. upper-roman-3",
    "(IV) double-paren-roman-4",
    "V) paren-roman-5",
    "vi. lower-roman-6",
    "(vii) double-paren-lower-roman-7",
    "viii) paren-lower-roman-8",
    "I. letter-9",
    "(J) double-paren-letter-10",
    "K) paren-letter-11",
    "l. lower-letter-12",
    "(m) double-paren-lower-letter-13",
    "n) paren-lower-letter-14",
]


def test_bullets():
    doc = Document("tests/data/test-formats.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]

    assert table.cell(4, 2).bullets[0] == "star-bullet-1"
    assert table.cell(4, 2).bullets[2] == "dash-bullet"
    assert table.cell(3, 2).formatted_bullets[4] == "• bullet-5"
    assert table.cell(4, 2).formatted_bullets[0] == "★ star-bullet-1"
    assert table.cell(4, 2).formatted_bullets[2] == "- dash-bullet"
    assert table.cell(6, 2).value == "the quick brown fox jumped"
    assert table.cell(8, 2).formatted_bullets[1] == "2. blue"
    assert table.cell(8, 3).formatted_bullets[2] == "C. apple"
    assert table.cell(9, 2).formatted_bullets == TEST_NUMBERED_REF


def test_bg_colors():
    doc = Document("tests/data/test-bgcolour.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables

    assert len(tables["Gradients"].cell(1, 1).style.bg_color) == 3
    assert tables["Gradients"].cell(0, 0).style.bg_color == [
        (86, 193, 255),
        (0, 77, 128),
    ]
    assert tables["Default Colours"].cell(1, 5).style.bg_color == (255, 66, 161)


def test_styles():
    doc = Document("tests/data/test-styles.numbers")
    sheets = doc.sheets
    table = sheets["Styles"].tables[0]

    for row_num, row in enumerate(table.iter_rows()):
        if row_num == 0:
            for cell in row[1:]:
                assert cell.style.name == cell.value
        elif row_num == 1:
            assert row[1].style.name == "CustomStyle1"
            assert row[1].style.bold
            assert row[1].style.italic
            assert not row[1].style.underline
            assert not row[1].style.strikethrough
            assert row[1].style.font_size == 12.0
            assert row[1].style.font_name == "Arial"
            assert row[2].style.font_color == RGB(0, 0, 0)

            assert row[2].style.name == "CustomStyle2"
            assert not row[2].style.bold
            assert not row[2].style.italic
            assert row[2].style.underline
            assert not row[2].style.strikethrough
            assert row[2].style.font_size == 13.0
            assert row[2].style.font_name == "Impact"
            assert row[2].style.font_color == RGB(0, 0, 0)

            assert row[3].style.name == "CustomStyle3"
            assert not row[3].style.bold
            assert not row[3].style.italic
            assert not row[3].style.underline
            assert row[3].style.strikethrough
            assert row[3].style.font_size == 10.0
            assert row[3].style.font_name == "Menlo"
            assert row[3].style.font_color == RGB(29, 177, 0)
        elif row_num == 2:
            assert row[1].style.bold
            assert row[1].style.italic
            assert not row[1].style.underline
            assert not row[1].style.strikethrough
            assert row[1].style.font_size == 11.0

            assert not row[2].style.bold
            assert not row[2].style.italic
            assert row[2].style.underline
            assert row[2].style.strikethrough
            assert row[2].style.font_size == 9.0

            assert not row[3].style.bold
            assert not row[3].style.italic
            assert not row[3].style.underline
            assert not row[3].style.strikethrough
            assert row[3].style.font_color == RGB(29, 177, 0)
            assert row[3].style.font_size == 12.0
        elif row[0].value == "Fonts":
            for cell in row[1:]:
                if not isinstance(cell, EmptyCell):
                    assert cell.style.font_name == cell.value
        elif row[0].value == "Alignment":
            for cell in row[1:]:
                if not isinstance(cell, EmptyCell):
                    ref = "_".join(
                        [
                            cell.style.alignment.horizontal.name,
                            cell.style.alignment.vertical.name,
                        ]
                    )
                    assert cell.value == ref


def test_header_styles():
    doc = Document("tests/data/test-styles.numbers")
    sheets = doc.sheets
    table = sheets["Headers"].tables[0]

    assert all([table.cell(0, row_num).style.bold for row_num in range(0, 4)])
    assert all([table.cell(3, row_num).style.bold for row_num in range(0, 3)])
    assert all([table.cell(8, row_num).style.bold for row_num in range(0, 3)])
    assert all([table.cell(9, row_num).style.bold for row_num in range(0, 4)])

    assert all([not table.cell(ref).style.bold for ref in ["E1", "B5", "E9"]])
    assert all([table.cell(ref).style.underline for ref in ["C1", "C4", "C9"]])
    assert all([table.cell(ref).style.italic for ref in ["B1", "B4", "B9"]])
    assert all([table.cell(ref).style.strikethrough for ref in ["D1", "A5", "D9"]])

    assert all(
        [
            table.cell(ref).style.font_color == RGB(29, 177, 0)
            for ref in ["A2", "A6", "A10"]
        ]
    )
    assert all(
        [
            table.cell(ref).style.bg_color == RGB(29, 177, 0)
            for ref in ["B2", "B6", "B10"]
        ]
    )
    assert all(
        [
            table.cell(ref).style.bg_color == [RGB(136, 250, 78), RGB(1, 113, 0)]
            for ref in ["C2", "C6", "C10"]
        ]
    )
    assert all(
        [
            table.cell(ref).style.bg_image.filename
            == "pexels-evg-kowalievska-1170986-16.jpg"
            for ref in ["D2", "E2", "B7", "C7", "D10", "E10"]
        ]
    )
    assert all(
        [
            len(table.cell(ref).style.bg_image.data) == 418932
            for ref in ["D2", "E2", "B7", "C7", "D10", "E10"]
        ]
    )


def test_style_exceptions():
    doc = Document()
    table = doc.sheets[0].tables[0]
    with pytest.raises(TypeError) as e:
        doc.add_style(bg_color=(0, 0, 0, 0))
    assert "RGB color must be an RGB" in str(e)
    with pytest.raises(TypeError) as e:
        doc.add_style(bg_color=(0, 0, 1.0))
    assert "RGB color must be an RGB" in str(e)

    style = doc.add_style(bg_color=(100, 100, 100))
    assert style.bg_color == RGB(100, 100, 100)

    styles = doc.styles
    assert "Title" in styles
    assert len(styles) == 7

    dummy = doc.add_style()
    assert dummy.name == "Custom Style 2"
    dummy = doc.add_style()
    assert dummy.name == "Custom Style 3"

    with pytest.raises(TypeError) as e:
        _ = Style(alignment=Alignment("invalid", "top"))
    assert "invalid horizontal alignment" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Style(alignment=Alignment("left", "invalid"))
    assert "invalid vertical alignment" in str(e)

    with pytest.raises(IndexError) as e:
        _ = table.set_cell_style(0, 0, "Blue Text")
    assert "style 'Blue Text' does not exist" in str(e)

    with pytest.raises(TypeError) as e:
        _ = table.set_cell_style(0, 0, object)
    assert "style must be a Style object or style name" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Style(alignment=[None, object()])
    assert "value must be an Alignment class" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Style(font_size="invalid")
    assert "size must be a float number" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Style(font_name=2.0)
    assert "font name must be a string" in str(e)

    for field in ["bold", "italic", "underline", "strikethrough"]:
        with pytest.raises(TypeError) as e:
            attrs = {field: "invalid"}
            _ = Style(**attrs)
        assert f"{field} argument must be boolean" in str(e)


def test_new_styles(tmp_path, pytestconfig):
    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "test-styles-new.numbers"

    doc = Document()
    table = doc.sheets[0].tables[0]
    red_text = doc.add_style(
        name="Red Text",
        font_name="Lucida Grande",
        font_color=RGB(230, 25, 25),
        font_size=14.0,
        bold=True,
        italic=True,
        alignment=Alignment("right", "middle"),
    )
    assert red_text.name == "Red Text"

    with pytest.raises(IndexError) as e:
        _ = doc.add_style(name="Red Text")
    assert "style 'Red Text' already exists" in str(e)

    table.write("B2", "Red", style=red_text)

    table.write("C2", "Blue", style="Heading")
    table.cell("C2").style.font_color = (0, 160, 255)

    green_bg = doc.add_style(bg_color=RGB(29, 177, 0))
    assert green_bg.name == "Custom Style 1"
    table.set_cell_style("D2", green_bg)
    assert table.cell("D2").style.bg_color == RGB(29, 177, 0)
    assert table.cell("D2").style.font_name == "Helvetica Neue"

    table.set_cell_style("E2", "Heading Red")
    assert table.cell("E2").style.font_color == RGB(238, 34, 12)
    assert table.cell("E2").style.bold

    table.set_cell_style("F2", "Body")
    table.set_cell_style("G2", "Body")
    table.cell("F2").style.bg_color = RGB(238, 34, 12)
    table.cell("F2").style.alignment = Alignment("right", "middle")

    doc.save(new_filename)

    new_doc = Document(new_filename)
    new_table = new_doc.sheets[0].tables[0]

    assert new_table.cell("B2").value == "Red"
    assert new_table.cell("B2").style.font_color == RGB(230, 25, 25)
    assert new_table.cell("B2").style.font_size == 14.0
    assert new_table.cell("B2").style.font_name == "Lucida Grande"
    assert new_table.cell("B2").style.bold
    assert new_table.cell("B2").style.italic
    assert new_table.cell("B2").style.alignment == Alignment("right", "middle")

    assert new_table.cell("C2").value == "Blue"
    assert new_table.cell("C2").style.font_color == RGB(0, 160, 255)
    assert new_table.cell("C2").style.name == "Heading"

    assert new_table.cell("D2").style.bg_color == RGB(29, 177, 0)
    assert new_table.cell("D2").style.font_name == "Helvetica Neue"

    assert new_table.cell("E2").style.name == "Heading Red"
    assert new_table.cell("E2").style.font_color == RGB(238, 34, 12)
    assert new_table.cell("E2").style.bold

    assert new_table.cell("F2").style.name == "Body"
    assert new_table.cell("G2").style.name == "Body"
    assert new_table.cell("F2").style.bg_color == RGB(238, 34, 12)
    assert new_table.cell("F2").style.alignment == Alignment("right", "middle")


def test_empty_styles(tmp_path, pytestconfig):
    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "test-styles-new.numbers"

    doc = Document()
    red_text = doc.add_style(
        name="Red Text",
        font_name="Lucida Grande",
        font_color=RGB(230, 25, 25),
        font_size=10.0,
        alignment=Alignment("center", "middle"),
    )
    table = doc.sheets[0].tables[0]
    row_header_style = table.cell(0, 0).style.name
    body_style = table.cell(1, 1).style.name

    table.write(5, 5, "data", style=red_text)
    doc.save(new_filename)

    new_doc = Document(new_filename)
    new_table = new_doc.sheets[0].tables[0]

    for row_num in range(0, table.num_header_rows):
        for col_num in range(0, table.num_header_cols):
            assert new_table.cell(row_num, col_num).style.name == row_header_style

    for row_num in range(table.num_header_rows, table.num_rows):
        for col_num in range(table.num_header_cols, table.num_cols):
            if row_num == 5 and col_num == 5:
                assert new_table.cell(5, 5).style.name == "Red Text"
            else:
                assert new_table.cell(row_num, col_num).style.name == body_style
