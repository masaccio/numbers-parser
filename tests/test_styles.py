from numbers_parser import Document, EmptyCell

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

    assert len(tables["Gradients"].cell(1, 1).bg_color) == 3
    assert tables["Gradients"].cell(0, 0).bg_color == [
        (86, 193, 255),
        (0, 77, 128),
    ]
    assert tables["Default Colours"].cell(1, 5).bg_color == (255, 66, 161)


def test_styles():
    doc = Document("tests/data/test-styles.numbers")
    sheets = doc.sheets
    table = sheets[0].tables[0]

    for row_num, row in enumerate(table.iter_rows()):
        if row_num == 0:
            for cell in row[1:]:
                assert cell.style_name == cell.value
        elif row_num == 1:
            assert row[1].style_name == "CustomStyle1"
            assert row[1].is_bold
            assert row[1].is_italic
            assert not row[1].is_underline
            assert not row[1].is_strikethrough
            assert row[1].font_size == 12.0
            assert row[1].font_name == "Arial"
            assert row[2].font_color == (0, 0, 0)

            assert row[2].style_name == "CustomStyle2"
            assert not row[2].is_bold
            assert not row[2].is_italic
            assert row[2].is_underline
            assert not row[2].is_strikethrough
            assert row[2].font_size == 13.0
            assert row[2].font_name == "Impact"
            assert row[2].font_color == (0, 0, 0)

            assert row[3].style_name == "CustomStyle3"
            assert not row[3].is_bold
            assert not row[3].is_italic
            assert not row[3].is_underline
            assert row[3].is_strikethrough
            assert row[3].font_size == 10.0
            assert row[3].font_name == "Menlo"
            assert row[3].font_color == (29, 177, 0)
        elif row[0].value == "Fonts":
            for cell in row[1:]:
                if not isinstance(cell, EmptyCell):
                    assert cell.font_name == cell.value

    # assert table.cell("A3").style_name == "Heading Red"
    # assert table.cell("B2").style_name == "ArialBold14"
    # assert table.cell("C2").style_name == "CourierNew12"
    # assert table.cell("B2").is_bold
    # assert not table.cell("B3").is_bold
    # assert table.cell("C2").is_italic
    # assert table.cell("D2").font_color == (29, 177, 0)
    # assert table.cell("B2").font_size == 14.0
    # assert table.cell("C2").font_size == 12.0
    # assert table.cell("D2").font_size == 11.0
    # assert table.cell("B2").font_name == "Arial-Black"
    # assert table.cell("C2").font_name == "CourierNewPSMT"
    # assert table.cell("D2").font_name == "HelveticaNeue"
