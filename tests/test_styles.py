"""Tests for cell and table styles."""

from collections import ChainMap
from pathlib import Path
from typing import Any

import pytest
from pytest_check import check

from numbers_parser import (
    RGB,
    Alignment,
    BackgroundImage,
    Cell,
    Document,
    HorizontalJustification,
    RichTextCell,
    Style,
    UnsupportedWarning,
    VerticalJustification,
)
from numbers_parser.cell import DEFAULT_ALIGNMENT_CLASS
from numbers_parser.constants import DEFAULT_FONT, DEFAULT_FONT_SIZE

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


def _style(cell: Cell) -> Style:
    """Return a cell's style after asserting it is not None."""
    assert cell.style is not None
    return cell.style


def test_bullets() -> None:
    """Test formatted bullets in cells."""
    doc = Document("tests/data/test-formats.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]

    cell_4_2 = table.cell(4, 2)
    assert isinstance(cell_4_2, RichTextCell)
    assert cell_4_2.bullets is not None
    assert cell_4_2.formatted_bullets is not None
    assert cell_4_2.bullets[0] == "star-bullet-1"
    assert cell_4_2.bullets[2] == "dash-bullet"

    cell_3_2 = table.cell(3, 2)
    assert isinstance(cell_3_2, RichTextCell)
    assert cell_3_2.formatted_bullets is not None
    assert cell_3_2.formatted_bullets[4] == "• bullet-5"

    assert cell_4_2.formatted_bullets[0] == "★ star-bullet-1"
    assert cell_4_2.formatted_bullets[2] == "- dash-bullet"
    assert table.cell(6, 2).value == "the quick brown fox jumped"

    cell_8_2 = table.cell(8, 2)
    assert isinstance(cell_8_2, RichTextCell)
    assert cell_8_2.formatted_bullets is not None
    assert cell_8_2.formatted_bullets[1] == "2. blue"

    cell_8_3 = table.cell(8, 3)
    assert isinstance(cell_8_3, RichTextCell)
    assert cell_8_3.formatted_bullets is not None
    assert cell_8_3.formatted_bullets[2] == "C. apple"

    cell_9_2 = table.cell(9, 2)
    assert isinstance(cell_9_2, RichTextCell)
    assert cell_9_2.formatted_bullets == TEST_NUMBERED_REF


def test_bg_colors() -> None:
    """Test cell background colors and gradients."""
    doc = Document("tests/data/test-bgcolour.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables

    grad_bg_1_1 = _style(tables["Gradients"].cell(1, 1)).bg_color
    assert isinstance(grad_bg_1_1, list)
    assert len(grad_bg_1_1) == 3
    assert _style(tables["Gradients"].cell(0, 0)).bg_color == [
        (86, 193, 255),
        (0, 77, 128),
    ]
    assert _style(tables["Default Colours"].cell(1, 5)).bg_color == (255, 66, 161)


STYLE_DEFAULTS = {
    "alignment": DEFAULT_ALIGNMENT_CLASS,
    "bg_image": None,
    "bg_color": None,
    "font_color": RGB(0, 0, 0),
    "font_size": DEFAULT_FONT_SIZE,
    "font_name": DEFAULT_FONT,
    "bold": False,
    "italic": False,
    "strikethrough": False,
    "underline": False,
    "first_indent": 0.0,
    "left_indent": 0.0,
    "right_indent": 0.0,
    "text_inset": 4.0,
}


def check_style(style: Style, **kwargs: Any) -> bool:  # noqa: ANN401
    """Check if a Style object matches expected attributes."""
    matches = []
    for attr, default_val in STYLE_DEFAULTS.items():
        matches.append(check.equal(getattr(style, attr), kwargs.get(attr, default_val)))
    return all(matches)


def decode_style_attrs(attrs: str) -> dict[str, Any]:
    """Decode style test specification string into an attribute dictionary."""
    if "," in attrs:
        return dict(ChainMap(*[decode_style_attrs(x) for x in attrs.split(",")]))
    if "=" in attrs:
        (attr, ref) = attrs.split("=")
        if isinstance(ref, str):
            if "color" in attr:  # pylint: disable=eval-used
                ref = eval(ref.replace(";", ","))  # pylint: disable=eval-used # noqa: S307
            elif "alignment" in attr:
                ref = eval(f"Alignment({ref.replace(';', ',')})")  # pylint: disable=eval-used # noqa: S307
            else:
                ref = eval(ref)  # pylint: disable=eval-used # noqa: S307

        return {attr: ref}
    return {attrs: True}


def invert_style_attrs(attrs: dict[str, Any]) -> dict[str, Any]:
    """Invert style attributes for modification testing."""
    for attr, default in STYLE_DEFAULTS.items():
        if attr not in attrs:
            attrs[attr] = default

    horizontal = HorizontalJustification((attrs["alignment"].horizontal + 1) % 4)
    vertical = VerticalJustification((attrs["alignment"].vertical + 1) % 3)
    attrs["alignment"] = Alignment(horizontal, vertical)
    if attrs["font_name"] == "Arial":
        attrs["font_name"] = "Menlo"
    else:
        attrs["font_name"] = "Arial"
    attrs["bold"] = not attrs["bold"]
    attrs["italic"] = not attrs["italic"]
    attrs["strikethrough"] = not attrs["strikethrough"]
    attrs["underline"] = not attrs["underline"]
    attrs["font_color"] = RGB(
        # 200 so black -> white text is still visible
        abs(200 - attrs["font_color"].r),
        abs(200 - attrs["font_color"].g),
        abs(200 - attrs["font_color"].b),
    )
    attrs["font_size"] = float(attrs["font_size"]) + 2.0
    attrs["first_indent"] = float(attrs["first_indent"]) + 2.0
    attrs["left_indent"] = float(attrs["left_indent"]) + 2.0
    attrs["right_indent"] = float(attrs["right_indent"]) + 2.0
    attrs["text_inset"] = float(attrs["text_inset"]) + 2.0
    return attrs


def test_all_styles(configurable_save_file: Path) -> None:
    """Test reading and resaving all cell styles."""
    doc = Document("tests/data/test-styles.numbers")
    table = doc.sheets["Styles"].tables[0]
    for row in range(table.num_rows):
        for col in range(table.num_cols):
            cell = table.cell(row, col)
            if not cell.value:
                continue
            attrs = decode_style_attrs(cell.value)
            valid = check_style(_style(cell), **attrs)
            assert valid

    # Re-save doc and check again
    doc.save(configurable_save_file)
    doc = Document(configurable_save_file)
    table = doc.sheets["Styles"].tables[0]
    for row in range(table.num_rows):
        for col in range(table.num_cols):
            cell = table.cell(row, col)
            if not cell.value:
                continue
            attrs = decode_style_attrs(cell.value)
            assert check_style(_style(cell), **attrs)


def test_all_style_changes(configurable_save_file: Path) -> None:
    """Test inverting, applying, and saving custom styles."""
    # Flip styles and re-save as new custom styles
    doc = Document("tests/data/test-styles.numbers")
    table = doc.sheets["Styles"].tables[0]
    for row in range(table.num_rows):
        for col in range(table.num_cols):
            cell = table.cell(row, col)
            if not cell.value:
                continue
            attrs = decode_style_attrs(cell.value)
            attrs = invert_style_attrs(attrs)
            if "name" in attrs:
                del attrs["name"]
            style = doc.add_style(**attrs)
            table.set_cell_style(row, col, style)

    doc.save(configurable_save_file)
    doc = Document(configurable_save_file)
    table = doc.sheets["Styles"].tables[0]
    style_num = 1
    for row in range(table.num_rows):
        for col in range(table.num_cols):
            cell = table.cell(row, col)
            if not cell.value:
                continue
            attrs = decode_style_attrs(cell.value)
            attrs = invert_style_attrs(attrs)
            cell_style = _style(cell)
            valid = check_style(cell_style, **attrs)
            assert valid
            assert cell_style.name == f"Custom Style {style_num}"
            style_num += 1

    style = doc.add_style(bold=True)
    assert style.name == f"Custom Style {style_num}"


def test_header_styles() -> None:
    """Test styles in header rows and columns."""
    doc = Document("tests/data/test-styles.numbers")
    sheets = doc.sheets
    table = sheets["Headers"].tables[0]

    assert all(_style(table.cell(0, row)).bold for row in range(4))
    assert all(_style(table.cell(3, row)).bold for row in range(3))
    assert all(_style(table.cell(8, row)).bold for row in range(3))
    assert all(_style(table.cell(9, row)).bold for row in range(4))

    assert all(not _style(table.cell(ref)).bold for ref in ["E1", "B5", "E9"])
    assert all(_style(table.cell(ref)).underline for ref in ["C1", "C4", "C9"])
    assert all(_style(table.cell(ref)).italic for ref in ["B1", "B4", "B9"])
    assert all(_style(table.cell(ref)).strikethrough for ref in ["D1", "A5", "D9"])

    assert all(_style(table.cell(ref)).font_color == RGB(29, 177, 0) for ref in ["A2", "A6", "A10"])
    assert all(_style(table.cell(ref)).bg_color == RGB(29, 177, 0) for ref in ["B2", "B6", "B10"])
    assert all(
        _style(table.cell(ref)).bg_color == [RGB(136, 250, 78), RGB(1, 113, 0)]
        for ref in ["C2", "C6", "C10"]
    )
    for ref in ["D2", "E2", "B7", "C7", "D10", "E10"]:
        bg_img = _style(table.cell(ref)).bg_image
        assert bg_img is not None
        assert bg_img.filename == "pexels-evg-kowalievska-1170986.jpg"
        assert len(bg_img.data) == 418932

    assert _style(table.cell("A9")).bold
    assert _style(table.cell("B9")).italic
    assert _style(table.cell("B9")).bold
    assert _style(table.cell("C9")).underline
    assert _style(table.cell("D9")).strikethrough
    assert not _style(table.cell("E9")).bold


def test_style_exceptions() -> None:
    """Test exceptions raised by invalid style parameters."""
    doc = Document()
    table = doc.sheets[0].tables[0]
    with pytest.raises(TypeError, match="RGB color must be an RGB"):
        doc.add_style(bg_color=(0, 0, 0, 0))  # type: ignore[arg-type]

    with pytest.raises(TypeError, match="RGB color must be an RGB"):
        doc.add_style(bg_color=(0, 0, 1.0))  # type: ignore[arg-type]

    style = Style(bg_color=(0, 0, 0))  # type: ignore[arg-type]
    with pytest.raises(TypeError, match="RGB color must be an RGB"):
        style.bg_color = object()  # type: ignore[assignment]

    style = doc.add_style(bg_color=(100, 100, 100))  # type: ignore[arg-type]
    assert style.bg_color == RGB(100, 100, 100)

    styles = doc.styles
    assert "Title" in styles
    assert len(styles) == 7

    dummy = doc.add_style()
    assert dummy.name == "Custom Style 2"
    dummy = doc.add_style()
    assert dummy.name == "Custom Style 3"

    style = Style(alignment=Alignment("left", "top"))
    style.alignment = None  # type: ignore[assignment]
    assert style.alignment.horizontal == HorizontalJustification.AUTO
    with pytest.raises(TypeError) as e:
        style.alignment = ("invalid", "top")  # type: ignore[assignment]
    assert "invalid horizontal alignment" in str(e)
    with pytest.raises(TypeError) as e:
        style.alignment = object()  # type: ignore[assignment]
    assert "Alignment must be an Alignment or a tuple" in str(e)
    with pytest.raises(TypeError) as e:
        style.alignment = (0, 0, 0)  # type: ignore[assignment]
    assert "Alignment must be an Alignment or a tuple" in str(e)
    with pytest.raises(TypeError) as e:
        _ = Style(alignment=Alignment("invalid", "top"))
    assert "invalid horizontal alignment" in str(e)
    with pytest.raises(TypeError) as e:
        _ = Style(alignment=Alignment("left", "invalid"))
    assert "invalid vertical alignment" in str(e)
    with pytest.raises(TypeError) as e:
        _ = Style(alignment=Alignment("left", "invalid"))
    with pytest.warns(UnsupportedWarning) as record:
        table.cell("A1").style = None
    assert len(record) == 1
    assert "cell style cannot be set" in str(record[0])

    with pytest.raises(IndexError) as idx_err:
        _ = table.set_cell_style(0, 0, "Blue Text")
    assert "style 'Blue Text' does not exist" in str(idx_err)

    with pytest.raises(TypeError) as e:
        _ = table.set_cell_style(0, 0, object)  # type: ignore[call-overload]
    assert "style must be a Style object or style name" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Style(alignment=[None, object()])  # type: ignore[arg-type]
    assert "Alignment must be an Alignment or a tuple" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Style(font_size="invalid")  # type: ignore[arg-type]
    assert "size must be a float number" in str(e)

    with pytest.raises(TypeError) as e:
        _ = Style(font_name=2.0)  # type: ignore[arg-type]
    assert "font name must be a string" in str(e)

    for field in ["bold", "italic", "underline", "strikethrough"]:
        with pytest.raises(TypeError) as e:  # noqa: PT012
            attrs: dict[str, Any] = {field: "invalid"}
            _ = Style(**attrs)
        assert f"{field} argument must be boolean" in str(e)


def test_new_styles(configurable_save_file: Path) -> None:
    """Test creating, applying, and saving new styles."""
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
    _style(table.cell("C2")).font_color = (0, 160, 255)  # type: ignore[assignment]

    green_bg = doc.add_style(bg_color=RGB(29, 177, 0))
    assert green_bg.name == "Custom Style 1"
    table.set_cell_style("D2", green_bg)
    assert _style(table.cell("D2")).bg_color == RGB(29, 177, 0)
    assert _style(table.cell("D2")).font_name == "Helvetica Neue"

    gradient_bg = doc.add_style(bg_color=[RGB(86, 193, 255), RGB(0, 77, 128)])
    table.set_cell_style("D3", gradient_bg)

    table.set_cell_style("E2", "Heading Red")
    assert _style(table.cell("E2")).font_color == RGB(238, 34, 12)
    assert _style(table.cell("E2")).bold

    table.set_cell_style("F2", "Body")
    table.set_cell_style("G2", "Body")
    _style(table.cell("F2")).bg_color = RGB(238, 34, 12)
    _style(table.cell("F2")).alignment = Alignment("right", "middle")

    table.write("F3", "Lorem ipsum dolor sit amet, consectetur adipiscing elit")
    table.write("F4", "Lorem ipsum dolor sit amet, consectetur adipiscing elit")
    assert _style(table.cell("F3")).text_wrap
    _style(table.cell("F3")).text_wrap = False

    doc.save(configurable_save_file)

    new_doc = Document(configurable_save_file)
    new_table = new_doc.sheets[0].tables[0]

    assert new_table.cell("B2").value == "Red"
    assert _style(new_table.cell("B2")).font_color == RGB(230, 25, 25)
    assert _style(new_table.cell("B2")).font_size == 14.0
    assert _style(new_table.cell("B2")).font_name == "Lucida Grande"
    assert _style(new_table.cell("B2")).bold
    assert _style(new_table.cell("B2")).italic
    assert _style(new_table.cell("B2")).alignment == Alignment("right", "middle")

    assert new_table.cell("C2").value == "Blue"
    assert _style(new_table.cell("C2")).font_color == RGB(0, 160, 255)
    assert _style(new_table.cell("C2")).name == "Heading"

    assert _style(new_table.cell("D2")).bg_color == RGB(29, 177, 0)
    assert _style(new_table.cell("D2")).font_name == "Helvetica Neue"

    assert _style(new_table.cell("D3")).bg_color == [RGB(86, 193, 255), RGB(0, 77, 128)]

    assert _style(new_table.cell("E2")).name == "Heading Red"
    assert _style(new_table.cell("E2")).font_color == RGB(238, 34, 12)
    assert _style(new_table.cell("E2")).bold

    assert _style(new_table.cell("F2")).name == "Body"
    assert _style(new_table.cell("G2")).name == "Body"
    assert _style(new_table.cell("F2")).bg_color == RGB(238, 34, 12)
    assert _style(new_table.cell("F2")).alignment == Alignment("right", "middle")

    assert not _style(new_table.cell("F3")).text_wrap


def test_empty_styles(configurable_save_file: Path) -> None:
    """Test preserving default styles on empty cells."""
    doc = Document()
    red_text = doc.add_style(
        name="Red Text",
        font_name="Lucida Grande",
        font_color=RGB(230, 25, 25),
        font_size=10.0,
        alignment=Alignment("center", "middle"),
    )
    table = doc.sheets[0].tables[0]
    row_header_style = _style(table.cell("A1")).name
    body_style = _style(table.cell("C2")).name

    table.write(5, 5, "data", style=red_text)
    doc.save(configurable_save_file)

    new_doc = Document(configurable_save_file)
    new_table = new_doc.sheets[0].tables[0]

    for row in range(table.num_header_rows):
        for col in range(table.num_header_cols):
            assert _style(new_table.cell(row, col)).name == row_header_style

    for row in range(table.num_header_rows, table.num_rows):
        for col in range(table.num_header_cols, table.num_cols):
            if row == 5 and col == 5:
                assert _style(new_table.cell(5, 5)).name == "Red Text"
            else:
                assert _style(new_table.cell(row, col)).name == body_style


def test_add_bg_image(configurable_save_file: Path) -> None:
    """Test adding background images to styles."""
    doc = Document()
    table = doc.sheets[0].tables[0]
    image_filename = "tests/data/cat.jpg"
    image_data = Path(image_filename).read_bytes()
    bg_image = BackgroundImage(image_data, image_filename)
    cats_bg_style = doc.add_style(name="Cats", bg_image=bg_image)

    with pytest.raises(TypeError) as e:
        _ = doc.add_style(name="Cats2", bg_image=object())  # type: ignore[arg-type]
    assert "bg_image must be a BackgroundImage object" in str(e)
    with pytest.raises(IndexError) as idx_err:
        _ = doc.add_style(name="Cats3", bg_image=BackgroundImage(image_data, image_filename))
    assert "cat.jpg: image already exists in document" in str(idx_err)

    table.write(0, 0, "❤️ cats", style=cats_bg_style)
    table.write(0, 1, "Still ❤️ cats", style=cats_bg_style)
    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    bg_0_0 = _style(table.cell(0, 0)).bg_image
    assert bg_0_0 is not None
    assert bg_0_0.filename == "cat.jpg"

    bg_0_1 = _style(table.cell(0, 1)).bg_image
    assert bg_0_1 is not None
    assert bg_0_1.filename == "cat.jpg"
