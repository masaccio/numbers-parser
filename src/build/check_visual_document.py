# ruff: noqa: S101, DTZ001

import subprocess
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

import pymupdf
from PIL import Image
from pixelmatch.contrib.PIL import pixelmatch
from rich import print  # noqa: A004
from rich.traceback import install

from numbers_parser import (
    RGB,
    Alignment,
    BackgroundImage,
    Border,
    BorderType,
    ControlFormattingType,
    Document,
    DurationStyle,
    FractionAccuracy,
    HorizontalJustification,
    MergedCell,
    NegativeNumberStyle,
    PaddingType,
    VerticalJustification,
    xl_cell_to_rowcol,
    xl_col_to_name,
    xl_col_to_offset,
    xl_range,
)
from numbers_parser.cell import CustomFormattingType
from numbers_parser.constants import CHECKBOX_TRUE_VALUE, STAR_RATING_VALUE

# Hijack the default exception hook to print beautiful tracebacks with local variables
install(show_locals=True)


def generate_doc_with_all_apis(filename: Path):
    doc = Document(num_header_rows=0, num_header_cols=0, num_cols=2, num_rows=20)
    sheet = doc.sheets[0]
    sheet.name = "API Test Sheet"
    table = sheet.tables[0]
    table.name = "API Exporter Reference"

    # Pre-allocate rows to avoid auto-expansion resizing issues
    table.col_width(0, 150)
    table.col_width(1, 420)
    table.row_height(1, 64)
    for r in range(2, table.num_rows):
        table.row_height(r, 32)

    row = 1  # Skip header row visually
    table.write(0, 0, "Rendered Feature")
    table.write(0, 1, "Explanation / API Usage")

    # --- Styles (without Background Image) ---
    style1 = doc.add_style(
        name="CustomStyle1",
        font_color=RGB(29, 177, 0),
        font_size=10.0,
        strikethrough=True,
        font_name="Menlo",
        bg_color=RGB(240, 240, 240),
        alignment=Alignment(HorizontalJustification.CENTER, VerticalJustification.MIDDLE),
        bold=True,
        italic=True,
        underline=True,
        first_indent=1.0,
        left_indent=1.0,
        right_indent=1.0,
        text_inset=1.0,
        text_wrap=False,
    )
    table.write(row, 0, "Styled Text", style=style1)
    table.write(
        row,
        1,
        'name="CustomStyle1", font_color=RGB(29, 177, 0), font_size=10.0, strikethrough=True, font_name="Menlo", bg_color=RGB(240, 240, 240), alignment=Alignment(HorizontalJustification.CENTER, VerticalJustification.MIDDLE), bold=True, italic=True, underline=True',
    )
    row += 1

    # --- Borders ---
    border1 = Border(width=2.5, color=RGB(255, 0, 0), style=BorderType.DASHES)
    table.write(row, 0, "Bordered Top")
    table.set_cell_border(row, 0, "top", border1)
    table.write(
        row, 1, 'Border(width=2.5, color=RGB(255, 0, 0), style=BorderType.DASHES) applied to "top"'
    )
    row += 1

    # --- Base Formatting ---
    table.write(row, 0, 255)
    table.set_cell_formatting(row, 0, "base", base=16, base_places=4, base_use_minus_sign=True)
    table.write(row, 1, "FormattingType.BASE: base=16, base_places=4, base_use_minus_sign=True")
    row += 1

    # --- Currency Formatting ---
    table.write(row, 0, -100.5)
    table.set_cell_formatting(
        row,
        0,
        "currency",
        currency_code="USD",
        decimal_places=2,
        negative_style=NegativeNumberStyle.MINUS,
        show_thousands_separator=True,
        use_accounting_style=True,
    )
    table.write(
        row,
        1,
        'FormattingType.CURRENCY: currency_code="USD", decimal_places=2, negative_style=NegativeNumberStyle.MINUS, show_thousands_separator=True, use_accounting_style=True',
    )
    row += 1

    # --- Fraction Formatting ---
    table.write(row, 0, 0.75)
    table.set_cell_formatting(row, 0, "fraction", fraction_accuracy=FractionAccuracy.QUARTERS)
    table.write(row, 1, "FormattingType.FRACTION: fraction_accuracy=FractionAccuracy.QUARTERS")
    row += 1

    # --- Datetime Formatting ---
    table.write(row, 0, datetime(2026, 1, 1, 12, 0))
    table.set_cell_formatting(row, 0, "datetime", date_time_format="yyyy-MM-dd HH:mm")
    table.write(row, 1, 'FormattingType.DATETIME: date_time_format="yyyy-MM-dd HH:mm"')
    row += 1

    # --- Duration Formatting ---
    table.write(row, 0, timedelta(days=1, hours=2))
    table.set_cell_formatting(row, 0, "duration", duration_style=DurationStyle.LONG)
    table.write(row, 1, "DURATION: duration_style=DurationStyle.LONG")
    row += 1

    # --- Tickbox Formatting ---
    table.write(row, 0, True)
    table.set_cell_formatting(row, 0, "tickbox")
    table.write(row, 1, 'FormattingType.TICKBOX: bool value (True) with "tickbox" format')
    row += 1

    # --- Rating Formatting ---
    table.write(row, 0, 3)
    table.set_cell_formatting(row, 0, "rating")
    table.write(row, 1, 'FormattingType.RATING: numeric value (3) with "rating" format')
    row += 1

    # --- Slider Formatting ---
    table.write(row, 0, 50)
    table.set_cell_formatting(
        row,
        0,
        "slider",
        control_format=ControlFormattingType.NUMBER,
        minimum=0.0,
        maximum=100.0,
        increment=10.0,
    )
    table.write(
        row,
        1,
        "FormattingType.SLIDER: control_format=ControlFormattingType.NUMBER, minimum=0.0, maximum=100.0, increment=10.0",
    )
    row += 1

    # --- Stepper Formatting ---
    table.write(row, 0, 5)
    table.set_cell_formatting(
        row,
        0,
        "stepper",
        control_format=ControlFormattingType.NUMBER,
        minimum=0.0,
        maximum=10.0,
        increment=1.0,
    )
    table.write(
        row,
        1,
        "FormattingType.STEPPER: control_format=ControlFormattingType.NUMBER, minimum=0.0, maximum=10.0, increment=1.0",
    )
    row += 1

    # --- Popup Formatting ---
    table.write(row, 0, "Option A")
    table.set_cell_formatting(
        row, 0, "popup", popup_values=["Option A", "Option B"], allow_none=False
    )
    table.write(
        row, 1, 'FormattingType.POPUP: popup_values=["Option A", "Option B"], allow_none=False'
    )
    row += 1

    # --- Custom Number Formatting ---
    custom_num = doc.add_custom_format(
        name="MyCustomNum",
        type="number",
        integer_format=PaddingType.ZEROS,
        decimal_format=PaddingType.SPACES,
        num_integers=4,
        num_decimals=2,
        show_thousands_separator=False,
    )
    table.write(row, 0, 42.1)
    table.set_cell_formatting(row, 0, "custom", format=custom_num)
    table.write(
        row,
        1,
        "NUMBER: integer_format=PaddingType.ZEROS, decimal_format=PaddingType.SPACES, num_integers=4, num_decimals=2",
    )
    row += 1

    # --- Custom Datetime Formatting ---
    custom_date = doc.add_custom_format(
        name="MyCustomDate", type="datetime", format="EEEE, d MMMM yyyy"
    )
    table.write(row, 0, datetime(2026, 10, 14))
    table.set_cell_formatting(row, 0, "custom", format=custom_date)
    table.write(row, 1, 'DATETIME: format="EEEE, d MMMM yyyy"')
    row += 1

    # --- Custom Text Formatting ---
    custom_text = doc.add_custom_format(name="MyCustomText", type="text", format="Prefix: %s")
    table.write(row, 0, "Test string")
    table.set_cell_formatting(row, 0, "custom", format=custom_text)
    table.write(row, 1, 'TEXT: format="Prefix: %s"')
    row += 1

    # --- Merging and Comments ---
    table.write(row, 0, "Merged Anchor")
    table.merge_cells(f"A{row + 1}:A{row + 2}")
    table.add_comment(f"A{row + 1}", "Comment on merged anchor", author="API Tester")
    table.add_comment(f"A{row + 1}", "Reply to comment", author="API Tester", reply=True)
    table.write(
        row,
        1,
        f'table.merge_cells("A{row + 1}:A{row + 2}"), table.add_comment(..., author="API Tester", reply=True)',
    )
    row += 2

    # --- XREFs and Utilities ---
    table.write(row, 0, "Xrefs test")
    table.write(
        row,
        1,
        f"xl_cell_to_rowcol('A1')={xl_cell_to_rowcol('A1')}, xl_col_to_name(2)={xl_col_to_name(2)}, xl_range(0,0,1,1)={xl_range(0, 0, 1, 1)}, xl_col_to_offset('B')={xl_col_to_offset('B')}",
    )
    row += 1

    # --- Background Image ---
    with open("tests/data/cat.jpg", "rb") as f:
        cat_data = f.read()
    cat_bg = BackgroundImage(cat_data, "cat.jpg")
    style_cat = doc.add_style(name="CatBackground", bg_image=cat_bg)
    table.write(row, 0, "Cat Image", style=style_cat)
    table.write(
        row,
        1,
        'doc.add_style(name="CatBackground", bg_image=BackgroundImage(cat_data, "cat.jpg"))',
    )
    table.row_height(row, 100)  # Give the cat room to breathe visually
    row += 1

    doc.save(filename)


def check_saved_document(filename: Path):
    doc = Document(filename)
    sheet = doc.sheets["API Test Sheet"]
    table = sheet.tables["API Exporter Reference"]

    # Check Styles
    style_cell = table.cell(1, 0)
    assert style_cell.value == "Styled Text"
    s = style_cell.style
    assert s.name == "CustomStyle1"
    assert s.font_color == RGB(29, 177, 0)
    assert s.font_size == 10.0
    assert s.strikethrough is True
    assert s.font_name == "Menlo"
    assert s.bg_color == RGB(240, 240, 240)
    assert s.alignment.horizontal == HorizontalJustification.CENTER
    assert s.alignment.vertical == VerticalJustification.MIDDLE
    assert s.bold is True
    assert s.italic is True
    assert s.underline is True
    assert s.first_indent == 1.0
    assert s.left_indent == 1.0
    assert s.right_indent == 1.0
    assert s.text_inset == 1.0
    assert s.text_wrap is False

    # Check Borders
    border_cell = table.cell(2, 0)
    b = border_cell.border.top
    assert b.width == 2.5
    assert b.color == RGB(255, 0, 0)
    assert b.style == BorderType.DASHES

    # Check Base Formatting
    base_cell = table.cell(3, 0)
    assert base_cell.value == 255
    assert base_cell.formatted_value == "00FF"

    # Check Currency Formatting
    curr_cell = table.cell(4, 0)
    assert curr_cell.value == -100.5

    # Check Fraction Formatting
    frac_cell = table.cell(5, 0)
    assert frac_cell.value == 0.75
    assert frac_cell.formatted_value == "3/4"

    # Check Datetime Formatting
    dt_cell = table.cell(6, 0)
    assert "2026-01-01 12:00" in dt_cell.formatted_value

    # Check Duration Formatting
    dur_cell = table.cell(7, 0)
    assert isinstance(dur_cell.value, timedelta)
    assert ":" in dur_cell.formatted_value

    # Check Tickbox Formatting
    tick_cell = table.cell(8, 0)
    assert tick_cell.value is True
    assert tick_cell.formatted_value == CHECKBOX_TRUE_VALUE

    # Check Rating Formatting
    rat_cell = table.cell(9, 0)
    assert rat_cell.value == 3
    assert rat_cell.formatted_value == STAR_RATING_VALUE * 3

    # Check Slider & Stepper
    slid_cell = table.cell(10, 0)
    assert slid_cell.formatted_value == "50.0"
    step_cell = table.cell(11, 0)
    assert step_cell.formatted_value == "5"

    # Check Popup Formatting
    pop_cell = table.cell(12, 0)
    assert pop_cell.value == "Option A"

    # Check Custom Formats inside doc.custom_formats
    assert "MyCustomNum" in doc.custom_formats
    c_num = doc.custom_formats["MyCustomNum"]
    assert c_num.type == CustomFormattingType.NUMBER
    c_num_cell = table.cell(13, 0)
    assert c_num_cell.value == 42.1
    assert "42" in c_num_cell.formatted_value

    assert "MyCustomDate" in doc.custom_formats
    c_date = doc.custom_formats["MyCustomDate"]
    assert c_date.type == CustomFormattingType.DATETIME
    c_date_cell = table.cell(14, 0)
    assert "2026" in c_date_cell.formatted_value

    assert "MyCustomText" in doc.custom_formats
    c_text = doc.custom_formats["MyCustomText"]
    assert c_text.type == CustomFormattingType.TEXT
    c_text_cell = table.cell(15, 0)
    assert "Prefix: Test string" in c_text_cell.formatted_value

    # Check Merging
    assert "A17:A18" in table.merge_ranges
    anchor_cell = table.cell("A17")
    merged_cell = table.cell("A18")
    assert merged_cell.is_merged is False
    assert anchor_cell.size == (2, 1)
    assert isinstance(merged_cell, MergedCell)

    # Check Comments
    comments = anchor_cell.comments
    assert len(comments) == 1
    assert comments[0].text == "Comment on merged anchor"
    assert comments[0].author == "API Tester"
    assert len(comments[0].replies) == 1
    assert comments[0].replies[0].text == "Reply to comment"
    assert comments[0].replies[0].author == "API Tester"

    # Check Background Image
    bg_cell = table.cell(19, 0)
    assert bg_cell.value == "Cat Image"
    s_bg = bg_cell.style
    assert s_bg.name == "CatBackground"
    assert s_bg.bg_image is not None
    assert s_bg.bg_image.filename == "cat.jpg"
    with open("tests/data/cat.jpg", "rb") as f:
        assert s_bg.bg_image.data == f.read()


def open_numbers_and_export_pdf(doc_filepath: Path, pdf_filepath: Path):
    """Open a document in Apple Numbers and export as PDF."""
    # Remove stale PDF to prevent false positives if export silently fails
    if pdf_filepath.exists():
        pdf_filepath.unlink()

    applescript = f"""
    tell application "Numbers"
        activate
        open POSIX file "{doc_filepath.resolve()}"
        
        -- Robust wait loop with timeout
        set waitCounter to 0
        repeat until exists front window
            delay 0.5
            set waitCounter to waitCounter + 1
            if waitCounter > 20 then
                error "Timeout: Numbers took too long to open the window."
            end if
        end repeat
        
        set targetDoc to front document
        export targetDoc to POSIX file "{pdf_filepath.resolve()}" as PDF
        close front document saving no
        quit
    end tell
    """

    try:
        subprocess.run(["osascript", "-e", applescript], check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        print(f"[bold red]AppleScript execution failed:[/bold red]\n{e.stderr}")
        raise


def compare_pdfs_visually(
    ref_pdf_path: str, test_pdf_path: str, diff_out_path: str, threshold: float = 0.1
) -> bool:
    ref_doc = pymupdf.open(ref_pdf_path)
    test_doc = pymupdf.open(test_pdf_path)
    all_pages_match = True

    if len(ref_doc) != len(test_doc):
        print(f"Page count mismatch: {len(ref_doc)} vs {len(test_doc)}")
        return False

    for page_num in range(len(ref_doc)):
        matrix = pymupdf.Matrix(2.0, 2.0)

        ref_pix = ref_doc[page_num].get_pixmap(matrix=matrix, alpha=True)
        test_pix = test_doc[page_num].get_pixmap(matrix=matrix, alpha=True)

        ref_img = Image.frombytes("RGBA", [ref_pix.width, ref_pix.height], ref_pix.samples)
        test_img = Image.frombytes("RGBA", [test_pix.width, test_pix.height], test_pix.samples)

        # Force the test image to exactly match the reference image dimensions
        if ref_img.size != test_img.size:
            normalized_test_img = Image.new("RGBA", ref_img.size, (255, 255, 255, 0))
            normalized_test_img.paste(test_img, (0, 0))
            test_img = normalized_test_img

        diff_img = Image.new("RGBA", ref_img.size)

        mismatch_pixels = pixelmatch(
            ref_img, test_img, diff_img, threshold=threshold, includeAA=True
        )

        if mismatch_pixels > 0:
            print(
                f"[red]Visual difference found on page {page_num + 1} ({mismatch_pixels} pixels mismatched)[/red]"
            )
            diff_file = f"{diff_out_path}_page_{page_num + 1}.png"
            diff_img.save(diff_file)
            all_pages_match = False

    return all_pages_match


def main():
    tempdir = Path(tempfile.gettempdir())
    doc_filepath = tempdir / "api-test.numbers"
    pdf_filepath = tempdir / "api-test.pdf"

    print(f"[green]Generating doc using all APIs to: {doc_filepath}[/green]")
    generate_doc_with_all_apis(doc_filepath)

    print("[green]Running API read assertions...[/green]")
    check_saved_document(doc_filepath)
    print("[green]Document API check passed[/green]")

    print(f"[green]Triggering AppleScript PDF Export to: {pdf_filepath}[/green]")
    open_numbers_and_export_pdf(doc_filepath, pdf_filepath)

    if not pdf_filepath.exists():
        print(f"[bold red]Failed to locate exported PDF at {pdf_filepath}[/bold red]")
        return

    print("[green]Comparing generated PDF against visual reference...[/green]")
    match = compare_pdfs_visually(
        ref_pdf_path="tests/data/visual-test.pdf",
        test_pdf_path=str(pdf_filepath),
        diff_out_path="visual_check",
    )

    if match:
        print("[bold green]Visual test passed flawlessly![/bold green]")
    else:
        print("[bold red]Visual test failed. Check the generated diff images.[/bold red]")


if __name__ == "__main__":
    main()
