import pytest

from numbers_parser import Document, FractionAccuracy


@pytest.mark.parametrize("accuracy", [FractionAccuracy.HALVES, FractionAccuracy.THREE])
@pytest.mark.parametrize(
    ("value", "expected"),
    [(-2.5, "-2 1/2"), (-0.5, "-1/2"), (-2.0, "-2")],
)
def test_negative_fraction_keeps_whole_and_sign(accuracy, value, expected):
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.write(0, 0, value)
    table.set_cell_formatting(0, 0, "fraction", fraction_accuracy=accuracy)

    assert table.cell(0, 0).formatted_value == expected
