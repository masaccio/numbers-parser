"""Exercise opt-in refusal of ambiguous named-header references."""

from datetime import datetime
from pathlib import Path

import pytest

from numbers_parser import Document, UnsupportedError


@pytest.mark.parametrize("variant", ["number", "formatted", "true", "false", "date", "text_true"])
def test_named_header_can_be_refused_without_changing_formula(
    tmp_path: Path, variant: str
) -> None:
    """Retain compatibility while exposing a safe alternative for consumers."""
    doc = Document("tests/data/create-formulas.numbers")
    source = doc.sheets["Main Sheet"].tables["Reference Tests"]
    target = doc.sheets["Main Sheet"].tables["Extra Headers"]
    node = doc._model.formula_ast(source._table_id)[source.cell("A55")._formula_id][0]
    row = 1
    if variant == "formatted":
        target.set_cell_formatting("A2", "number", decimal_places=0)
    elif variant in {"true", "false"}:
        row = 3 if variant == "true" else 4
        node.AST_row.row = row - 54
    elif variant == "date":
        target.write("A2", datetime(2024, 2, 3, 14, 30))  # noqa: DTZ001
    elif variant == "text_true":
        target.write("A4", "another")
        target.write("A2", "TRUE")
    path = tmp_path / f"{variant}.numbers"
    doc.save(path)
    reopened = Document(path)
    target = reopened.sheets["Main Sheet"].tables["Extra Headers"]
    assert all(isinstance(target.cell(row, col).value, str) for col in (1, 2))
    cell = reopened.sheets["Main Sheet"].tables["Reference Tests"].cell("A55")
    assert cell.formula.startswith("COUNTA(")
    with pytest.raises(UnsupportedError, match="named header reference"):
        _ = cell.formula_without_header_names


def test_coordinate_formula_stays_available() -> None:
    """Ordinary coordinates are still rendered by both APIs."""
    table = Document("tests/data/test-10.numbers").sheets[0].tables[0]
    assert table.cell("A1").formula_without_header_names is None
    assert table.cell("B2").formula == "A1+A2"
    assert table.cell("B2").formula_without_header_names == "A1+A2"
