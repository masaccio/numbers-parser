"""Keep formula range endpoints bound to their original tables."""

from pathlib import Path

import pytest

from numbers_parser import Document, UnsupportedError
from numbers_parser.formula import Formula
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.xrefs import CellRange


class Endpoint(CellRange):
    """A resolved endpoint with an explicit display name and target."""

    def __init__(self, target: int, text: str) -> None:
        self.to_table_id = target
        self._ref_error = False
        self.text = text

    def __str__(self) -> str:
        return self.text


def render(start: str | Endpoint, end: str | Endpoint) -> str:
    """Render a colon-node range from two endpoints."""
    formula = Formula(None, 0, 0, 0)
    formula.push(start)
    formula.push(end)
    formula.range(0, 0, None)
    return str(formula)


def test_same_target_retains_full_prefix() -> None:
    """Qualified endpoints retain the complete sheet and table name."""
    assert render(Endpoint(1, "Other::Values::A1"), Endpoint(1, "Other::Values::A2")) == (
        "Other::Values::A1:A2"
    )


def test_equal_display_names_do_not_hide_different_targets() -> None:
    """Different table IDs remain distinct even when their names match."""
    with pytest.raises(UnsupportedError):
        render(Endpoint(1, "Values::A1"), Endpoint(2, "Values::A2"))


def test_dynamic_endpoint_keeps_existing_rendering() -> None:
    """Function-derived endpoints remain available without name rewriting."""
    assert render("A1", "INDEX(A1:A3,1)") == "A1:INDEX(A1:A3,1)"


def test_saved_mixed_target_range_does_not_become_local(tmp_path: Path) -> None:
    """A real formula archive with different table targets must not be retargeted."""
    doc = Document("tests/data/create-formulas.numbers")
    table = doc.sheets["Main Sheet"].tables["Reference Tests"]
    first, second = table.cell("A10"), table.cell("A15")
    formulas = doc._model.formula_ast(table._table_id)
    nodes = formulas[first._formula_id]
    node_type = TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive
    first_ref, second_ref, function = node_type(), node_type(), node_type()
    first_ref.CopyFrom(nodes[0])
    second_ref.CopyFrom(formulas[second._formula_id][0])
    function.CopyFrom(nodes[1])
    del nodes[:]
    nodes.add().CopyFrom(first_ref)
    nodes.add().CopyFrom(second_ref)
    nodes.add().AST_node_type = (
        TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_name["COLON_NODE"].number
    )
    nodes.add().CopyFrom(function)
    path = tmp_path / "mixed-target.numbers"
    doc.save(path)
    reopened = Document(path).sheets["Main Sheet"].tables["Reference Tests"].cell("A10")
    with pytest.raises(UnsupportedError):
        _ = reopened.formula
