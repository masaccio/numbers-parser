"""Expose grouping state even when the modern owner is absent."""

from pathlib import Path

import pytest

from numbers_parser import Document, UnsupportedError
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives


def test_grouping_checks_legacy_owner_after_save(tmp_path: Path) -> None:
    """An embedded legacy owner must not look like an ungrouped table."""
    doc = Document("tests/data/test-categories.numbers")
    table = doc.sheets[0].tables["Categories"]
    assert table.has_grouping is True
    doc._model.table_model(table._table_id).ClearField("category_owner")
    path = tmp_path / "legacy-owner.numbers"
    doc.save(path)
    reopened = Document(path).sheets[0].tables["Categories"]
    assert reopened.categorized_data() is None
    assert reopened.has_grouping is True


def test_grouping_rejects_missing_state() -> None:
    """Unknown grouping state is not evidence that grouping is disabled."""
    doc = Document("tests/data/test-10.numbers")
    table = doc.sheets[0].tables[0]
    assert table.has_grouping is False
    model = doc._model.table_model(table._table_id)
    model.ClearField("category_owner")
    model.ClearField("category_owner_deprecated")
    with pytest.raises(UnsupportedError):
        _ = table.has_grouping


def test_grouping_checks_both_owners() -> None:
    """Either enabled owner makes the table grouped."""
    doc = Document("tests/data/test-categories.numbers")
    table = doc.sheets[0].tables["Categories"]
    model = doc._model.table_model(table._table_id)
    model.category_owner_deprecated.group_by[0].is_enabled = False
    assert table.has_grouping is True
    model.category_owner_deprecated.group_by[0].is_enabled = True
    owner = doc._model.get_object(
        model.category_owner.identifier, TSTArchives.CategoryOwnerRefArchive
    )
    group = doc._model.get_object(owner.group_by[0].identifier, TSTArchives.GroupByArchive)
    group.is_enabled = False
    assert table.has_grouping is True
