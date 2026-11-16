"""Tests for deprecated API changes."""

import pytest

from numbers_parser import Document


def test_api_change() -> None:
    """Test deprecation warnings for cell image properties."""
    doc = Document("tests/data/issue-43.numbers")
    table = doc.sheets[0].tables[0]

    with pytest.warns(
        DeprecationWarning,
        match="image_data is deprecated and will be removed in the future",
    ) as record:
        image_data = table.cell("A1").image_data
    assert len(record) == 1
    assert image_data is not None
    assert len(image_data) == 87857

    with pytest.warns(
        DeprecationWarning,
        match="image_filename is deprecated and will be removed in the future",
    ) as record:
        image_filename = table.cell("A1").image_filename
    assert len(record) == 1
    assert image_filename is not None
    assert image_filename == "pasted-image.png"

    with pytest.warns(
        DeprecationWarning,
        match="image_data is deprecated and will be removed in the future",
    ) as record:
        image_data = table.cell("C1").image_data
    assert len(record) == 1
    assert image_data is None

    with pytest.warns(
        DeprecationWarning,
        match="image_filename is deprecated and will be removed in the future",
    ) as record:
        image_filename = table.cell("C1").image_filename
    assert len(record) == 1
    assert image_filename is None
