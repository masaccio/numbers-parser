"""Parse and extract data from Apple Numbers spreadsheets."""

import importlib.metadata
import os
import plistlib
import re
import warnings

from numbers_parser.constants import SUPPORTED_NUMBERS_VERSIONS

with warnings.catch_warnings():
    # Protobuf
    warnings.filterwarnings("ignore", category=DeprecationWarning)
    from numbers_parser.cell import (
        RGB,
        Alignment,
        BackgroundImage,
        BoolCell,
        Border,
        BorderType,
        BulletedTextCell,
        Cell,
        CellBorder,
        Comment,
        CustomFormatting,
        DateCell,
        DurationCell,
        EmptyCell,
        ErrorCell,
        Formatting,
        HorizontalJustification,
        MergeAnchor,
        MergedCell,
        MergeReference,
        NumberCell,
        RichTextCell,
        Style,
        TextCell,
        VerticalJustification,
    )
    from numbers_parser.constants import (
        CellPadding,
        CellType,
        CellValueType,
        ControlFormattingType,
        DurationStyle,
        DurationUnits,
        FormattingType,
        FormatType,
        FractionAccuracy,
        NegativeNumberStyle,
        PaddingType,
    )
    from numbers_parser.document import Document, Sheet, Table
    from numbers_parser.exceptions import (
        FileError,
        FileFormatError,
        FormulaError,
        NotImplementedError,  # noqa: A004
        NumbersError,
        UnsupportedError,
        UnsupportedWarning,
    )
    from numbers_parser.xrefs import (
        xl_cell_to_rowcol,
        xl_col_to_name,
        xl_col_to_offset,
        xl_range,
        xl_rowcol_to_cell,
    )

__all__ = [
    "RGB",
    "Alignment",
    "BackgroundImage",
    "BoolCell",
    "Border",
    "BorderType",
    "BulletedTextCell",
    "Cell",
    "CellBorder",
    "CellPadding",
    "CellType",
    "CellValueType",
    "Comment",
    "ControlFormattingType",
    "CustomFormatting",
    "DateCell",
    "Document",
    "DurationCell",
    "DurationStyle",
    "DurationUnits",
    "EmptyCell",
    "ErrorCell",
    "FileError",
    "FileFormatError",
    "FormatType",
    "Formatting",
    "FormattingType",
    "FormulaError",
    "FractionAccuracy",
    "HorizontalJustification",
    "MergeAnchor",
    "MergeReference",
    "MergedCell",
    "NegativeNumberStyle",
    "NotImplementedError",
    "NumberCell",
    "NumbersError",
    "PaddingType",
    "RichTextCell",
    "Sheet",
    "Style",
    "Table",
    "TextCell",
    "UnsupportedError",
    "UnsupportedWarning",
    "VerticalJustification",
    "xl_cell_to_rowcol",
    "xl_col_to_name",
    "xl_col_to_offset",
    "xl_range",
    "xl_rowcol_to_cell",
]

__version__ = importlib.metadata.version("numbers-parser")

_DEFAULT_NUMBERS_INSTALL_PATH = "/Applications/Numbers.app"
_VERSION_PLIST_PATH = "Contents/version.plist"


def _get_version() -> str:
    """Return the version."""
    return __version__


def _check_installed_numbers_version() -> str | None:
    """Check installed numbers version."""
    try:
        with open(
            os.path.join(_DEFAULT_NUMBERS_INSTALL_PATH, _VERSION_PLIST_PATH),
            "rb",
        ) as fp:
            version_dict = plistlib.load(fp)
    except OSError:
        return None
    installed_version = version_dict["CFBundleShortVersionString"]

    installed_version = re.sub(r"(\d+)\.(\d+)\.\d+", r"\1.\2", installed_version)
    if installed_version not in SUPPORTED_NUMBERS_VERSIONS:
        warnings.warn(
            f"Numbers version {installed_version} not tested with this version",
            stacklevel=2,
        )
    return installed_version


_ = _check_installed_numbers_version()
