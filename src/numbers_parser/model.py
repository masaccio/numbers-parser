"""Internal data model and archive operations for Numbers documents."""

# ruff: noqa: SLF001
from __future__ import annotations

import logging
import re
from array import array
from collections import defaultdict
from datetime import datetime, timedelta
from hashlib import sha1
from itertools import chain
from math import floor
from pathlib import Path
from struct import pack
from typing import TYPE_CHECKING, Any, cast
from warnings import warn

from numbers_parser.bullets import (
    BULLET_CONVERSION,
    BULLET_PREFIXES,
    BULLET_SUFFIXES,
)
from numbers_parser.cell import (
    RGB,
    Alignment,
    Border,
    BorderType,
    Cell,
    Comment,
    CustomFormatting,
    Formatting,
    FormattingType,
    HorizontalJustification,
    MergeAnchor,
    MergedCell,
    MergeReference,
    PaddingType,
    Style,
    VerticalJustification,
    _decode_date_format,
)
from numbers_parser.constants import (
    ALLOWED_FORMATTING_PARAMETERS,
    COLON_TRACT_NODE,
    CUSTOM_FORMAT_TYPE_MAP,
    CUSTOM_TEXT_PLACEHOLDER,
    DEFAULT_COLUMN_WIDTH,
    DEFAULT_DOCUMENT,
    DEFAULT_FONT,
    DEFAULT_PASSWORD_HINT,
    DEFAULT_PRE_BNC_BYTES,
    DEFAULT_ROW_HEIGHT,
    DEFAULT_TABLE_OFFSET,
    DEFAULT_TEXT_INSET,
    DEFAULT_TEXT_WRAP,
    DEFAULT_TILE_SIZE,
    DOCUMENT_ID,
    EPOCH,
    FONT_FAMILY_DEFAULT,
    FONT_MAP,
    FORMAT_TYPE_MAP,
    MAX_TILE_SIZE,
    PACKAGE_ID,
    CellInteractionType,
    CellValueType,
    FormatType,
    OwnerKind,
)
from numbers_parser.containers import ObjectStore, _ProtoT
from numbers_parser.exceptions import UnsupportedError, UnsupportedWarning
from numbers_parser.formula import TableFormulas
from numbers_parser.generated import TNArchives_pb2 as TNArchives
from numbers_parser.generated import TSAArchives_pb2 as TSAArchives
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.generated import TSDArchives_pb2 as TSDArchives
from numbers_parser.generated import TSKArchives_pb2 as TSKArchives
from numbers_parser.generated import TSPArchiveMessages_pb2 as TSPArchiveMessages
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages
from numbers_parser.generated import TSSArchives_pb2 as TSSArchives
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.generated import TSWPArchives_pb2 as TSWPArchives
from numbers_parser.generated.TSDArchives_pb2 import (
    StrokePatternArchive as StrokePattern,
)
from numbers_parser.generated.TSWPArchives_pb2 import (
    CharacterStylePropertiesArchive as CharacterStyle,
)
from numbers_parser.iwafile import find_extension
from numbers_parser.numbers_cache import Cacheable, cache
from numbers_parser.numbers_uuid import NumbersUUID, uuid_to_hex
from numbers_parser.xrefs import CellRange, ScopedNameRefCache

if TYPE_CHECKING:
    from collections.abc import Sequence

    from google.protobuf.message import Message as ProtoMessage

    from numbers_parser.generated.TSCEArchives_pb2 import ASTNodeArrayArchive

    ASTNodeArchive = ASTNodeArrayArchive.ASTNodeArchive
    ASTColonTractArchive = ASTNodeArrayArchive.ASTColonTractArchive


logger = logging.getLogger(__name__)
debug = logger.debug


class MergeCells:
    """Track merged-cell anchors and references in a table."""

    _references: dict[tuple[int, int], MergeReference | MergeAnchor | None]

    def __init__(self) -> None:
        """Initialize the instance state."""
        self._references = defaultdict(lambda: None)

    def add_reference(self, row: int, col: int, rect: tuple) -> None:
        """Add reference."""
        self._references[(row, col)] = MergeReference(*rect)

    def add_anchor(self, row: int, col: int, size: tuple) -> None:
        """Add anchor."""
        self._references[(row, col)] = MergeAnchor(size)

    def is_merge_reference(self, row_col: tuple) -> bool:
        # defaultdict will default this to False for missing entries
        """Return whether merge reference is available."""
        return isinstance(self._references[row_col], MergeReference)

    def is_merge_anchor(self, row_col: tuple) -> bool:
        # defaultdict will default this to False for missing entries
        """Return whether merge anchor is available."""
        return isinstance(self._references[row_col], MergeAnchor)

    def get(self, row_col: tuple) -> MergeAnchor | MergeReference | None:
        """Return the merge marker stored at the row and column."""
        return self._references[row_col]

    def size(self, row_col: tuple) -> tuple[int, int]:
        """Return the dimensions of the merged cell range."""
        ref = self._references[row_col]
        if isinstance(ref, MergeAnchor):
            return ref.size
        return (0, 0)

    def rect(self, row_col: tuple) -> tuple | None:
        """Return the coordinates covered by the merged cell range."""
        ref = self._references[row_col]
        if isinstance(ref, MergeReference):
            return ref.rect
        return None

    def merge_cells(self) -> list[tuple[int, int]]:
        """Merge cells."""
        return [k for k, v in self._references.items() if self.is_merge_anchor(k)]


class DataLists(Cacheable):
    """Model for TST.DataList with caching and key generation for new values."""

    _model: NumbersModel
    _datalists: dict[int, dict]
    _datalist_name: str

    def __init__(
        self,
        model: NumbersModel,
        datalist_name: str,
        value_attr: str,
    ) -> None:
        """Initialize the instance state."""
        self._model = model
        self._datalists = {}
        self._value_attr = value_attr
        self._datalist_name = datalist_name

    @cache()
    def add_table(self, table_id: int) -> None:
        """Cache a new datalist for a table if not already seen."""
        table_model = self._model.table_model(table_id)
        base_data_store = table_model.base_data_store
        datalist_id = getattr(base_data_store, self._datalist_name).identifier
        datalist = self._model.get_object(datalist_id, TSTArchives.TableDataList)

        max_key = 0
        self._datalists[table_id] = {}
        self._datalists[table_id]["by_key"] = {}
        self._datalists[table_id]["by_value"] = {}
        self._datalists[table_id]["key_index"] = {}
        self._datalists[table_id]["datalist"] = datalist
        self._datalists[table_id]["id"] = datalist_id
        for i, entry in enumerate(sorted(datalist.entries, key=lambda x: x.key)):
            if entry.key > max_key:
                max_key = entry.key
                self._datalists[table_id]["by_key"][entry.key] = entry
                self._datalists[table_id]["key_index"][entry.key] = i
                value_key = self.value_key(getattr(entry, self._value_attr))
                self._datalists[table_id]["by_value"][value_key] = entry.key
        self._datalists[table_id]["next_key"] = max_key + 1

    def id(self, table_id: int) -> int:
        """Return the archive identifier for this object."""
        self.add_table(table_id)
        return self._datalists[table_id]["id"]

    def lookup_value(
        self,
        table_id: int,
        key: int,
    ) -> TSTArchives.TableDataList.ListEntry:
        """Return the an entry in a table's datalist matching a key."""
        self.add_table(table_id)
        return self._datalists[table_id]["by_key"][key]

    def value_key(self, value: str | ProtoMessage) -> str:
        """Return the key used to identify this cell value."""
        if hasattr(value, "DESCRIPTOR"):
            return repr(value)
        return value

    def init(self, table_id: int) -> None:
        """Remove all entries from a datalist."""
        self.add_table(table_id)
        self._datalists[table_id]["by_key"] = {}
        self._datalists[table_id]["by_value"] = {}
        self._datalists[table_id]["key_index"] = {}
        self._datalists[table_id]["next_key"] = 1
        self._datalists[table_id]["datalist"].nextListID = 1
        self._datalists[table_id]["datalist"].entries.clear()

    def lookup_key(self, table_id: int, value: str | ProtoMessage) -> int:
        """
        Return the key associated with a value for a particular table entry.

        If the value is not in the datalist, allocate a new entry with the
        next available key.
        """
        self.add_table(table_id)
        value_key = self.value_key(value)
        if value_key not in self._datalists[table_id]["by_value"]:
            key = self._datalists[table_id]["next_key"]
            self._datalists[table_id]["next_key"] += 1
            self._datalists[table_id]["datalist"].nextListID += 1
            attrs = {"key": key, self._value_attr: value, "refcount": 1}
            entry = TSTArchives.TableDataList.ListEntry(**attrs)
            self._datalists[table_id]["datalist"].entries.append(entry)
            self._datalists[table_id]["by_key"][key] = entry
            self._datalists[table_id]["key_index"][key] = (
                len(self._datalists[table_id]["datalist"].entries) - 1
            )
            self._datalists[table_id]["by_value"][value_key] = key
        else:
            value_key = self.value_key(value)
            key = self._datalists[table_id]["by_value"][value_key]
            index = self._datalists[table_id]["key_index"][key]
            self._datalists[table_id]["datalist"].entries[index].refcount += 1

        return key


class NumbersModel(Cacheable):
    """
    Load Numbers document objects and abstract their internal structures.

    Do not use this class in application code.
    """

    def __init__(self, filepath: Path | None, password: str | None) -> None:
        """Initialize the instance state."""
        if filepath is None:
            filepath = Path(DEFAULT_DOCUMENT)
        self.objects = ObjectStore(filepath, password)
        self._merge_cells: dict[int, MergeCells] = defaultdict(MergeCells)
        self._row_heights: dict[int, dict[int, int]] = {}
        self._col_widths: dict[int, dict[int, int]] = {}
        self._table_formats = DataLists(self, "format_table", "format")
        self.table_styles = DataLists(self, "styleTable", "reference")
        self._table_strings = DataLists(self, "stringTable", "string")
        self._table_comments = DataLists(self, "commentStorageTable", "comment_storage")
        self._control_specs = DataLists(self, "control_cell_spec_table", "cell_spec")
        self._formulas = DataLists(self, "formula_table", "formula")
        self._table_data: dict[int, list[list[Cell]]] = {}
        self.table_categories_data: dict[int, dict[str, Any] | None] = {}
        self.table_categories_row_mapper: dict[int, dict[int, int] | None] = {}
        self._styles: dict[str, Style] = {}
        self._update_strokes: dict[int, bool] = {}
        self.images: dict[bytes, int] = {}
        self._custom_formats: dict[str, CustomFormatting] = {}
        self._custom_format_archives: dict[str, TSKArchives.CustomFormatArchive] = {}
        self._custom_format_uuids: dict[str, TSPMessages.UUID] = {}
        self.name_ref_cache = ScopedNameRefCache(self)
        self.missing_fonts: dict[str, bool] = {}
        self.calculate_table_uuid_map()

    def save(
        self,
        filepath: Path,
        package: bool,
        password: str | None,
        hint: str = DEFAULT_PASSWORD_HINT,
    ) -> None:
        """Save the value."""
        self.refresh_pivot_tables()
        self.objects.save(filepath, package, password, hint)

    def get_object(self, key: int, expected_type: type[_ProtoT]) -> _ProtoT:
        """Fetch an object from the object store and cast it to the expected Protobuf class."""
        return self.objects.get_as(key, expected_type)

    def find_refs(self, ref: str) -> list[int]:
        """Find refs."""
        return self.objects.find_refs(ref)

    @property
    def document_archive(self) -> TNArchives.DocumentArchive:
        """Return the Document protobuf archive for the current document."""
        return self.get_object(DOCUMENT_ID, TNArchives.DocumentArchive)

    @property
    def package_metadata(self) -> TSPArchiveMessages.PackageMetadata:
        """Return the metadata archive for the current document."""
        return self.get_object(PACKAGE_ID, TSPArchiveMessages.PackageMetadata)

    @cache(num_args=0)
    def calc_engine_id(self) -> int:
        """Return the CalculationEngine ID for the current document."""
        try:
            ce_id = self.find_refs("CalculationEngineArchive")
            return ce_id[0]
        except IndexError as e:
            # Issue #17 breaks but the CalculationEngine is fundamental
            # to the function of numbers-parser
            msg = "Documents without a CalculationEngine archive are unsupported"
            raise UnsupportedError(msg) from e

    @cache(num_args=0)
    def calc_engine(self) -> TSCEArchives.CalculationEngineArchive:
        """Return the CalculationEngine object for the current document."""
        return self.get_object(self.calc_engine_id(), TSCEArchives.CalculationEngineArchive)

    def sheet_ids(self) -> list[int]:
        """Return the sheet identifiers associated with this table."""
        return [
            o.identifier
            for o in self.document_archive.sheets
            if isinstance(self.objects[o.identifier], TNArchives.SheetArchive)
        ]

    def sheet_name(self, sheet_id: int, value: str | None = None) -> str:
        """Return the name of the sheet containing the table."""
        sheet_archive = self.get_object(sheet_id, TNArchives.SheetArchive)
        if value is None:
            return sheet_archive.name
        sheet_archive.name = value
        return value

    def set_table_data(self, table_id: int, data: list) -> None:
        """Set table data."""
        self._table_data[table_id] = data

    def table_info_id(self, table_id: int) -> int:
        """Return the TableInfoArchive ID for a given tale ID."""
        ids = [
            table_info_id
            for table_info_id in self.objects.find_refs("TableInfoArchive")
            if self.get_object(table_info_id, TSTArchives.TableInfoArchive).tableModel.identifier
            == table_id
        ]
        return ids[0]

    def table_info(self, table_id: int = 0, table_info_id: int = 0) -> TSTArchives.TableInfoArchive:
        """Return the TableInfo archive given its ID."""
        if not table_info_id:
            table_info_id = self.table_info_id(table_id)
        return self.get_object(table_info_id, TSTArchives.TableInfoArchive)

    def table_model(self, table_id: int) -> TSTArchives.TableModelArchive:
        """Return the TableModel archive for a table ID."""
        return self.get_object(table_id, TSTArchives.TableModelArchive)

    def table_ids(self, sheet_id: int | None = None) -> list:
        """
        Return table IDs for a sheet or for the whole document.

        If ``sheet_id`` is ``None``, return IDs for all tables.
        """
        table_info_ids = self.find_refs("TableInfoArchive")
        return [
            self.table_info(table_info_id=t_id).tableModel.identifier
            for t_id in table_info_ids
            if (
                sheet_id is None
                or self.table_info(table_info_id=t_id).super.parent.identifier == sheet_id
            )
        ]

    def number_of_rows(self, table_id: int, num_rows: int | None = None) -> int:
        """Return the table row count."""
        if num_rows is not None:
            self.table_model(table_id).number_of_rows = num_rows
        return self.table_model(table_id).number_of_rows

    def number_of_columns(self, table_id: int, num_cols: int | None = None) -> int:
        """Return the table column count."""
        if num_cols is not None:
            self.table_model(table_id).number_of_columns = num_cols
        return self.table_model(table_id).number_of_columns

    def table_name(self, table_id: int, value: str | None = None) -> str:
        """Return the name of a table."""
        if value is None:
            return self.table_model(table_id).table_name
        self.table_model(table_id).table_name = value
        return value

    def table_names(self) -> list[str]:
        """Return the names of all tables in the model."""
        return list(
            chain.from_iterable(
                [[self.table_name(tid) for tid in self.table_ids(sid)] for sid in self.sheet_ids()],
            ),
        )

    def table_name_enabled(self, table_id: int, enabled: bool | None = None) -> bool:
        """Return whether the table name is displayed."""
        if enabled is not None:
            self.table_model(table_id).table_name_enabled = enabled
            return False
        return self.table_model(table_id).table_name_enabled

    def caption_enabled(self, table_id: int, enabled: bool | None = None) -> bool:
        """Return whether the table caption is displayed."""
        table_info = self.table_info(table_id=table_id)
        if enabled is not None:
            table_info.super.caption_hidden = not enabled
            return False
        caption_info_id = table_info.super.caption.identifier
        caption_archive = self.get_object(caption_info_id, TSDArchives.StandinCaptionArchive)
        if caption_archive.DESCRIPTOR.name == "StandinCaptionArchive":
            return False
        return not table_info.super.caption_hidden

    def find_style_id(self, style_substr: str) -> int:
        """Find style id."""
        stylesheet = self.get_object(self.stylesheet_id(), TSSArchives.StylesheetArchive)
        caption_styles = [
            x for x in stylesheet.identifier_to_style_map if style_substr in x.identifier.lower()
        ]
        return caption_styles[0].style.identifier

    def caption_paragraph_style_id(self) -> int:
        """Return the paragraph style identifier for the table caption."""
        style_map = {
            style_id: self.get_object(style_id, TSWPArchives.ParagraphStyleArchive)
            for style_id in self.find_refs("ParagraphStyleArchive")
            if "Caption" in self.get_object(style_id, TSWPArchives.ParagraphStyleArchive).super.name
        }
        return next(iter(style_map.keys()))

    @cache(num_args=0)
    def stylesheet_id(self) -> int:
        """Return the identifier of the table stylesheet."""
        return self.find_refs("StylesheetArchive")[0]

    def set_reference(self, obj: TSPMessages.Reference, ref_id: int) -> None:
        """Set reference."""
        obj.MergeFrom(TSPMessages.Reference(identifier=ref_id))

    def create_path_source_archive(self, table_id: int) -> TSDArchives.PathSourceArchive:
        """Create path source archive."""
        box_size = 100.0
        return TSDArchives.PathSourceArchive(
            horizontalFlip=False,
            verticalFlip=False,
            bezier_path_source=TSDArchives.BezierPathSourceArchive(
                naturalSize=TSPMessages.Size(width=self.table_width(table_id), height=0.0),
                path=TSPMessages.Path(
                    elements=[
                        TSPMessages.Path.Element(
                            type=TSPMessages.Path.ElementType.moveTo,
                            points=[TSPMessages.Point(x=0.0, y=0.0)],
                        ),
                        TSPMessages.Path.Element(
                            type=TSPMessages.Path.ElementType.lineTo,
                            points=[TSPMessages.Point(x=box_size, y=0.0)],
                        ),
                        TSPMessages.Path.Element(
                            type=TSPMessages.Path.ElementType.lineTo,
                            points=[TSPMessages.Point(x=box_size, y=box_size)],
                        ),
                        TSPMessages.Path.Element(
                            type=TSPMessages.Path.ElementType.lineTo,
                            points=[TSPMessages.Point(x=0.0, y=box_size)],
                        ),
                        TSPMessages.Path.Element(
                            type=TSPMessages.Path.ElementType.closeSubpath,
                        ),
                        TSPMessages.Path.Element(
                            type=TSPMessages.Path.ElementType.moveTo,
                            points=[TSPMessages.Point(x=0.0, y=0.0)],
                        ),
                    ],
                ),
            ),
        )

    def create_caption_archive(self, table_id: int) -> None:
        """Create caption archive."""
        table_info = self.table_info(table_id=table_id)
        caption_placement_id, _ = self.objects.create_object_from_dict(
            "CalculationEngine",
            {
                "caption_anchor_location": 1,  # UPPER_LEFT
                "drawable_anchor_location": 7,  # LOWER_LEFT
            },
            TSAArchives.CaptionPlacementArchive,
        )
        caption_info_id, caption_info = self.objects.create_object_from_dict(
            "CalculationEngine",
            {"childInfoKind": "Caption", "placement": {"identifier": caption_placement_id}},
            TSAArchives.CaptionInfoArchive,
        )
        storage_id, _ = self.objects.create_object_from_dict(
            "CalculationEngine",
            {
                "text": ["Caption"],
                "in_document": True,
                "style_sheet": {"identifier": self.stylesheet_id()},
                "table_para_style": {
                    "entries": [
                        {
                            "character_index": 0,
                            "object": {"identifier": self.caption_paragraph_style_id()},
                        },
                    ],
                },
                "table_list_style": {
                    "entries": [
                        {
                            "character_index": 0,
                            "object": {"identifier": self.find_style_id("liststyle")},
                        },
                    ],
                },
                "table_para_starts": {"entries": [{"character_index": 0, "first": 0, "second": 0}]},
                "table_para_bidi": {"entries": [{"character_index": 0, "first": 0, "second": 0}]},
                "table_drop_cap_style": {"entries": [{"character_index": 0}]},
            },
            TSWPArchives.StorageArchive,
        )
        for object_id in [
            storage_id,
            self.find_style_id("caption"),
            self.find_style_id("liststyle"),
            self.caption_paragraph_style_id(),
        ]:
            self.add_component_reference(
                object_id,
                location="CalculationEngine",
                component_id=self.stylesheet_id(),
            )
        caption_info.super.MergeFrom(
            TSWPArchives.ShapeInfoArchive(
                is_text_box=True,
                owned_storage=TSPMessages.Reference(identifier=storage_id),
                deprecated_storage=TSPMessages.Reference(identifier=storage_id),
                super=TSDArchives.ShapeArchive(
                    super=self.create_drawable(self.table_info_id(table_id), 0, 500, flags=1),
                    style={"identifier": self.find_style_id("caption")},
                    strokePatternOffsetDistance=0.0,
                    pathsource=self.create_path_source_archive(table_id),
                ),
            ),
        )

        self.set_reference(table_info.super.caption, caption_info_id)
        component = self.metadata_component_info(self.calc_engine_id())
        component.object_uuid_map_entries.append(
            TSPArchiveMessages.ObjectUUIDMapEntry(
                identifier=caption_info_id,
                uuid=NumbersUUID().protobuf2,
            ),
        )

    def caption_text(self, table_id: int, caption: str | None = None) -> str:
        """Return the text displayed in the table caption."""
        table_info = self.table_info(table_id=table_id)
        caption_info_id = table_info.super.caption.identifier
        caption_archive = self.get_object(caption_info_id, TSAArchives.CaptionInfoArchive)

        if caption_archive.DESCRIPTOR.name == "StandinCaptionArchive":
            if caption is None:
                return "Caption"
            self.create_caption_archive(table_id)
            caption_archive = self.get_object(
                table_info.super.caption.identifier,
                TSAArchives.CaptionInfoArchive,
            )
        else:
            caption_archive = cast("TSAArchives.CaptionInfoArchive", caption_archive)

        storage_archive = self.get_object(
            caption_archive.super.owned_storage.identifier,
            TSWPArchives.StorageArchive,
        )
        if caption is not None:
            storage_archive.text.clear()
            storage_archive.text.append(caption)
            return "Caption"
        if len(storage_archive.text) == 0:
            return "Caption"
        return storage_archive.text[0]

    @cache(num_args=0)
    def custom_format_map(self) -> dict[str, TSKArchives.CustomFormatArchive]:
        """Return the custom-format mapping for the document."""
        custom_format_list = self.get_object(
            self.document_archive.super.custom_format_list.identifier,
            TSKArchives.CustomFormatListArchive,
        )
        return {
            NumbersUUID(u).hex: custom_format_list.custom_formats[i]
            for i, u in enumerate(custom_format_list.uuids)
        }

    @cache(num_args=2)
    def table_format(self, table_id: int, key: int) -> TSKArchives.FormatStructArchive:
        """Return the format associated with a format ID for a particular table."""
        return self._table_formats.lookup_value(table_id, key).format

    @cache(num_args=3)
    def format_archive(
        self,
        table_id: int,
        format_type: FormattingType,
        formatting: Formatting,
    ) -> int:
        """Create a table format from a Formatting spec and return the table format ID."""
        attrs = {x: getattr(formatting, x) for x in ALLOWED_FORMATTING_PARAMETERS[format_type]}
        attrs["format_type"] = FORMAT_TYPE_MAP[format_type]

        format_archive = TSKArchives.FormatStructArchive(**attrs)
        return self._table_formats.lookup_key(table_id, format_archive)

    def cell_popup_model(self, parent_id: int, formatting: Formatting) -> int:
        """Return the popup-menu model for a cell."""
        tsce_items: list[dict[str, Any]] = [{"cell_value_type": "NIL_TYPE"}]
        for item in formatting.popup_values:
            if isinstance(item, str):
                tsce_items.append(
                    {
                        "cell_value_type": "STRING_TYPE",
                        "string_value": {
                            "value": item,
                            "format": {"format_type": FormatType.TEXT},
                        },
                    },
                )
            else:
                tsce_items.append(
                    {
                        "cell_value_type": "NUMBER_TYPE",
                        "number_value": {
                            "value": item,
                            "format": {"format_type": FormatType.DECIMAL},
                        },
                    },
                )
        popup_menu_id, _ = self.objects.create_object_from_dict(
            f"Index/Tables/DataList-{parent_id}",
            {"tsce_item": tsce_items},
            TSTArchives.PopUpMenuModel,
            True,  # noqa: FBT003
        )
        return popup_menu_id

    def control_cell_archive(
        self,
        table_id: int,
        format_type: FormattingType,
        formatting: Formatting,
    ) -> int:
        """Create control cell archive from a Formatting spec and return the table format ID."""
        if format_type == FormattingType.TICKBOX:
            cell_spec = TSTArchives.CellSpecArchive(interaction_type=CellInteractionType.TOGGLE)
        elif format_type == FormattingType.RATING:
            cell_spec = TSTArchives.CellSpecArchive(
                interaction_type=CellInteractionType.RATING,
                range_control_min=0.0,
                range_control_max=5.0,
                range_control_inc=1.0,
            )
        elif format_type == FormattingType.SLIDER:
            cell_spec = TSTArchives.CellSpecArchive(
                interaction_type=CellInteractionType.SLIDER,
                range_control_min=formatting.minimum,
                range_control_max=formatting.maximum,
                range_control_inc=formatting.increment,
            )
        elif format_type == FormattingType.STEPPER:
            cell_spec = TSTArchives.CellSpecArchive(
                interaction_type=CellInteractionType.STEPPER,
                range_control_min=formatting.minimum,
                range_control_max=formatting.maximum,
                range_control_inc=formatting.increment,
            )
        else:  # POPUP
            popup_id = self.cell_popup_model(self._control_specs.id(table_id), formatting)
            cell_spec = TSTArchives.CellSpecArchive(
                interaction_type=CellInteractionType.POPUP,
                chooser_control_popup_model=TSPMessages.Reference(identifier=popup_id),
                chooser_control_start_w_first=not (formatting.allow_none),
            )
        return self._control_specs.lookup_key(table_id, cell_spec)

    def add_custom_decimal_format_archive(self, formatting: CustomFormatting) -> None:
        """Create a custom format from the format spec."""
        integer_format = formatting.integer_format
        decimal_format = formatting.decimal_format
        num_integers = formatting.num_integers
        num_decimals = formatting.num_decimals
        show_thousands_separator = formatting.show_thousands_separator

        if num_integers == 0:
            format_string = ""
        elif integer_format == PaddingType.NONE:
            format_string = "#" * num_integers
        else:
            format_string = "0" * num_integers
        if num_integers > 6:  # noqa: PLR2004
            format_string = re.sub(r"(...)(...)$", r",\1,\2", format_string)
        elif num_integers > 3:
            format_string = re.sub(r"(...)$", r",\1", format_string)
        if num_decimals > 0:
            if decimal_format == PaddingType.NONE:
                format_string += "." + "#" * num_decimals
            else:
                format_string += "." + "0" * num_decimals

        min_integer_width = (
            num_integers if num_integers > 0 and integer_format != PaddingType.NONE else 0
        )
        num_nonspace_decimal_digits = num_decimals if decimal_format == PaddingType.ZEROS else 0
        num_nonspace_integer_digits = num_integers if integer_format == PaddingType.ZEROS else 0
        index_from_right_last_integer = num_decimals + 1 if num_integers > 0 else num_decimals
        # Empirically correct:
        if index_from_right_last_integer == 1:
            index_from_right_last_integer = 0
        elif index_from_right_last_integer == 0:
            index_from_right_last_integer = 1
        decimal_width = num_decimals if decimal_format == PaddingType.SPACES else 0
        is_complex = "0" in format_string and (
            min_integer_width > 0 or num_nonspace_decimal_digits == 0
        )

        format_archive = TSKArchives.CustomFormatArchive(
            name=formatting.name,
            format_type_pre_bnc=FormatType.CUSTOM_NUMBER,
            format_type=FormatType.CUSTOM_NUMBER,
            default_format=TSKArchives.FormatStructArchive(
                contains_integer_token=num_integers > 0,
                custom_format_string=format_string,
                decimal_width=decimal_width,
                format_type=FormatType.CUSTOM_NUMBER,
                fraction_accuracy=0xFFFFFFFD,
                index_from_right_last_integer=index_from_right_last_integer,
                is_complex=is_complex,
                min_integer_width=min_integer_width,
                num_hash_decimal_digits=0,
                num_nonspace_decimal_digits=num_nonspace_decimal_digits,
                num_nonspace_integer_digits=num_nonspace_integer_digits,
                requires_fraction_replacement=False,
                scale_factor=1.0,
                show_thousands_separator=show_thousands_separator and num_integers > 0,
                total_num_decimal_digits=decimal_width,
                use_accounting_style=False,
            ),
        )
        self.add_custom_format_archive(formatting, format_archive)

    def add_custom_datetime_format_archive(self, formatting: CustomFormatting) -> None:
        """Add custom datetime format archive."""
        format_archive = TSKArchives.CustomFormatArchive(
            name=formatting.name,
            format_type_pre_bnc=FormatType.CUSTOM_DATE,
            format_type=FormatType.CUSTOM_DATE,
            default_format=TSKArchives.FormatStructArchive(
                custom_format_string=formatting.format,
                format_type=FormatType.CUSTOM_DATE,
            ),
        )
        self.add_custom_format_archive(formatting, format_archive)

    def add_custom_format_archive(
        self,
        formatting: CustomFormatting,
        format_archive: TSKArchives.CustomFormatArchive,
    ) -> None:
        """Add custom format archive."""
        format_uuid = NumbersUUID().protobuf2
        formatting_name = formatting.name or ""
        self.custom_formats[formatting_name] = formatting
        self._custom_format_archives[formatting_name] = format_archive
        self._custom_format_uuids[formatting_name] = format_uuid

        custom_format_list = self.get_object(
            self.document_archive.super.custom_format_list.identifier,
            TSKArchives.CustomFormatListArchive,
        )
        custom_format_list.custom_formats.append(format_archive)
        custom_format_list.uuids.append(format_uuid)

    def custom_format_id(self, table_id: int, formatting: CustomFormatting) -> int:
        """Look up the custom format and return the format ID for the table."""
        format_type = CUSTOM_FORMAT_TYPE_MAP[formatting.type]
        format_uuid = self._custom_format_uuids[formatting.name or ""]
        custom_format = TSKArchives.FormatStructArchive(
            format_type=format_type,
            custom_uid=TSPMessages.UUID(lower=format_uuid.lower, upper=format_uuid.upper),
        )
        return self._table_formats.lookup_key(table_id, custom_format)

    def add_custom_text_format_archive(self, formatting: CustomFormatting) -> None:
        """Add custom text format archive."""
        format_string = formatting.format.replace("%s", CUSTOM_TEXT_PLACEHOLDER)
        format_archive = TSKArchives.CustomFormatArchive(
            name=formatting.name,
            format_type_pre_bnc=FormatType.CUSTOM_TEXT,
            format_type=FormatType.CUSTOM_TEXT,
            default_format=TSKArchives.FormatStructArchive(
                custom_format_string=format_string,
                format_type=FormatType.CUSTOM_TEXT,
            ),
        )
        self.add_custom_format_archive(formatting, format_archive)

    @cache(num_args=2)
    def table_style(self, table_id: int, key: int) -> TSTArchives.CellStyleArchive:
        """Return the style associated with a style ID for a particular table."""
        style_entry = self.table_styles.lookup_value(table_id, key)
        return self.get_object(style_entry.reference.identifier, TSTArchives.CellStyleArchive)

    @cache(num_args=2)
    def table_string(self, table_id: int, key: int) -> str:
        """Return the string associated with a string ID for a particular table."""
        try:
            return self._table_strings.lookup_value(table_id, key).string
        except KeyError:
            return ""

    def init_table_strings(self, table_id: int) -> None:
        """Cache table strings reference and delete all existing keys/values."""
        self._table_strings.init(table_id)

    def table_string_key(self, table_id: int, value: str) -> int:
        """
        Return the key for a string in a table's string list.

        If the string is absent, allocate an entry with the next available key.
        """
        return self._table_strings.lookup_key(table_id, value)

    @cache(num_args=0)
    def owner_id_map(self) -> dict[int, str]:
        """
        Map owner IDs to their UUIDs.

        Return the mapping from integer owner IDs to 128-bit UUID strings.
        """
        calc_engine = self.calc_engine()
        if calc_engine is None:
            return {}

        owner_id_map = {}
        for e in calc_engine.dependency_tracker.owner_id_map.map_entry:
            owner_id_map[e.internal_owner_id] = NumbersUUID(e.owner_id).hex
        return owner_id_map

    def calculate_table_uuid_map(self) -> None:
        # Each Table Model has a UUID which is used in references to the table.

        # For haunted owner archive types, map formula_owner_uids to their base_owner_uids
        """Calculate table uuid map."""
        haunted_owner_ids = [
            dependency_id
            for dependency_id in self.find_refs("FormulaOwnerDependenciesArchive")
            if self.get_object(
                dependency_id,
                TSCEArchives.FormulaOwnerDependenciesArchive,
            ).owner_kind
            == OwnerKind.HAUNTED_OWNER
        ]
        if len(haunted_owner_ids) == 0:
            # Some older documents (see issue-18) do not use FormulaOwnerDependenciesArchive
            self._table_id_to_base_id = {}
            return

        formula_owner_to_base_owner_map = {}
        for obj_id in haunted_owner_ids:
            archive = self.get_object(obj_id, TSCEArchives.FormulaOwnerDependenciesArchive)
            formula_owner_uid = uuid_to_hex(archive.formula_owner_uid)
            base_owner_uid = uuid_to_hex(archive.base_owner_uid)
            formula_owner_to_base_owner_map[formula_owner_uid] = base_owner_uid

        # Map table IDs to the base_owner_uids of the formula owners that match
        # the table model's haunted owner
        self._table_id_to_base_id = {
            table_id: formula_owner_to_base_owner_map.get(
                uuid_to_hex(self.table_model(table_id).haunted_owner.owner_uid),
            )
            for table_id in self.table_ids()
        }
        self._table_base_id_to_formula_owner_id = {
            uuid_to_hex(
                self.get_object(
                    obj_id,
                    TSCEArchives.FormulaOwnerDependenciesArchive,
                ).base_owner_uid,
            ): obj_id
            for obj_id in haunted_owner_ids
        }

    @cache()
    def table_base_id(self, table_id: int) -> str | None:
        """Find the base UUID for a table."""
        # Table can be empty if the document does not use FormulaOwnerDependenciesArchive
        return self._table_id_to_base_id.get(table_id)

    def add_merge_range(
        self,
        table_id: int,
        row_start: int,
        row_end: int,
        col_start: int,
        col_end: int,
    ) -> None:
        """Add merge range."""
        debug(
            "Add merge: table_id=%d, [%d,%d]->[%d,%d]",
            table_id,
            row_start,
            row_end,
            col_start,
            col_end,
        )
        size = (row_end - row_start + 1, col_end - col_start + 1)
        for row in range(row_start, row_end + 1):
            for col in range(col_start, col_end + 1):
                self._merge_cells[table_id].add_reference(
                    row,
                    col,
                    (row_start, col_start, row_end, col_end),
                )
        self._merge_cells[table_id].add_anchor(row_start, col_start, size)

    @cache()
    def calculate_merges_using_formula_stores(self, table_id: int) -> int:
        """Calculate merges using formula stores."""

        def get_range_end(
            archive: ASTColonTractArchive.ASTColonTractAbsoluteRangeArchive,
        ) -> int:
            # range_end is optional
            """Return the end coordinate of a range, or its start if unset."""
            return archive.range_end if archive.HasField("range_end") else archive.range_begin

        table_model = self.table_model(table_id)
        formulas = table_model.merge_owner.formula_store.formulas
        if len(formulas) == 0:
            debug("table=%s: no formula store merges", self.table_name(table_id))
            return 0

        merge_count = 0
        for formula in formulas:
            node = formula.formula.AST_node_array.AST_node[0]
            if node.AST_node_type != COLON_TRACT_NODE:
                continue
            merge_count += 1
            self.add_merge_range(
                table_id,
                node.AST_colon_tract.absolute_row[0].range_begin,
                get_range_end(node.AST_colon_tract.absolute_row[0]),
                node.AST_colon_tract.absolute_column[0].range_begin,
                get_range_end(node.AST_colon_tract.absolute_column[0]),
            )
        debug("table=%s: %d formula store merges found", self.table_name(table_id), merge_count)
        return merge_count

    @cache()
    def calculate_merges_using_dependency_archives(self, table_id: int) -> int:
        """Extract all the merge cell ranges for the Table."""
        owner_id_map = self.owner_id_map()
        table_base_id = self.table_base_id(table_id)

        formula_table_ids = self.find_refs("FormulaOwnerDependenciesArchive")
        merge_count = 0
        for formula_id in formula_table_ids:
            dependencies = self.get_object(formula_id, TSCEArchives.FormulaOwnerDependenciesArchive)
            if dependencies.owner_kind != OwnerKind.MERGE_OWNER:
                continue
            for record in dependencies.range_dependencies.back_dependency:
                to_owner_id = record.internal_range_reference.owner_id
                if owner_id_map[to_owner_id] == table_base_id:
                    record_range = record.internal_range_reference.range
                    merge_count += 1
                    self.add_merge_range(
                        table_id,
                        record_range.top_left_row,
                        record_range.bottom_right_row,
                        record_range.top_left_column,
                        record_range.bottom_right_column,
                    )
        debug(
            "table=%s: %d dependency archive merges found",
            self.table_name(table_id),
            merge_count,
        )
        return merge_count

    @cache()
    def calculate_merges_using_region_map(self, table_id: int) -> None:
        """Calculate merges using region map."""
        base_data_store = self.table_model(table_id).base_data_store
        if base_data_store.merge_region_map.identifier == 0:
            debug("table=%s: no merge_region_map", self.table_name(table_id))
            return

        merge_count = 0
        cell_ranges = self.get_object(
            base_data_store.merge_region_map.identifier,
            TSTArchives.MergeRegionMapArchive,
        )
        for cell_range in cell_ranges.cell_range:
            (col_start, row_start) = (
                cell_range.origin.packedData >> 16,
                cell_range.origin.packedData & 0xFFFF,
            )
            (num_columns, num_rows) = (
                cell_range.size.packedData >> 16,
                cell_range.size.packedData & 0xFFFF,
            )
            row_end = row_start + num_rows - 1
            col_end = col_start + num_columns - 1
            merge_count += 1
            self.add_merge_range(table_id, row_start, row_end, col_start, col_end)
        debug(
            "table=%s: %d merge_region_map merges found",
            self.table_name(table_id),
            merge_count,
        )

    def merge_cells(self, table_id: int) -> MergeCells:
        """Merge cells."""
        if self.calculate_merges_using_formula_stores(table_id) > 0:
            return self._merge_cells[table_id]
        if self.calculate_merges_using_dependency_archives(table_id) > 0:
            return self._merge_cells[table_id]
        self.calculate_merges_using_region_map(table_id)
        return self._merge_cells[table_id]

    def table_id_to_sheet_id(self, table_id: int) -> int:
        """Return the sheet identifier containing a table."""
        return next(
            sheet_id for sheet_id in self.sheet_ids() if table_id in self.table_ids(sheet_id)
        )

    @cache()
    def table_uuids_to_id(self, table_uuid: str) -> int:
        """Map table UUIDs to their model identifiers."""
        return next(
            (
                table_id
                for sheet_id in self.sheet_ids()
                for table_id in self.table_ids(sheet_id)
                if table_uuid == self.table_base_id(table_id)
            ),
            0,
        )

    def node_to_ref(  # noqa: C901
        self,
        table_id: int,
        row: int,
        col: int,
        node: ASTNodeArchive,
    ) -> CellRange:
        """Convert a formula node into a cell reference."""

        def resolve_range(
            is_absolute: bool,
            absolute_list: Sequence[ASTColonTractArchive.ASTColonTractAbsoluteRangeArchive],
            relative_list: Sequence[ASTColonTractArchive.ASTColonTractRelativeRangeArchive],
            offset: int,
            max_val: int,
        ) -> int:
            """Resolve range."""
            if is_absolute:
                return absolute_list[0].range_begin
            if not relative_list and absolute_list[0].range_begin == max_val:
                return max_val
            return offset + relative_list[0].range_begin

        def resolve_range_end(
            is_absolute: bool,
            absolute_list: Sequence[ASTColonTractArchive.ASTColonTractAbsoluteRangeArchive],
            relative_list: Sequence[ASTColonTractArchive.ASTColonTractRelativeRangeArchive],
            offset: int,
            max_val: int,
        ) -> int:
            """Resolve range end."""
            if is_absolute:
                return range_end(absolute_list[0])
            if not relative_list and range_end(absolute_list[0]) == max_val:
                return max_val
            return offset + range_end(relative_list[0])

        if node.HasField("AST_cross_table_reference_extra_info"):
            table_uuid = NumbersUUID(node.AST_cross_table_reference_extra_info.table_id).hex
            to_table_id = self.table_uuids_to_id(table_uuid)
        else:
            to_table_id = table_id

        if node.HasField("AST_colon_tract"):
            row_begin = resolve_range(
                node.AST_sticky_bits.begin_row_is_absolute,
                node.AST_colon_tract.absolute_row,
                node.AST_colon_tract.relative_row,
                row,
                0x7FFFFFFF,
            )

            row_end = resolve_range_end(
                node.AST_sticky_bits.end_row_is_absolute,
                node.AST_colon_tract.absolute_row,
                node.AST_colon_tract.relative_row,
                row,
                0x7FFFFFFF,
            )

            col_begin = resolve_range(
                node.AST_sticky_bits.begin_column_is_absolute,
                node.AST_colon_tract.absolute_column,
                node.AST_colon_tract.relative_column,
                col,
                0x7FFF,
            )

            col_end = resolve_range_end(
                node.AST_sticky_bits.end_column_is_absolute,
                node.AST_colon_tract.absolute_column,
                node.AST_colon_tract.relative_column,
                col,
                0x7FFF,
            )

            return CellRange(
                _model=self,
                row_start=None if row_begin == 0x7FFFFFFF else row_begin,  # noqa: PLR2004
                row_end=None if row_end == 0x7FFFFFFF else row_end,  # noqa: PLR2004
                col_start=None if col_begin == 0x7FFF else col_begin,  # noqa: PLR2004
                col_end=None if col_end == 0x7FFF else col_end,  # noqa: PLR2004
                row_start_is_abs=node.AST_sticky_bits.begin_row_is_absolute,
                row_end_is_abs=node.AST_sticky_bits.end_row_is_absolute,
                col_start_is_abs=node.AST_sticky_bits.begin_column_is_absolute,
                col_end_is_abs=node.AST_sticky_bits.end_column_is_absolute,
                from_table_id=table_id,
                to_table_id=to_table_id,
            )

        row = node.AST_row.row if node.AST_row.absolute else row + node.AST_row.row
        col = node.AST_column.column if node.AST_column.absolute else col + node.AST_column.column
        if node.HasField("AST_row") and not node.HasField("AST_column"):
            return CellRange(
                _model=self,
                row_start=row,
                row_start_is_abs=node.AST_row.absolute,
                from_table_id=table_id,
                to_table_id=to_table_id,
            )

        if node.HasField("AST_column") and not node.HasField("AST_row"):
            return CellRange(
                _model=self,
                col_start=col,
                col_start_is_abs=node.AST_column.absolute,
                from_table_id=table_id,
                to_table_id=to_table_id,
            )

        return CellRange(
            _model=self,
            row_start=row,
            col_start=col,
            row_start_is_abs=node.AST_row.absolute,
            col_start_is_abs=node.AST_column.absolute,
            from_table_id=table_id,
            to_table_id=to_table_id,
        )

    @cache()
    def formula_ast(
        self,
        table_id: int,
    ) -> dict[int, Sequence[TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive]]:
        """Return the parsed formula AST for a table."""
        bds = self.table_model(table_id).base_data_store
        formula_table = self.get_object(bds.formula_table.identifier, TSTArchives.TableDataList)
        formulas = {}
        for formula in formula_table.entries:
            formulas[formula.key] = formula.formula.AST_node_array.AST_node
        return formulas

    @cache()
    def storage_buffers(self, table_id: int) -> dict[int, list[bytes | None]]:
        # Keyed by absolute row index: rows without data have no rowInfo
        # so the position in rowInfos is not the row number.
        """Return the cell storage buffers for a table."""
        buffers = {}
        bds = self.table_model(table_id).base_data_store
        tile_size = bds.tiles.tile_size or DEFAULT_TILE_SIZE
        for tile_ref in bds.tiles.tiles:
            tile = self.get_object(tile_ref.tile.identifier, TSTArchives.Tile)
            if not tile.last_saved_in_BNC:
                msg = "Pre-BNC storage is unsupported"
                raise UnsupportedError(msg)
            for r in tile.rowInfos:
                buffer = get_storage_buffers_for_row(
                    r.cell_storage_buffer,
                    r.cell_offsets,
                    self.number_of_columns(table_id),
                    r.has_wide_offsets,
                )
                buffers[tile_ref.tileid * tile_size + r.tile_row_index] = buffer
        return buffers

    @cache(num_args=3)
    def storage_buffer(self, table_id: int, row: int, col: int) -> bytes | None:
        """Return the storage buffer for a cell."""
        row_buffers = self.storage_buffers(table_id).get(row)
        if row_buffers is None or col >= len(row_buffers):
            return None
        return row_buffers[col]

    def recalculate_row_headers(self, table_id: int, data: list) -> None:
        """Recalculate the table row-header metadata."""
        current_row_heights = {}
        for row in range(self.number_of_rows(table_id)):
            current_row_heights[row] = self.row_height(table_id, row)

        base_data_store = self.table_model(table_id).base_data_store
        buckets = self.get_object(
            base_data_store.rowHeaders.buckets[0].identifier,
            TSTArchives.HeaderStorageBucket,
        )
        buckets.headers.clear()

        for row, cells in enumerate(data):
            num_cols = len(data) - sum(isinstance(x, MergedCell) for x in cells)
            height = current_row_heights[row]
            header = TSTArchives.HeaderStorageBucket.Header(
                index=row,
                numberOfCells=num_cols,
                size=height,
                hidingState=0,
            )
            buckets.headers.append(header)

    def recalculate_column_headers(self, table_id: int, data: list) -> None:
        """Recalculate the table column-header metadata."""
        current_column_widths = {}
        for col in range(self.number_of_columns(table_id)):
            current_column_widths[col] = self.col_width(table_id, col)

        base_data_store = self.table_model(table_id).base_data_store
        buckets = self.get_object(
            base_data_store.columnHeaders.identifier,
            TSTArchives.HeaderStorageBucket,
        )
        buckets.headers.clear()
        # Transpose data to get columns
        col_data = [list(x) for x in zip(*data, strict=False)]

        for col, cells in enumerate(col_data):
            num_rows = len(cells) - sum(isinstance(x, MergedCell) for x in cells)
            width = current_column_widths[col]
            header = TSTArchives.HeaderStorageBucket.Header(
                index=col,
                numberOfCells=num_rows,
                size=width,
                hidingState=0,
            )
            buckets.headers.append(header)

    def recalculate_merged_cells(self, table_id: int) -> None:
        """Recalculate merged-cell ranges for the table."""
        merge_cells = self.merge_cells(table_id)

        merge_map_id, merge_map = self.objects.create_object_from_dict(
            "CalculationEngine",
            {},
            TSTArchives.MergeRegionMapArchive,
        )

        merge_cells = self.merge_cells(table_id)
        for row_col in merge_cells.merge_cells():
            size = merge_cells.size(row_col)
            cell_id = TSTArchives.CellID(packedData=row_col[1] << 16 | row_col[0])
            table_size = TSTArchives.TableSize(packedData=size[1] << 16 | size[0])
            cell_range = TSTArchives.CellRange(origin=cell_id, size=table_size)
            merge_map.cell_range.append(cell_range)

        base_data_store = self.table_model(table_id).base_data_store
        self.set_reference(base_data_store.merge_region_map, merge_map_id)

    def recalculate_row_info(
        self,
        _table_id: int,
        data: list,
        tile_row_offset: int,
        row: int,
    ) -> TSTArchives.TileRowInfo:
        """Recalculate the row metadata for the table."""
        row_info = TSTArchives.TileRowInfo()
        row_info.storage_version = 5
        row_info.tile_row_index = row - tile_row_offset
        row_info.cell_count = 0
        cell_storage = b""

        offsets = [-1] * len(data[0])
        current_offset = 0

        for col in range(len(data[row])):
            buffer = data[row][col]._to_buffer()
            if buffer is not None:
                cell_storage += buffer
                # Always use wide offsets
                offsets[col] = current_offset >> 2
                current_offset += len(buffer)

                row_info.cell_count += 1

        row_info.cell_offsets = pack(f"<{len(offsets)}h", *offsets)
        row_info.cell_offsets_pre_bnc = DEFAULT_PRE_BNC_BYTES
        row_info.cell_storage_buffer = cell_storage
        row_info.cell_storage_buffer_pre_bnc = DEFAULT_PRE_BNC_BYTES
        row_info.has_wide_offsets = True
        return row_info

    @cache()
    def metadata_component_info(
        self,
        reference: str | int | None = None,
    ) -> TSPArchiveMessages.ComponentInfo:
        """Return the component info of an object in the document metadata given it's name or ID."""
        component_map = {c.identifier: c for c in self.package_metadata.components}
        if isinstance(reference, str):
            component_ids = [
                x for x, c in component_map.items() if c.preferred_locator == reference
            ]
        else:
            component_ids = [x for x, c in component_map.items() if c.identifier == reference]
        return cast("TSPArchiveMessages.ComponentInfo", component_map[component_ids[0]])

    def add_component_metadata(self, object_id: int, parent: str, locator: str) -> None:
        """Add a new ComponentInfo record to the parent object in the document metadata."""
        if locator is not None:
            component_info = TSPArchiveMessages.ComponentInfo(
                identifier=object_id,
                is_stored_outside_object_archive=False,
                document_read_version=[2, 0, 0],
                document_write_version=[2, 0, 0],
                save_token=1,
                locator=locator.format(object_id),
                preferred_locator=re.sub(r"\-\d+.*", "", locator),
            )
        else:
            component_info = TSPArchiveMessages.ComponentInfo(
                identifier=object_id,
                is_stored_outside_object_archive=False,
                document_read_version=[2, 0, 0],
                document_write_version=[2, 0, 0],
                save_token=1,
                preferred_locator=parent,
            )
        self.package_metadata.components.append(component_info)
        self.add_component_reference(object_id, location=parent)

    def add_component_reference(
        self,
        object_id: int,
        location: str | int | None = None,
        component_id: int | None = None,
        is_weak: bool = False,
    ) -> None:
        """Add an external reference to an object in a metadata component."""
        component = self.metadata_component_info(location or component_id)
        params: dict[str, Any] = (
            {"object_identifier": object_id, "component_identifier": component_id}
            if component_id is not None
            else {"component_identifier": object_id}
        )
        if is_weak:
            params["is_weak"] = True
        component.external_references.append(
            TSPArchiveMessages.ComponentExternalReference(**params),
        )

    def recalculate_table_data(self, table_id: int, data: list) -> None:
        """Recalculate table cell data and related metadata."""
        table_model = self.table_model(table_id)
        table_model.number_of_rows = len(data)
        table_model.number_of_columns = len(data[0])

        self.init_table_strings(table_id)
        self.recalculate_row_headers(table_id, data)
        self.recalculate_column_headers(table_id, data)
        self.recalculate_merged_cells(table_id)
        self.update_paragraph_styles()
        self.update_cell_styles(table_id, data)
        self.update_cell_borders(table_id, data)

        self.objects.remove_unreferenced_objects()
        self.update_table_metadata_ranges(table_id, len(data), len(data[0]))

        tile_idx = 0
        max_tile_idx = len(data) >> 8
        base_data_store = self.table_model(table_id).base_data_store
        base_data_store.tiles.ClearField("tiles")
        if len(data[0]) > MAX_TILE_SIZE:
            base_data_store.tiles.should_use_wide_rows = True

        while tile_idx <= max_tile_idx:
            row_start = tile_idx * MAX_TILE_SIZE
            if (len(data) - row_start) > MAX_TILE_SIZE:
                num_rows = MAX_TILE_SIZE
                row_end = row_start + MAX_TILE_SIZE
            else:
                num_rows = len(data) - row_start
                row_end = row_start + num_rows

            tile_dict = {
                "maxColumn": 0,
                "maxRow": 0,
                "numCells": 0,
                "numrows": num_rows,
                "storage_version": 5,
                "rowInfos": [],
                "last_saved_in_BNC": True,
                "should_use_wide_rows": True,
            }
            tile_id, tile = self.objects.create_object_from_dict(
                "Index/Tables/Tile-{}",
                tile_dict,
                TSTArchives.Tile,
            )
            for row in range(row_start, row_end):
                row_info = self.recalculate_row_info(table_id, data, row_start, row)
                tile.rowInfos.append(row_info)

            tile_ref = TSTArchives.TileStorage.Tile()
            tile_ref.tileid = tile_idx
            tile_ref.tile.MergeFrom(TSPMessages.Reference(identifier=tile_id))
            base_data_store.tiles.tiles.append(tile_ref)
            base_data_store.tiles.tile_size = MAX_TILE_SIZE

            self.add_component_metadata(tile_id, "CalculationEngine", "Tables/Tile-{}")

            tile_idx += 1

        self.objects.update_object_file_store()

    def create_string_table(self) -> tuple[int, TSTArchives.TableDataList]:
        """Create string table."""
        table_strings_id, table_strings = self.objects.create_object_from_dict(
            "Index/Tables/DataList-{}",
            {"listType": TSTArchives.TableDataList.ListType.STRING, "nextListID": 1},
            TSTArchives.TableDataList,
        )
        self.add_component_metadata(table_strings_id, "CalculationEngine", "Tables/DataList-{}")
        return table_strings_id, table_strings

    def table_height(self, table_id: int) -> int:
        """Return the height of a table in points."""
        height = 0.0
        for row in range(self.number_of_rows(table_id)):
            height += self.row_height(table_id, row)
        return floor(height)

    def row_height(self, table_id: int, row: int, height: int | None = None) -> int:
        """Return or update the height of a table row."""
        if height is not None:
            if table_id not in self._row_heights:
                self._row_heights[table_id] = {}
            self._row_heights[table_id][row] = height
            return height

        if table_id in self._row_heights and row in self._row_heights[table_id]:
            return self._row_heights[table_id][row]

        table_model = self.table_model(table_id)
        bds = self.table_model(table_id).base_data_store
        buckets = self.get_object(
            bds.rowHeaders.buckets[0].identifier,
            TSTArchives.HeaderStorageBucket,
        ).headers
        bucket_map = {x.index: x for x in buckets}
        if row in bucket_map and bucket_map[row].size != 0.0:
            height = round(bucket_map[row].size)
        else:
            height = round(table_model.default_row_height)

        return height

    def table_width(self, table_id: int) -> int:
        """Return the width of a table in points."""
        width = 0.0
        for row in range(self.number_of_columns(table_id)):
            width += self.col_width(table_id, row)
        return round(width)

    def col_width(self, table_id: int, col: int, width: int | None = None) -> int:
        """Return or update the width of a table column."""
        if width is not None:
            if table_id not in self._col_widths:
                self._col_widths[table_id] = {}
            self._col_widths[table_id][col] = width
            return width

        if table_id in self._col_widths and col in self._col_widths[table_id]:
            return self._col_widths[table_id][col]

        table_model = self.table_model(table_id)
        bds = self.table_model(table_id).base_data_store
        buckets = self.get_object(
            bds.columnHeaders.identifier,
            TSTArchives.HeaderStorageBucket,
        ).headers
        bucket_map = {x.index: x for x in buckets}
        if col in bucket_map and bucket_map[col].size != 0.0:
            width = round(bucket_map[col].size)
        else:
            width = round(table_model.default_column_width)
        return width

    def num_header_rows(self, table_id: int, num_headers: int | None = None) -> int:
        """Return/set the number of header rows."""
        table_model = self.table_model(table_id)
        if num_headers is not None:
            table_model.number_of_header_rows = num_headers
        return table_model.number_of_header_rows

    def num_header_cols(self, table_id: int, num_headers: int | None = None) -> int:
        """Return/set the number of header columns."""
        table_model = self.table_model(table_id)
        if num_headers is not None:
            table_model.number_of_header_columns = num_headers
        return table_model.number_of_header_columns

    def table_coordinates(self, table_id: int) -> tuple[float, float]:
        """Return the coordinates and dimensions of a table."""
        table_info = self.table_info(table_id=table_id)
        return (
            table_info.super.geometry.position.x,
            table_info.super.geometry.position.y,
        )

    def is_a_pivot_table(self, table_id: int) -> bool:
        """Table is a pivot table."""
        table_info = self.table_info(table_id=table_id)
        return table_info.is_a_pivot_table

    def last_table_offset(self, sheet_id: int) -> float:
        """Y offset of the last table in a sheet."""
        table_id = self.table_ids(sheet_id)[-1]
        y_offset = next(
            self.table_info(table_id=x).super.geometry.position.y
            for x in self.table_ids(sheet_id)
            if x == table_id
        )

        return self.table_height(table_id) + y_offset

    def create_drawable(  # noqa: PLR0917
        self,
        sheet_id: int,
        x: float | None,
        y: float | None,
        flags: int = 3,
        height: float = 231.0,
        width: float = 494.0,
    ) -> TSDArchives.DrawableArchive:
        """Create a DrawableArchive for a new table in a sheet."""
        table_x = x if x is not None else 0.0
        table_y = y if y is not None else self.last_table_offset(sheet_id) + DEFAULT_TABLE_OFFSET
        return TSDArchives.DrawableArchive(
            parent=TSPMessages.Reference(identifier=sheet_id),
            geometry=TSDArchives.GeometryArchive(
                angle=0.0,
                flags=flags,
                position=TSPMessages.Point(x=table_x, y=table_y),
                size=TSPMessages.Size(height=height, width=width),
            ),
        )

    def add_table(  # noqa: PLR0917
        self,
        sheet_id: int,
        table_name: str,
        from_table_id: int,
        x: float | None,
        y: float | None,
        num_rows: int,
        num_cols: int,
        number_of_header_rows: int = 1,
        number_of_header_columns: int = 1,
    ) -> int:
        """Add table."""
        from_table = self.get_object(from_table_id, TSTArchives.TableModelArchive)

        table_strings_id, _ = self.create_string_table()

        # Build a minimal table duplicating references from the source table
        from_table_refs = field_references(from_table)
        table_model_id, table_model = self.objects.create_object_from_dict(
            "CalculationEngine",
            {
                "table_id": str(NumbersUUID()).upper(),
                "number_of_rows": num_rows,
                "number_of_columns": num_cols,
                "table_name": table_name,
                "table_name_enabled": True,
                "default_row_height": DEFAULT_ROW_HEIGHT,
                "default_column_width": DEFAULT_COLUMN_WIDTH,
                "number_of_header_rows": number_of_header_rows,
                "number_of_header_columns": number_of_header_columns,
                "header_rows_frozen": True,
                "header_columns_frozen": True,
                **from_table_refs,
            },
            TSTArchives.TableModelArchive,
        )
        # Suppress Numbers assertions for tables sharing the same data
        table_model.category_owner.identifier = 0

        column_headers_id, _ = self.objects.create_object_from_dict(
            "Index/Tables/HeaderStorageBucket-{}",
            {"bucketHashFunction": 1},
            TSTArchives.HeaderStorageBucket,
        )
        self.add_component_metadata(
            column_headers_id,
            "CalculationEngine",
            "Tables/HeaderStorageBucket-{}",
        )

        sidecar_id, _ = self.objects.create_object_from_dict(
            "CalculationEngine",
            {"max_order": 1, "column_count": 0, "row_count": 0},
            TSTArchives.StrokeSidecarArchive,
        )
        self.set_reference(table_model.stroke_sidecar, sidecar_id)

        style_table_id, _ = self.objects.create_object_from_dict(
            "Index/Tables/DataList-{}",
            {"listType": TSTArchives.TableDataList.ListType.STYLE, "nextListID": 1},
            TSTArchives.TableDataList,
        )
        self.add_component_metadata(style_table_id, "CalculationEngine", "Tables/DataList-{}")

        formula_table_id, _ = self.objects.create_object_from_dict(
            "Index/Tables/TableDataList-{}",
            {"listType": TSTArchives.TableDataList.ListType.FORMULA, "nextListID": 1},
            TSTArchives.TableDataList,
        )
        self.add_component_metadata(
            formula_table_id,
            "CalculationEngine",
            "Tables/TableDataList-{}",
        )

        format_table_pre_bnc_id, _ = self.objects.create_object_from_dict(
            "Index/Tables/TableDataList-{}",
            {"listType": TSTArchives.TableDataList.ListType.STYLE, "nextListID": 1},
            TSTArchives.TableDataList,
        )
        self.add_component_metadata(
            format_table_pre_bnc_id,
            "CalculationEngine",
            "Tables/TableDataList-{}",
        )

        data_store_refs = field_references(from_table.base_data_store)
        data_store_refs["stringTable"] = {"identifier": table_strings_id}
        data_store_refs["columnHeaders"] = {"identifier": column_headers_id}
        data_store_refs["styleTable"] = {"identifier": style_table_id}
        data_store_refs["formula_table"] = {"identifier": formula_table_id}
        data_store_refs["format_table_pre_bnc"] = {"identifier": format_table_pre_bnc_id}
        table_model.base_data_store.MergeFrom(
            TSTArchives.DataStore(
                rowHeaders=TSTArchives.HeaderStorage(bucketHashFunction=1),
                nextRowStripID=1,
                nextColumnStripID=0,
                rowTileTree=TSTArchives.TableRBTree(),
                columnTileTree=TSTArchives.TableRBTree(),
                tiles=TSTArchives.TileStorage(
                    tile_size=DEFAULT_TILE_SIZE,
                    should_use_wide_rows=True,
                ),
                **data_store_refs,
            ),
        )

        row_headers_id, _ = self.objects.create_object_from_dict(
            "Index/Tables/HeaderStorageBucket-{}",
            {"bucketHashFunction": 1},
            TSTArchives.HeaderStorageBucket,
        )

        self.add_component_metadata(
            row_headers_id,
            "CalculationEngine",
            "Tables/HeaderStorageBucket-{}",
        )
        table_model.base_data_store.rowHeaders.buckets.append(
            TSPMessages.Reference(identifier=row_headers_id),
        )

        table_info_id, table_info = self.objects.create_object_from_dict(
            "CalculationEngine",
            {},
            TSTArchives.TableInfoArchive,
        )
        table_info.tableModel.MergeFrom(TSPMessages.Reference(identifier=table_model_id))
        table_info.super.MergeFrom(self.create_drawable(sheet_id, x, y))

        haunted_owner_uuid = self.add_formula_owner(
            table_info_id,
            num_rows,
            num_cols,
            number_of_header_rows,
            number_of_header_columns,
        )
        table_model.haunted_owner.owner_uid.MergeFrom(haunted_owner_uuid.protobuf2)
        self.calculate_table_uuid_map()

        self._table_data[table_model_id] = [
            [Cell._empty_cell(table_model_id, row, col, self) for col in range(num_cols)]
            for row in range(num_rows)
        ]
        self.recalculate_table_data(table_model_id, self._table_data[table_model_id])

        self.add_component_reference(
            table_info_id,
            location="Document",
            component_id=self.calc_engine_id(),
        )
        self.create_caption_archive(table_model_id)
        self.caption_enabled(table_model_id, False)  # noqa: FBT003

        self.get_object(sheet_id, TNArchives.SheetArchive).drawable_infos.append(
            TSPMessages.Reference(identifier=table_info_id),
        )

        self.name_ref_cache.mark_dirty()
        return table_model_id

    def add_formula_owner(
        self,
        table_info_id: int,
        num_rows: int,
        num_cols: int,
        number_of_header_rows: int,
        number_of_header_columns: int,
    ) -> NumbersUUID:
        """
        Create formula-owner metadata for a table.

        This archive enables cross-references to cells in the table.
        """
        _ = num_rows
        formula_owner_uuid = NumbersUUID()
        calc_engine = self.calc_engine()
        owner_id_map = calc_engine.dependency_tracker.owner_id_map.map_entry
        next_owner_id = max(x.internal_owner_id for x in owner_id_map) + 1
        volatile_dependencies: dict[str, dict[str, Any]] = {
            "volatile_time_cells": {},
            "volatile_random_cells": {},
            "volatile_locale_cells": {},
            "volatile_sheet_table_name_cells": {},
            "volatile_remote_data_cells": {},
            "volatile_geometry_cell_refs": {},
        }
        total_range_for_table = {
            "top_left_column": 0,
            "top_left_row": 0,
            "bottom_right_column": num_cols - 1,
            "bottom_right_row": num_cols - 1,
        }
        body_range_for_table = {
            "top_left_column": number_of_header_columns,
            "top_left_row": number_of_header_rows,
            "bottom_right_column": num_cols - 1,
            "bottom_right_row": num_cols - 1,
        }

        formula_deps_id, _ = self.objects.create_object_from_dict(
            "CalculationEngine",
            {
                "formula_owner_uid": formula_owner_uuid.dict2,
                "internal_formula_owner_id": next_owner_id,
                "owner_kind": OwnerKind.TABLE_MODEL,
                "cell_dependencies": {},
                "range_dependencies": {},
                "volatile_dependencies": volatile_dependencies,
                "spanning_column_dependencies": {
                    "total_range_for_table": total_range_for_table,
                    "body_range_for_table": body_range_for_table,
                },
                "spanning_row_dependencies": {
                    "total_range_for_table": total_range_for_table,
                    "body_range_for_table": body_range_for_table,
                },
                "whole_owner_dependencies": {"dependent_cells": {}},
                "cell_errors": {},
                "formula_owner": {"identifier": table_info_id},
                "tiled_cell_dependencies": {},
                "uuid_references": {},
                "tiled_range_dependencies": {},
            },
            TSCEArchives.FormulaOwnerDependenciesArchive,
        )
        calc_engine.dependency_tracker.formula_owner_dependencies.append(
            TSPMessages.Reference(identifier=formula_deps_id),
        )
        owner_id_map.append(
            TSCEArchives.OwnerIDMapArchive.OwnerIDMapArchiveEntry(
                internal_owner_id=next_owner_id,
                owner_id=formula_owner_uuid.protobuf4,
            ),
        )

        formula_owner_uuid = NumbersUUID()
        base_owner_uuid = NumbersUUID()
        next_owner_id += 1
        null_range_ref = {
            "top_left_column": 0x7FFF,
            "top_left_row": 0x7FFFFFFF,
            "bottom_right_column": 0x7FFF,
            "bottom_right_row": 0x7FFFFFFF,
        }
        spanning_dependencies = {
            "total_range_for_table": null_range_ref,
            "body_range_for_table": null_range_ref,
        }
        formula_deps_id, _ = self.objects.create_object_from_dict(
            "CalculationEngine",
            {
                "formula_owner_uid": formula_owner_uuid.dict2,
                "internal_formula_owner_id": next_owner_id,
                "owner_kind": OwnerKind.HAUNTED_OWNER,
                "cell_dependencies": {},
                "range_dependencies": {},
                "volatile_dependencies": volatile_dependencies,
                "spanning_column_dependencies": spanning_dependencies,
                "spanning_row_dependencies": spanning_dependencies,
                "whole_owner_dependencies": {"dependent_cells": {}},
                "cell_errors": {},
                "base_owner_uid": base_owner_uuid.dict2,
                "tiled_cell_dependencies": {},
                "uuid_references": {},
                "tiled_range_dependencies": {},
            },
            TSCEArchives.FormulaOwnerDependenciesArchive,
        )
        calc_engine.dependency_tracker.formula_owner_dependencies.append(
            TSPMessages.Reference(identifier=formula_deps_id),
        )
        owner_id_map.append(
            TSCEArchives.OwnerIDMapArchive.OwnerIDMapArchiveEntry(
                internal_owner_id=next_owner_id,
                owner_id=formula_owner_uuid.protobuf4,
            ),
        )
        return formula_owner_uuid

    def add_sheet(self, sheet_name: str) -> int:
        """Add a new sheet with a copy of a table from another sheet."""
        sheet_id, _ = self.objects.create_object_from_dict(
            "Document",
            {"name": sheet_name},
            TNArchives.SheetArchive,
        )

        self.add_component_reference(
            sheet_id,
            location="CalculationEngine",
            component_id=DOCUMENT_ID,
            is_weak=True,
        )

        self.document_archive.sheets.append(TSPMessages.Reference(identifier=sheet_id))

        return sheet_id

    @property
    def styles(self) -> dict[str, Style]:
        """Return the styles available in the document."""
        if not self._styles:
            self._styles = self.available_paragraph_styles()
        return self._styles

    @cache(num_args=0)
    def available_paragraph_styles(self) -> dict[str, Style]:
        """Return the paragraph styles available in the document."""
        presets = find_extension(
            self.get_object(self.document_archive.theme.identifier, TNArchives.ThemeArchive).super,
            "paragraph_style_presets",
        )

        styles = {}
        for x in presets:
            style_obj = self.get_object(x.identifier, TSWPArchives.ParagraphStyleArchive)
            name = style_obj.super.name
            styles[name] = Style(
                alignment=Alignment(
                    HorizontalJustification(style_obj.para_properties.alignment),
                    VerticalJustification(0),
                ),
                font_color=self.cell_font_color(style_obj),
                font_size=self.cell_font_size(style_obj),
                font_name=self.cell_font_family(style_obj),
                bold=self.cell_is_bold(style_obj),
                italic=self.cell_is_italic(style_obj),
                underline=self.cell_is_underline(style_obj),
                strikethrough=self.cell_is_strikethrough(style_obj),
                name=self.cell_style_name(style_obj),
                _text_style_obj_id=x.identifier,
                _font_details=self.cell_font_details(style_obj),
            )

        for style in styles.values():
            # Override __setattr__ behavior for builtin styles
            style.__dict__["_update_text_style"] = False
            style.__dict__["_update_cell_style"] = False

        return styles

    def add_paragraph_style(self, style: Style) -> int:
        """Add paragraph style."""
        if style.underline:
            underline = CharacterStyle.UnderlineType.kSingleUnderline
        else:
            underline = CharacterStyle.UnderlineType.kNoUnderline
        if style.strikethrough:
            strikethru = CharacterStyle.StrikethruType.kSingleStrikethru
        else:
            strikethru = CharacterStyle.StrikethruType.kNoStrikethru

        name_str = style.name if style.name is not None else "Custom Style"
        style_id_name = "numbers-parser-" + name_str.lower().replace(" ", "-")
        para_style_id, para_style = self.objects.create_object_from_dict(
            "DocumentStylesheet",
            {
                "super": {
                    "name": style.name,
                    "style_identifier": style_id_name,
                },
                "override_count": 1,
                "char_properties": {
                    "font_color": {
                        "model": "rgb",
                        "r": style.font_color.r / 255,
                        "g": style.font_color.g / 255,
                        "b": style.font_color.b / 255,
                        "a": 1.0,
                        "rgbspace": "srgb",
                    },
                    "bold": style.bold or (style._font_details or {})["bold"],
                    "italic": style.italic or (style._font_details or {})["italic"],
                    "underline": underline,
                    "strikethru": strikethru,
                    "font_size": style.font_size,
                    "font_name": (style._font_details or {})["name"],
                    "tsd_fill": {
                        "color": {
                            "model": "rgb",
                            "r": style.font_color.r / 255,
                            "g": style.font_color.g / 255,
                            "b": style.font_color.b / 255,
                            "a": 1.0,
                            "rgbspace": "srgb",
                        },
                    },
                },
                "para_properties": {
                    "alignment": style.alignment.horizontal,
                    "first_line_indent": style.first_indent,
                    "left_indent": style.left_indent,
                    "right_indent": style.right_indent,
                },
            },
            TSWPArchives.ParagraphStyleArchive,
        )
        stylesheet_id = self.document_archive.stylesheet.identifier
        para_style.super.stylesheet.MergeFrom(TSPMessages.Reference(identifier=stylesheet_id))
        stylesheet_archive = self.get_object(stylesheet_id, TSSArchives.StylesheetArchive)
        stylesheet_archive.styles.append(TSPMessages.Reference(identifier=para_style_id))
        stylesheet_archive.identifier_to_style_map.append(
            TSSArchives.StylesheetArchive.IdentifiedStyleEntry(
                identifier=style_id_name,
                style=TSPMessages.Reference(identifier=para_style_id),
            ),
        )

        presets = find_extension(
            self.get_object(self.document_archive.theme.identifier, TNArchives.ThemeArchive).super,
            "paragraph_style_presets",
        )
        presets.append(TSPMessages.Reference(identifier=para_style_id))
        self.styles[name_str] = style
        return para_style_id

    def update_paragraph_style(self, style: Style) -> None:
        """Update paragraph style."""
        if style.underline:
            underline = CharacterStyle.UnderlineType.kSingleUnderline
        else:
            underline = CharacterStyle.UnderlineType.kNoUnderline
        if style.strikethrough:
            strikethru = CharacterStyle.StrikethruType.kSingleStrikethru
        else:
            strikethru = CharacterStyle.StrikethruType.kNoStrikethru
        style_obj = self.get_object(
            style._text_style_obj_id or 0,
            TSWPArchives.ParagraphStyleArchive,
        )
        style_obj.char_properties.font_color.r = style.font_color.r / 255
        style_obj.char_properties.font_color.g = style.font_color.g / 255
        style_obj.char_properties.font_color.b = style.font_color.b / 255
        style_obj.char_properties.bold = style.bold
        style_obj.char_properties.italic = style.italic
        style_obj.char_properties.underline = underline
        style_obj.char_properties.strikethru = strikethru
        style_obj.char_properties.font_size = style.font_size
        style_obj.char_properties.font_name = (style._font_details or {})["name"]
        style_obj.char_properties.tsd_fill.color.r = style.font_color.r / 255
        style_obj.char_properties.tsd_fill.color.g = style.font_color.g / 255
        style_obj.char_properties.tsd_fill.color.b = style.font_color.b / 255
        style_obj.para_properties.alignment = cast(
            "TSWPArchives.ParagraphStylePropertiesArchive.TextAlignmentType.ValueType",
            int(style.alignment.horizontal),
        )
        style_obj.para_properties.first_line_indent = style.first_indent
        style_obj.para_properties.left_indent = style.left_indent
        style_obj.para_properties.right_indent = style.right_indent

    def update_paragraph_styles(self) -> None:
        """
        Create paragraph style archives for newly added styles.

        Update existing archives for styles that have changed.
        """
        new_styles = [x for x in self.styles.values() if x._text_style_obj_id is None]
        updated_styles = [
            x
            for x in self.styles.values()
            if x._text_style_obj_id is not None and x._update_text_style
        ]
        for style in new_styles:
            style._text_style_obj_id = self.add_paragraph_style(style)
            style._update_text_style = True

        for style in updated_styles:
            self.update_paragraph_style(style)

    def update_cell_styles(self, table_id: int, data: list[list[Cell]]) -> None:
        """Create or update archives for cells with changed styles."""
        _ = table_id
        cell_styles = {}
        for _, cells in enumerate(data):
            for _, cell in enumerate(cells):
                if cell._style is not None and cell._style._update_cell_style:
                    fingerprint = (
                        str(cell._style.alignment.vertical)
                        + str(cell._style.first_indent)
                        + str(cell._style.left_indent)
                        + str(cell._style.right_indent)
                        + str(cell._style.text_inset)
                        + str(cell._style.text_wrap)
                    )

                    if cell._style.bg_color is not None:
                        if isinstance(cell._style.bg_color, list):
                            for color in cell._style.bg_color:
                                fingerprint += str(color.r) + str(color.g) + str(color.b)
                        else:
                            fingerprint += (
                                str(cell._style.bg_color.r)
                                + str(cell._style.bg_color.g)
                                + str(cell._style.bg_color.b)
                            )

                    if cell._style.bg_image is not None:
                        fingerprint += cell._style.bg_image.filename

                    if fingerprint not in cell_styles:
                        cell_styles[fingerprint] = self.add_cell_style(cell._style)
                    cell._style._cell_style_obj_id = cell_styles[fingerprint]

    def add_cell_style(self, style: Style) -> int:
        """Add cell style."""
        if style.bg_image is not None:
            digest = sha1(style.bg_image.data).digest()  # noqa: S324
            if digest in self.images:
                image_id = self.images[digest]
            else:
                datas = self.package_metadata.datas
                image_id = self.next_image_identifier()
                datas.append(
                    TSPArchiveMessages.DataInfo(
                        identifier=image_id,
                        digest=digest,
                        preferred_file_name=style.bg_image.filename,
                        file_name=style.bg_image.filename,
                        materialized_length=len(style.bg_image.data),
                    ),
                )
                self.images[digest] = image_id
            color_attrs = {
                "cell_fill": {
                    "image": {
                        "technique": "ScaleToFill",
                        "imagedata": {"identifier": image_id},
                        "interpretsUntaggedImageDataAsGeneric": False,
                    },
                },
            }
        elif style.bg_color is not None:
            if isinstance(style.bg_color, list):
                stops = []
                for i, color in enumerate(style.bg_color):
                    fraction = i / (len(style.bg_color) - 1) if len(style.bg_color) > 1 else 1.0
                    stops.append(
                        {
                            "color": {
                                "model": "rgb",
                                "r": color.r / 255,
                                "g": color.g / 255,
                                "b": color.b / 255,
                                "a": 1.0,
                                "rgbspace": "srgb",
                            },
                            "fraction": float(fraction),
                            "inflection": 0.5,
                        },
                    )
                color_attrs = {
                    "cell_fill": {
                        "gradient": {
                            "type": "Linear",
                            "opacity": 1.0,
                            "advancedGradient": True,
                            "anglegradient": {"gradientangle": 0.0},
                            "stops": stops,
                        },
                    },
                }
            else:
                color_attrs = {
                    "cell_fill": {
                        "color": {
                            "model": "rgb",
                            "r": style.bg_color.r / 255,
                            "g": style.bg_color.g / 255,
                            "b": style.bg_color.b / 255,
                            "a": 1.0,
                            "rgbspace": "srgb",
                        },
                    },
                }
        else:
            color_attrs = {}

        cell_style_id, cell_style = self.objects.create_object_from_dict(
            "DocumentStylesheet",
            {
                "super": {
                    "name": style.name,
                    "style_identifier": "",
                },
                "override_count": 1,
                "cell_properties": {
                    **color_attrs,
                    "padding": {
                        "left": style.text_inset,
                        "top": style.text_inset,
                        "right": style.text_inset,
                        "bottom": style.text_inset,
                    },
                    "text_wrap": style.text_wrap,
                    "vertical_alignment": style.alignment.vertical,
                },
            },
            TSTArchives.CellStyleArchive,
        )
        style_id_name = f"numbers-parser-custom-{cell_style_id}"
        cell_style.super.style_identifier = style_id_name

        stylesheet_id = self.document_archive.stylesheet.identifier
        cell_style.super.stylesheet.MergeFrom(TSPMessages.Reference(identifier=stylesheet_id))
        stylesheet_archive = self.get_object(stylesheet_id, TSSArchives.StylesheetArchive)
        stylesheet_archive.styles.append(TSPMessages.Reference(identifier=cell_style_id))
        stylesheet_archive.identifier_to_style_map.append(
            TSSArchives.StylesheetArchive.IdentifiedStyleEntry(
                identifier=style_id_name,
                style=TSPMessages.Reference(identifier=cell_style_id),
            ),
        )

        return cell_style_id

    def text_style_object_id(self, cell: Cell) -> int | None:
        """Return the archive identifier for a cell text style."""
        if cell._text_style_id is None:
            return None
        entry = self.table_styles.lookup_value(cell._table_id, cell._text_style_id)
        return entry.reference.identifier

    def cell_style_object_id(self, cell: Cell) -> int | None:
        """Return the archive identifier for a cell style."""
        if cell._cell_style_id is None:
            return None
        entry = self.table_styles.lookup_value(cell._table_id, cell._cell_style_id)
        return entry.reference.identifier

    def custom_style_name(self) -> str:
        """
        Return the next available custom style name.

        Names use the highest available number in the current document.
        """
        stylesheet_id = self.document_archive.stylesheet.identifier
        current_styles = self.styles.keys()
        custom_styles = [x for x in current_styles if re.fullmatch(r"Custom Style \d+", x)]
        for style_entry in self.get_object(
            stylesheet_id,
            TSSArchives.StylesheetArchive,
        ).identifier_to_style_map:
            style_id = style_entry.style.identifier
            style_name = getattr(
                self.get_object(style_id, TSWPArchives.ShapeStyleArchive).super,
                "name",
                "",
            )
            if re.fullmatch(r"Custom Style \d+", style_name):
                custom_styles.append(style_name)

        if len(custom_styles) > 0:
            offset = len("Custom Style ")
            custom_style_ids = [int(x[offset:]) for x in custom_styles]
            return "Custom Style " + str(custom_style_ids[-1] + 1)
        return "Custom Style 1"

    @property
    def custom_formats(self) -> dict[str, CustomFormatting]:
        """Return the custom formats defined in the document."""
        if not self._custom_formats:
            custom_format_list = self.get_object(
                self.document_archive.super.custom_format_list.identifier,
                TSKArchives.CustomFormatListArchive,
            )
            custom_formats = custom_format_list.custom_formats
            custom_format_names = [x.name for x in custom_formats]
            custom_format_uuids = list(custom_format_list.uuids)
            for i, format_name in enumerate(custom_format_names):
                self._custom_formats[format_name] = CustomFormatting.from_archive(custom_formats[i])
                self._custom_format_archives[format_name] = custom_formats[i]
                self._custom_format_uuids[format_name] = custom_format_uuids[i]

        return self._custom_formats

    def custom_format_name(self) -> str:
        """
        Return the next available custom format name.

        Names use the highest available number in the current document.
        """
        current_formats = self.custom_formats.keys()
        if "Custom Format" not in current_formats:
            return "Custom Format"
        format_ids = [
            m.group(1) for x in current_formats if (m := re.fullmatch(r"Custom Format (\d+)", x))
        ]
        if len(format_ids) > 0:
            last_id = int(format_ids[-1])
            return f"Custom Format {last_id + 1}"
        return "Custom Format 1"

    @cache()
    def table_formulas(self, table_id: int) -> TableFormulas:
        """Return the formula collection for a table."""
        return TableFormulas(self, table_id)

    @cache(num_args=2)
    def table_rich_text(  # noqa: C901
        self,
        table_id: int,
        string_key: int,
    ) -> dict | None:
        """Extract bullets and hyperlinks from a rich text data cell."""
        rich_text_table = self.get_object(
            self.table_model(table_id).base_data_store.rich_text_table.identifier,
            TSTArchives.TableDataList,
        )
        for entry in rich_text_table.entries:  # pragma: no branch
            if string_key == entry.key:
                payload = self.get_object(
                    entry.rich_text_payload.identifier,
                    TSTArchives.RichTextPayloadArchive,
                )
                payload_storage = self.get_object(
                    payload.storage.identifier,
                    TSWPArchives.StorageArchive,
                )
                smartfield_entries = payload_storage.table_smartfield.entries
                cell_text = payload_storage.text[0] if len(payload_storage.text) > 0 else ""

                hyperlinks = []
                for i, e in enumerate(smartfield_entries):
                    if e.object.identifier:
                        obj = self.objects[e.object.identifier]
                        if isinstance(obj, TSWPArchives.HyperlinkFieldArchive):
                            start = e.character_index
                            if i < len(smartfield_entries) - 1:
                                end = smartfield_entries[i + 1].character_index
                            else:
                                end = len(cell_text)
                            url_text = cell_text[start:end]
                            hyperlinks.append((url_text, obj.url_ref))

                bullets = []
                bullet_chars = []
                payload_entries = payload_storage.table_para_style.entries
                table_list_styles = payload_storage.table_list_style.entries
                offsets = [e.character_index for e in payload_entries]
                table_list_style = None
                for i, offset in enumerate(offsets):
                    if i == len(offsets) - 1:
                        bullets.append(cell_text[offset:])
                    else:
                        # Remove the last character (always newline)
                        bullets.append(cell_text[offset : offsets[i + 1] - 1])

                    # Re-use last style if there is none defined for this bullet
                    if i < len(table_list_styles):
                        table_list_style = table_list_styles[i]

                    if table_list_style is None:
                        bullet_char = None
                    else:
                        bullet_style = self.get_object(
                            table_list_style.object.identifier,
                            TSWPArchives.ListStyleArchive,
                        )
                        if len(bullet_style.strings) > 0:
                            bullet_char = bullet_style.strings[0]
                        elif len(bullet_style.number_types) > 0:
                            number_type = bullet_style.number_types[0]
                            bullet_char = formatted_number(number_type, i)
                        else:
                            bullet_char = None

                    bullet_chars.append(bullet_char)

                return {
                    "text": cell_text,
                    "bulleted": any(c is not None for c in bullet_chars),
                    "bullets": bullets,
                    "bullet_chars": bullet_chars,
                    "hyperlinks": hyperlinks,
                }
        return None

    def table_comments(self, table_id: int, key: int) -> list[Comment]:
        """Resolve a comment ID to a list of Comment objects."""
        entry = self._table_comments.lookup_value(table_id, key)
        storage_id = entry.comment_storage.identifier
        return [self._resolve_comment_storage(storage_id)]

    def _resolve_comment_storage(self, storage_id: int) -> Comment:
        """Recursively fetch a comment and its replies into a Comment object."""
        storage = self.get_object(storage_id, TSDArchives.CommentStorageArchive)

        author_id = storage.author.identifier
        author_name = (
            self.get_object(author_id, TSKArchives.AnnotationAuthorArchive).name
            if author_id in self.objects
            else "Unknown"
        )
        creation_date = EPOCH + timedelta(seconds=storage.creation_date.seconds)

        replies = [
            self._resolve_comment_storage(ref.identifier) for ref in getattr(storage, "replies", [])
        ]

        return Comment(
            author=author_name,
            text=storage.text,
            creation_date=creation_date,
            replies=replies,
        )

    def author_id(self, author_name: str, table_id: int) -> int:
        """Get or create an annotation author."""
        storage_id = self.find_refs("AnnotationAuthorStorageArchive")[0]
        storage = self.get_object(storage_id, TSKArchives.AnnotationAuthorStorageArchive)

        for author_ref in storage.annotation_author:
            author = self.get_object(author_ref.identifier, TSKArchives.AnnotationAuthorArchive)
            if author.name == author_name:
                return author_ref.identifier

        # Append new authors to the same file that holds the Author Storage array
        new_author_id, _ = self.objects.create_object_from_dict(
            "AnnotationAuthorStorage",
            {
                "name": author_name,
                "color": {
                    "model": "rgb",
                    "r": 0.0,
                    "g": 0.0,
                    "b": 0.0,
                    "a": 1.0,
                    "rgbspace": "srgb",
                },
            },
            TSKArchives.AnnotationAuthorArchive,
            append=True,
        )
        storage.annotation_author.append(TSPMessages.Reference(identifier=new_author_id))
        bds = self.table_model(table_id).base_data_store
        self.add_component_reference(
            object_id=new_author_id,
            location=bds.commentStorageTable.identifier,
            component_id=storage_id,
        )
        return new_author_id

    def add_comment_to_cell(  # noqa: PLR0917
        self,
        table_id: int,
        row: int,
        col: int,
        text: str,
        author: str,
        reply: bool,
    ) -> None:
        """Add comment to cell."""
        cell = self._table_data[table_id][row][col]

        if reply and cell._comment_id is None:
            msg = "Cannot reply to a comment in a cell with no comments"
            raise ValueError(msg)
        if not reply and cell._comment_id is not None:
            msg = "Cannot reply to existing comment without reply=True"
            raise ValueError(msg)

        author_id = self.author_id(author, table_id)
        creation_date = (datetime.now() - EPOCH).total_seconds()  # noqa: DTZ005
        datalist_id = self._table_comments.id(table_id)
        datalist_file = self.objects._object_to_filename_map[datalist_id].strip(".iwa")
        new_comment_id, _ = self.objects.create_object_from_dict(
            datalist_file,
            {
                "text": text,
                "creation_date": {"seconds": creation_date},
                "author": {"identifier": author_id},
                "storage_uuid": NumbersUUID().dict2,
            },
            TSDArchives.CommentStorageArchive,
            append=True,
        )
        new_comment_ref = TSPMessages.Reference(identifier=new_comment_id)
        if reply:
            entry = self._table_comments.lookup_value(table_id, cell._comment_id or 0)
            root_comment_id = entry.comment_storage.identifier
            root_comment = self.get_object(root_comment_id, TSDArchives.CommentStorageArchive)
            root_comment.replies.append(new_comment_ref)
        else:
            key = self._table_comments.lookup_key(table_id, new_comment_ref)
            cell._comment_id = key

    def cell_text_style(self, cell: Cell) -> TSWPArchives.ParagraphStyleArchive:
        """Return the cell's text style or default header, footer, or body style."""
        if cell._text_style_id is not None:
            style_entry = self.table_styles.lookup_value(cell._table_id, cell._text_style_id)
            return self.get_object(
                style_entry.reference.identifier,
                TSWPArchives.ParagraphStyleArchive,
            )

        table_model = self.table_model(cell._table_id)
        if cell.row in range(table_model.number_of_header_rows):
            return self.get_object(
                table_model.header_row_text_style.identifier,
                TSWPArchives.ParagraphStyleArchive,
            )
        if cell.col in range(table_model.number_of_header_columns):
            return self.get_object(
                table_model.header_column_text_style.identifier,
                TSWPArchives.ParagraphStyleArchive,
            )
        if table_model.number_of_footer_rows > 0:
            start_row_num = table_model.number_of_rows - table_model.number_of_footer_rows
            end_row_num = start_row_num + table_model.number_of_footer_rows
            if cell.row in range(start_row_num, end_row_num):
                return self.get_object(
                    table_model.footer_row_text_style.identifier,
                    TSWPArchives.ParagraphStyleArchive,
                )
        return self.get_object(
            table_model.body_text_style.identifier,
            TSWPArchives.ParagraphStyleArchive,
        )

    def cell_alignment(self, cell: Cell) -> Alignment:
        """Return the cell alignment."""
        text_style = self.cell_text_style(cell)
        horizontal = HorizontalJustification(int(self.para_property(text_style, "alignment")))

        if cell._cell_style_id is None:
            vertical = VerticalJustification.TOP
        else:
            cell_style = self.table_style(cell._table_id, cell._cell_style_id)
            vertical = VerticalJustification(
                cast("int", self.cell_property(cell_style, "vertical_alignment")),
            )
        return Alignment(horizontal, vertical)

    def cell_bg_color(self, cell: Cell) -> RGB | list[RGB] | None:
        """Return the cell bg color."""
        if cell._cell_style_id is None:
            return None

        style = self.table_style(cell._table_id, cell._cell_style_id)
        cell_properties = style.cell_properties.cell_fill

        if cell_properties.HasField("color"):
            return rgb(cell_properties.color)
        if cell_properties.HasField("gradient"):
            return [(rgb(s.color)) for s in cell_properties.gradient.stops]
        return None

    def char_property(
        self,
        style: TSWPArchives.ParagraphStyleArchive,
        field: str,
    ) -> Any:  # noqa: ANN401
        """Return a character property from the style or its parent."""
        if not style.char_properties.HasField(field):  # type: ignore[arg-type]
            parent = self.get_object(
                style.super.parent.identifier,
                TSWPArchives.ParagraphStyleArchive,
            )
            return getattr(parent.char_properties, field)
        return getattr(style.char_properties, field)

    def para_property(
        self,
        style: TSWPArchives.ParagraphStyleArchive,
        field: str,
    ) -> Any:  # noqa: ANN401
        """Return a paragraph property from the style or its parent."""
        if not style.para_properties.HasField(field):  # type: ignore[arg-type]
            parent = self.get_object(
                style.super.parent.identifier,
                TSWPArchives.ParagraphStyleArchive,
            )
            return getattr(parent.para_properties, field)
        return getattr(style.para_properties, field)

    def cell_property(
        self,
        style: TSTArchives.CellStyleArchive,
        field: str,
    ) -> Any:  # noqa: ANN401
        """Return a cell property from the style or its parent."""
        if not style.cell_properties.HasField(field):  # type: ignore[arg-type]
            parent = self.get_object(
                style.super.parent.identifier,
                TSTArchives.CellStyleArchive,
            )
            return getattr(parent.cell_properties, field)
        return getattr(style.cell_properties, field)

    def cell_is_bold(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> bool:
        """Return whether the cell is bold."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return bool(self.char_property(style, "bold"))

    def cell_is_italic(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> bool:
        """Return whether the cell is italic."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return bool(self.char_property(style, "italic"))

    def cell_is_underline(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> bool:
        """Return whether the cell is underline."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        underline = self.char_property(style, "underline")
        return underline != CharacterStyle.UnderlineType.kNoUnderline

    def cell_is_strikethrough(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> bool:
        """Return whether the cell is strikethrough."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        strikethru = self.char_property(style, "strikethru")
        return strikethru != CharacterStyle.StrikethruType.kNoStrikethru

    def cell_style_name(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> str:
        """Return the cell style name."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return style.super.name

    def cell_font_color(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> RGB:
        """Return the cell font color."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return rgb(self.char_property(style, "font_color"))

    def cell_font_size(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> float:
        """Return the cell font size."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.char_property(style, "font_size")

    def cell_font_family(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> str:
        """Return the cell font family."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        font_name = self.char_property(style, "font_name")
        if font_name not in FONT_MAP:
            if font_name not in self.missing_fonts:
                warn(
                    f"Custom font '{font_name}' unsupported; falling back to {DEFAULT_FONT}",
                    UnsupportedWarning,
                    stacklevel=2,
                )
                self.missing_fonts[font_name] = True
            return DEFAULT_FONT

        return cast("str", FONT_MAP[font_name]["family"])

    def cell_font_details(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> dict[str, Any]:
        """Return the cell font details."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        font_name = self.char_property(style, "font_name")
        if font_name not in FONT_MAP:
            return FONT_FAMILY_DEFAULT[DEFAULT_FONT]
        return FONT_MAP[font_name]

    def cell_first_indent(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> float:
        """Return the cell first indent."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.para_property(style, "first_line_indent")

    def cell_left_indent(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> float:
        """Return the cell left indent."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.para_property(style, "left_indent")

    def cell_right_indent(self, obj: Cell | TSWPArchives.ParagraphStyleArchive) -> float:
        """Return the cell right indent."""
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.para_property(style, "right_indent")

    def cell_text_inset(self, cell: Cell) -> float:
        """Return the cell text inset."""
        if cell._cell_style_id is None:
            return DEFAULT_TEXT_INSET
        style = self.table_style(cell._table_id, cell._cell_style_id)
        padding = cast("TSWPArchives.PaddingArchive", self.cell_property(style, "padding"))
        # Padding is always identical (only one UI setting)
        return padding.left

    def cell_text_wrap(self, cell: Cell) -> bool:
        """Return the cell text wrap."""
        if cell._cell_style_id is None:
            return DEFAULT_TEXT_WRAP
        style = self.table_style(cell._table_id, cell._cell_style_id)
        return bool(self.cell_property(style, "text_wrap"))

    def stroke_type(self, stroke_run: TSTArchives.StrokeLayerArchive.StrokeRunArchive) -> str:
        """Return the stroke type for a stroke run."""
        stroke_type = stroke_run.stroke.pattern.type
        if stroke_type == StrokePattern.StrokePatternType.TSDSolidPattern:
            return "solid"
        if stroke_type == StrokePattern.StrokePatternType.TSDPattern:
            if stroke_run.stroke.pattern.pattern[0] < 1.0:
                return "dots"
            return "dashes"
        return "none"

    def cell_for_stroke(self, table_id: int, side: str, row: int, col: int) -> Cell | None:
        """Return the cell for stroke."""
        data = self._table_data[table_id]
        if row < 0 or col < 0:
            return None
        if row >= len(data) or col >= len(data[row]):
            return None
        cell = self._table_data[table_id][row][col]
        if isinstance(cell, MergedCell):
            if (
                (side == "top" and row == cell.row_start)
                or (side == "right" and col == cell.col_end)
                or (side == "bottom" and row == cell.row_end)
                or (side == "left" and col == cell.col_start)
            ):
                return cell
        elif cell.is_merged:
            # No borders for merged edges of anchor cells
            size = cell.size or (1, 1)
            if (
                side in ["top", "left"]
                or (side == "right" and size[1] == 1)
                or (side == "bottom" and size[0] == 1)
            ):
                return cell
        else:
            return cell
        return None

    def set_cell_border(  # noqa: C901
        self,
        table_id: int,
        row: int,
        col: int,
        side: str,
        border_value: Border,
    ) -> None:
        """Set the 2 borders adjacent to a stroke if within the table range."""
        self._update_strokes[table_id] = True
        if side == "top":
            bottom_cell = self.cell_for_stroke(table_id, "bottom", row - 1, col)
            if (cell := self.cell_for_stroke(table_id, "top", row, col)) is not None:
                cell._border.top = border_value
            if bottom_cell is not None:
                bottom_cell._border.bottom = border_value
        elif side == "right":
            left_cell = self.cell_for_stroke(table_id, "left", row, col + 1)
            if (cell := self.cell_for_stroke(table_id, "right", row, col)) is not None:
                cell._border.right = border_value
            if left_cell is not None:
                left_cell._border.left = border_value
        elif side == "bottom":
            top_cell = self.cell_for_stroke(table_id, "top", row + 1, col)
            if (cell := self.cell_for_stroke(table_id, "bottom", row, col)) is not None:
                cell._border.bottom = border_value
            if top_cell is not None:
                top_cell._border.top = border_value
        else:  # left border
            right_cell = self.cell_for_stroke(table_id, "right", row, col - 1)
            if (cell := self.cell_for_stroke(table_id, "left", row, col)) is not None:
                cell._border.left = border_value
            if right_cell is not None:
                right_cell._border.right = border_value

    def extract_strokes_in_layers(
        self,
        table_id: int,
        layer_ids: Sequence,
        side: str,
    ) -> list[tuple[int, int, int, int, str, Border]]:
        """Extract strokes in layers."""
        strokes = []
        for layer_id in layer_ids:
            stroke_layer = self.get_object(layer_id.identifier, TSTArchives.StrokeLayerArchive)
            for stroke_run in stroke_layer.stroke_runs:
                border_value = Border(
                    width=round(stroke_run.stroke.width, 2),
                    color=rgb(stroke_run.stroke.color),
                    style=self.stroke_type(stroke_run),
                )
                if side in ["top", "bottom"]:
                    start_row = stroke_layer.row_column_index
                    start_column = stroke_run.origin
                    for col in range(start_column, start_column + stroke_run.length):
                        strokes.append(  # noqa: PERF401
                            (
                                stroke_run.order,
                                table_id,
                                start_row,
                                col,
                                side,
                                border_value,
                            ),
                        )
                else:
                    start_row = stroke_run.origin
                    start_column = stroke_layer.row_column_index
                    for row in range(start_row, start_row + stroke_run.length):
                        strokes.append(  # noqa: PERF401
                            (
                                stroke_run.order,
                                table_id,
                                row,
                                start_column,
                                side,
                                border_value,
                            ),
                        )
        return strokes

    @cache()
    def extract_strokes(self, table_id: int) -> None:
        """Extract strokes."""
        sidecar = self.get_object(
            self.table_model(table_id).stroke_sidecar.identifier,
            TSTArchives.StrokeSidecarArchive,
        )
        strokes = []
        strokes.extend(
            self.extract_strokes_in_layers(table_id, sidecar.top_row_stroke_layers, "top"),
        )
        strokes.extend(
            self.extract_strokes_in_layers(table_id, sidecar.left_column_stroke_layers, "left"),
        )
        strokes.extend(
            self.extract_strokes_in_layers(
                table_id,
                sidecar.right_column_stroke_layers,
                "right",
            ),
        )
        strokes.extend(
            self.extract_strokes_in_layers(
                table_id,
                sidecar.bottom_row_stroke_layers,
                "bottom",
            ),
        )
        for _, stroke_table_id, row, col, side, border_value in sorted(
            strokes,
            key=lambda value: value[0],
        ):
            self.set_cell_border(stroke_table_id, row, col, side, border_value)

    def create_stroke(
        self,
        origin: int,
        length: int,
        border_value: Border,
    ) -> TSTArchives.StrokeLayerArchive.StrokeRunArchive:
        """Create stroke."""
        line_cap = TSDArchives.StrokeArchive.LineCap.ButtCap
        line_join = TSDArchives.LineJoin.MiterJoin
        if border_value.style == BorderType.SOLID:
            pattern = TSDArchives.StrokePatternArchive(
                type=StrokePattern.StrokePatternType.TSDSolidPattern,
                phase=0.0,
                count=0,
                pattern=[0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            )
        elif border_value.style == BorderType.DASHES:
            pattern = TSDArchives.StrokePatternArchive(
                type=StrokePattern.StrokePatternType.TSDPattern,
                phase=0.0,
                count=2,
                pattern=[2.0, 2.0, 0.0, 0.0, 0.0, 0.0],
            )
        elif border_value.style == BorderType.DOTS:
            pattern = TSDArchives.StrokePatternArchive(
                type=StrokePattern.StrokePatternType.TSDPattern,
                phase=0.0,
                count=2,
                pattern=[0.0001, 2.0, 0.0, 0.0, 0.0, 0.0],
            )
            line_cap = TSDArchives.StrokeArchive.LineCap.RoundCap
            line_join = TSDArchives.LineJoin.RoundJoin
        else:
            pattern = TSDArchives.StrokePatternArchive(
                type=StrokePattern.StrokePatternType.TSDEmptyPattern,
                phase=0.0,
                count=0,
                pattern=[0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            )

        color = TSPMessages.Color(
            model=TSPMessages.Color.rgb,
            rgbspace=TSPMessages.Color.srgb,
            r=border_value.color.r / 255,
            g=border_value.color.g / 255,
            b=border_value.color.b / 255,
            a=1.0,
        )
        width = border_value.width
        return TSTArchives.StrokeLayerArchive.StrokeRunArchive(
            origin=origin,
            length=length,
            order=1,  # When we re-save, strokes are optimized to be flat
            stroke=TSDArchives.StrokeArchive(
                color=color,
                width=width,
                cap=line_cap,
                join=line_join,
                miter_limit=4.0,
                pattern=pattern,
            ),
        )

    def update_cell_borders(self, table_id: int, data: list) -> None:  # noqa: C901
        """Consolidate identical strokes and then generate stoke archives."""
        if table_id not in self._update_strokes:
            return

        num_rows = self.table_model(table_id).number_of_rows
        num_cols = self.table_model(table_id).number_of_columns

        # Clear existing stroke layers to prepare for rebuilt, optimized strokes
        table_model = self.table_model(table_id)
        sidecar = self.get_object(
            table_model.stroke_sidecar.identifier,
            TSTArchives.StrokeSidecarArchive,
        )
        sidecar.top_row_stroke_layers.clear()
        sidecar.bottom_row_stroke_layers.clear()
        sidecar.left_column_stroke_layers.clear()
        sidecar.right_column_stroke_layers.clear()

        # Horizontal strokes (all but bottom of table)
        for row in range(num_rows):
            col = 0
            while col < num_cols:
                border = data[row][col]._border.top
                if border is not None:
                    length = 1
                    while col + length < num_cols and data[row][col + length]._border.top == border:
                        length += 1
                    self.add_stroke(table_id, row, col, "top", border, length)
                    col += length
                else:
                    col += 1

        # Bottom horizontal stroke
        row = num_rows - 1
        col = 0
        while col < num_cols:
            border = data[row][col]._border.bottom
            if border is not None:
                length = 1
                while col + length < num_cols and data[row][col + length]._border.bottom == border:
                    length += 1
                self.add_stroke(table_id, row, col, "bottom", border, length)
                col += length
            else:
                col += 1

        # Vertical strokes (all but right hand side of table)
        for col in range(num_cols):
            row = 0
            while row < num_rows:
                border = data[row][col]._border.left
                if border is not None:
                    length = 1
                    while (
                        row + length < num_rows and data[row + length][col]._border.left == border
                    ):
                        length += 1
                    self.add_stroke(table_id, row, col, "left", border, length)
                    row += length
                else:
                    row += 1

        # Right hand side of table
        col = num_cols - 1
        row = 0
        while row < num_rows:
            border = data[row][col]._border.right
            if border is not None:
                length = 1
                while row + length < num_rows and data[row + length][col]._border.right == border:
                    length += 1
                self.add_stroke(table_id, row, col, "right", border, length)
                row += length
            else:
                row += 1

    def add_stroke(  # noqa: PLR0917
        self,
        table_id: int,
        row: int,
        col: int,
        side: str,
        border_value: Border,
        length: int,
    ) -> None:
        """Add stroke."""
        table_model = self.table_model(table_id)
        sidecar = self.get_object(
            table_model.stroke_sidecar.identifier,
            TSTArchives.StrokeSidecarArchive,
        )
        sidecar.row_count = table_model.number_of_rows
        sidecar.column_count = table_model.number_of_columns
        sidecar.max_order = 2

        if side == "top":
            layer_ids = sidecar.top_row_stroke_layers
            row_column_index = row
            origin = col
        elif side == "right":
            layer_ids = sidecar.right_column_stroke_layers
            row_column_index = col
            origin = row
        elif side == "bottom":
            layer_ids = sidecar.bottom_row_stroke_layers
            row_column_index = row
            origin = col
        else:  # left border
            layer_ids = sidecar.left_column_stroke_layers
            row_column_index = col
            origin = row

        stroke_layer = None
        for layer_id in layer_ids:
            tmp_stroke_layer = self.get_object(layer_id.identifier, TSTArchives.StrokeLayerArchive)
            if tmp_stroke_layer.row_column_index == row_column_index:
                stroke_layer = tmp_stroke_layer
        if stroke_layer is not None:
            stroke_layer.stroke_runs.append(self.create_stroke(origin, length, border_value))
            stroke_layer.stroke_runs.sort(key=lambda x: x.origin)
        else:
            stroke_layer_id, stroke_layer = self.objects.create_object_from_dict(
                "CalculationEngine",
                {
                    "row_column_index": row_column_index,
                },
                TSTArchives.StrokeLayerArchive,
            )
            stroke_layer.stroke_runs.append(self.create_stroke(origin, length, border_value))
            layer_ids.append(TSPMessages.Reference(identifier=stroke_layer_id))

    def store_image(self, data: bytes, filename: str) -> None:
        """Store image data in the file store."""
        stored_filename = f"Data/{filename}"
        if stored_filename in self.objects.file_store:
            msg = f"{filename}: image already exists in document"
            raise IndexError(msg)
        self.objects.file_store[stored_filename] = data

    def next_image_identifier(self) -> int:
        """Return the next available ID in the list of images in the document."""
        datas = self.package_metadata.datas
        image_ids = [x.identifier for x in datas]
        # datas never appears to be an empty list (default themes include images)
        return max(image_ids) + 1

    @cache(num_args=0)
    def group_uuid_values(self) -> dict[NumbersUUID, str | float | bool | datetime | None]:
        """Group cell values by their associated UUID identifiers."""

        def cell_value_to_key(
            cell_value: TSCEArchives.CellValueArchive,
        ) -> str | float | bool | datetime | None:
            """Convert a CellValueArchive to a key."""
            cell_value_type = cell_value.cell_value_type
            if cell_value_type == CellValueType.STRING_TYPE:
                return cell_value.string_value.value
            if cell_value_type == CellValueType.NUMBER_TYPE:
                return cell_value.number_value.value
            if cell_value_type == CellValueType.BOOLEAN_TYPE:
                return cell_value.boolean_value.value
            if cell_value_type == CellValueType.DATE_TYPE:
                return _decode_date_format(
                    cell_value.date_value.format.date_time_format,
                    EPOCH + timedelta(seconds=cell_value.date_value.value),
                )
            return None

        return {
            NumbersUUID(
                self.get_object(
                    group_id,
                    TSTArchives.GroupByArchive.GroupNodeArchive,
                ).group_uid,
            ): cell_value_to_key(
                self.get_object(
                    group_id,
                    TSTArchives.GroupByArchive.GroupNodeArchive,
                ).group_cell_value,
            )
            for group_id in self.find_refs("GroupNodeArchive")
        }

    @cache()
    def calculate_table_categories(  # noqa: C901
        self,
        table_id: int,
    ) -> None:
        """Calculate table categories."""
        category_owner_id = self.table_model(table_id).category_owner.identifier
        if not category_owner_id:
            self.table_categories_data[table_id] = None
            self.table_categories_row_mapper[table_id] = None
            return

        category_archive_id = (
            self.get_object(category_owner_id, TSTArchives.CategoryOwnerRefArchive)
            .group_by[0]
            .identifier
        )
        category_archive = self.get_object(category_archive_id, TSTArchives.GroupByArchive)
        if not category_archive.is_enabled:
            self.table_categories_data[table_id] = None
            self.table_categories_row_mapper[table_id] = None
            return

        table_info = self.table_info(table_id=table_id)
        category_order = self.get_object(
            table_info.category_order.identifier,
            TSTArchives.CategoryOrderArchive,
        )
        row_uid_map = self.get_object(
            category_order.uid_map.identifier,
            TSTArchives.ColumnRowUIDMapArchive,
        )

        group_uuids = self.group_uuid_values()
        row_uuid_to_offset = {
            NumbersUUID(uuid): row for row, uuid in enumerate(category_archive.row_uid_lookup.uuids)
        }
        row_uid_for_index = [
            NumbersUUID(row_uid_map.sorted_row_uids[i]) for i in row_uid_map.row_uid_for_index
        ]

        def parent_relationships(
            parent: NumbersUUID | None,
            children: Sequence[TSTArchives.GroupByArchive.GroupNodeArchive],
            group_parents: dict[NumbersUUID, NumbersUUID | None],
        ) -> None:
            """Map model objects to their parent relationships."""
            for child in children:
                child_uuid = NumbersUUID(child.group_uid)
                group_parents[child_uuid] = parent
                if len(child.child) > 0:
                    parent_relationships(child_uuid, child.child, group_parents)

        group_parents: dict[NumbersUUID, NumbersUUID | None] = {}
        parent_relationships(None, category_archive.group_node_root.child, group_parents)

        row = 0
        row_mapper: dict[int, int] = {}
        nodes: dict[NumbersUUID, dict[str, Any]] = {}
        root_children: dict[Any, Any] = {}
        stack: list[NumbersUUID] = []

        for uuid in row_uid_for_index:
            if uuid in group_uuids:
                parent = group_parents.get(uuid)
                nodes[uuid] = {"key": group_uuids[uuid], "children": {}, "rows": []}

                if parent is None:
                    root_children[nodes[uuid]["key"]] = nodes[uuid]
                else:
                    parent_node = nodes[parent]
                    parent_node["children"][nodes[uuid]["key"]] = nodes[uuid]

                while stack and stack[-1] != parent:
                    stack.pop()
                stack.append(uuid)
            else:
                mapped_row = row_uuid_to_offset[uuid]
                if stack:
                    nodes[stack[-1]]["rows"].append(self._table_data[table_id][mapped_row])

                row_mapper[row] = mapped_row
                row += 1

        def node_to_structure(node: dict[str, Any]) -> dict[str, Any]:
            """Convert a protobuf node into a structured representation."""
            if not node["children"]:
                return node["rows"]
            out = {}
            for child_key, child_node in node["children"].items():
                out[child_key] = node_to_structure(child_node)
            return out

        self.table_categories_data[table_id] = {}
        for key, node in root_children.items():
            self.table_categories_data[table_id][key] = node_to_structure(node)

        self.table_categories_row_mapper[table_id] = {
            row: row_uuid_to_offset[uuid]
            for row, uuid in enumerate(
                uuid for uuid in row_uid_for_index if uuid not in group_uuids
            )
        }

    def refresh_pivot_tables(self) -> None:
        """
        Signal a pivot-table data refresh for the next document load.

        Bump refresh UUIDs and timestamps while preserving rule UUIDs to keep
        layout and styling unchanged.
        """
        for sheet_id in self.sheet_ids():
            for table_id in self.table_ids(sheet_id):
                if not self.is_a_pivot_table(table_id):
                    continue

                table_model = self.table_model(table_id)
                pivot_owner = self.get_object(
                    table_model.pivot_owner.identifier,
                    TSTArchives.PivotOwnerArchive,
                )

                # Only touch the data refresh trackers. Do not alter rule UUIDs.
                pivot_owner.refresh_uid.CopyFrom(NumbersUUID().protobuf2)
                pivot_owner.refresh_timestamp = (datetime.now() - EPOCH).total_seconds()  # noqa: DTZ005

    def update_table_metadata_ranges(  # noqa: C901
        self,
        table_id: int,
        num_rows: int,
        num_cols: int,
    ) -> None:
        """Keep UID maps and dependency bounds in sync with data dimensions."""
        table_model = self.table_model(table_id)

        def _update_map(uid_map_ref: TSPMessages.Reference) -> None:  # noqa: C901
            """Update map."""
            if not uid_map_ref or uid_map_ref.identifier == 0:
                return
            uid_map = self.get_object(uid_map_ref.identifier, TSTArchives.ColumnRowUIDMapArchive)

            col_uuids = [
                NumbersUUID(uid_map.sorted_column_uids[i]) for i in uid_map.column_uid_for_index
            ]
            row_uuids = [NumbersUUID(uid_map.sorted_row_uids[i]) for i in uid_map.row_uid_for_index]

            if num_cols > len(col_uuids):
                col_uuids.extend([NumbersUUID() for _ in range(num_cols - len(col_uuids))])
            elif num_cols < len(col_uuids):
                col_uuids = col_uuids[:num_cols]

            if num_rows > len(row_uuids):
                row_uuids.extend([NumbersUUID() for _ in range(num_rows - len(row_uuids))])
            elif num_rows < len(row_uuids):
                row_uuids = row_uuids[:num_rows]

            sorted_col_uuids = sorted(col_uuids, key=lambda x: x.int)
            sorted_row_uuids = sorted(row_uuids, key=lambda x: x.int)

            uid_map.sorted_column_uids.clear()
            for u in sorted_col_uuids:
                uid_map.sorted_column_uids.append(u.protobuf2)

            uid_map.column_uid_for_index.clear()
            uid_map.column_index_for_uid.clear()
            for _ in sorted_col_uuids:
                uid_map.column_index_for_uid.append(0)
            for col_idx, u in enumerate(col_uuids):
                sorted_idx = sorted_col_uuids.index(u)
                uid_map.column_uid_for_index.append(sorted_idx)
                uid_map.column_index_for_uid[sorted_idx] = col_idx

            uid_map.sorted_row_uids.clear()
            for u in sorted_row_uuids:
                uid_map.sorted_row_uids.append(u.protobuf2)

            uid_map.row_uid_for_index.clear()
            uid_map.row_index_for_uid.clear()
            for _ in sorted_row_uuids:
                uid_map.row_index_for_uid.append(0)
            for row_idx, u in enumerate(row_uuids):
                sorted_idx = sorted_row_uuids.index(u)
                uid_map.row_uid_for_index.append(sorted_idx)
                uid_map.row_index_for_uid[sorted_idx] = row_idx

        _update_map(table_model.base_column_row_uids)

        table_info = self.table_info(table_id=table_id)
        table_info_id = self.table_info_id(table_id)
        _update_map(table_info.view_column_row_uids)

        calc_engine = self.calc_engine()
        for dep_ref in calc_engine.dependency_tracker.formula_owner_dependencies:
            dep_archive = self.get_object(
                dep_ref.identifier,
                TSCEArchives.FormulaOwnerDependenciesArchive,
            )
            if (
                dep_archive.owner_kind == OwnerKind.TABLE_MODEL
                and dep_archive.formula_owner.identifier == table_info_id
            ):
                for field in ["spanning_column_dependencies", "spanning_row_dependencies"]:
                    spanning = getattr(dep_archive, field)
                    spanning.total_range_for_table.bottom_right_row = num_rows - 1
                    spanning.total_range_for_table.bottom_right_column = num_cols - 1
                    spanning.body_range_for_table.bottom_right_row = num_rows - 1
                    spanning.body_range_for_table.bottom_right_column = num_cols - 1


def rgb(obj: TSPMessages.Color) -> RGB:
    """Convert a TSPArchives.Color into an RGB tuple."""
    return RGB(round(obj.r * 255), round(obj.g * 255), round(obj.b * 255))


def range_end(
    obj: ASTColonTractArchive.ASTColonTractRelativeRangeArchive
    | ASTColonTractArchive.ASTColonTractAbsoluteRangeArchive,
) -> int:
    """Select end range for a IndexSetArchive.IndexSetEntry."""
    if obj.HasField("range_end"):
        return obj.range_end
    return obj.range_begin


def formatted_number(
    number_type: TSWPArchives.ListStyleArchive.LabelType.ValueType,
    index: int,
) -> str:
    """Format an index bullet for the requested type."""
    bullet_char = BULLET_PREFIXES[number_type]
    bullet_char += BULLET_CONVERSION[number_type](index)
    bullet_char += BULLET_SUFFIXES[number_type]

    return bullet_char


def get_storage_buffers_for_row(
    storage_buffer: bytes,
    offsets: bytes,
    num_cols: int,
    has_wide_offsets: bool,
) -> list[bytes | None]:
    """
    Extract storage buffers for each cell in a table row.

    Args:
    ----
        storage_buffer:  cell_storage_buffer or cell_storage_buffer for a table row
        offsets: 16-bit cell offsets for a table row
        num_cols: number of columns in a table row
        has_wide_offsets: use 4-byte offsets rather than 1-byte offset

    Returns:
    -------
         data: list of bytes for each cell in a row, or None if empty

    """
    offset_list = array("h", offsets).tolist()
    if has_wide_offsets:
        offset_list = [o * 4 for o in offset_list]

    data: list[bytes | None] = []
    for col in range(num_cols):
        if col >= len(offset_list):
            break

        start = offset_list[col]
        if start < 0:
            data.append(None)
            continue

        if col == (len(offset_list) - 1):
            end = len(storage_buffer)
        else:
            end = None
            # Find next positive offset
            for i, x in enumerate(offset_list[col + 1 :]):
                if x >= 0:
                    end = offset_list[col + i + 1]
                    break
            if end is None:
                end = len(storage_buffer)
        data.append(storage_buffer[start:end])

    return data


def field_references(obj: ProtoMessage) -> dict[str, Any]:
    """Return a dict of all fields in an object that are references to other objects."""
    return {
        x[0].name: {"identifier": getattr(obj, x[0].name).identifier}
        for x in obj.ListFields()
        if isinstance(getattr(obj, x[0].name), TSPMessages.Reference)
    }
