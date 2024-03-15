import re
from array import array
from collections import defaultdict
from hashlib import sha1
from pathlib import Path
from struct import pack
from typing import Dict, List, Tuple, Union

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
    xl_col_to_name,
    xl_rowcol_to_cell,
)
from numbers_parser.constants import (
    ALLOWED_FORMATTING_PARAMETERS,
    CUSTOM_FORMAT_TYPE_MAP,
    CUSTOM_TEXT_PLACEHOLDER,
    DEFAULT_COLUMN_WIDTH,
    DEFAULT_DOCUMENT,
    DEFAULT_PRE_BNC_BYTES,
    DEFAULT_ROW_HEIGHT,
    DEFAULT_TABLE_OFFSET,
    DEFAULT_TEXT_INSET,
    DEFAULT_TEXT_WRAP,
    DEFAULT_TILE_SIZE,
    DOCUMENT_ID,
    FORMAT_TYPE_MAP,
    MAX_TILE_SIZE,
    PACKAGE_ID,
    CellInteractionType,
    FormatType,
    OwnerKind,
)
from numbers_parser.containers import ObjectStore
from numbers_parser.exceptions import UnsupportedError
from numbers_parser.formula import TableFormulas
from numbers_parser.generated import TNArchives_pb2 as TNArchives
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.generated import TSDArchives_pb2 as TSDArchives
from numbers_parser.generated import TSKArchives_pb2 as TSKArchives
from numbers_parser.generated import TSPArchiveMessages_pb2 as TSPArchiveMessages
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages
from numbers_parser.generated import TSSArchives_pb2 as TSSArchives
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.generated import TSWPArchives_pb2 as TSWPArchives
from numbers_parser.generated.fontmap import FONT_NAME_TO_FAMILY
from numbers_parser.generated.TSDArchives_pb2 import (
    StrokePatternArchive as StrokePattern,
)
from numbers_parser.generated.TSWPArchives_pb2 import (
    CharacterStylePropertiesArchive as CharacterStyle,
)
from numbers_parser.iwafile import find_extension
from numbers_parser.numbers_cache import Cacheable, cache
from numbers_parser.numbers_uuid import NumbersUUID


def create_font_name_map(font_map: dict) -> dict:
    new_font_map = {}
    for k, v in font_map.items():
        if v not in new_font_map:
            new_font_map[v] = k
    return new_font_map


FONT_FAMILY_TO_NAME = create_font_name_map(FONT_NAME_TO_FAMILY)


class MergeCells:
    def __init__(self):
        self._references = defaultdict(lambda: False)

    def add_reference(self, row: int, col: int, rect: Tuple):
        self._references[(row, col)] = MergeReference(*rect)

    def add_anchor(self, row: int, col: int, size: Tuple):
        self._references[(row, col)] = MergeAnchor(size)

    def is_merge_reference(self, row_col: Tuple) -> bool:
        # defaultdict will default this to False for missing entries
        return isinstance(self._references[row_col], MergeReference)

    def is_merge_anchor(self, row_col: Tuple) -> bool:
        # defaultdict will default this to False for missing entries
        return isinstance(self._references[row_col], MergeAnchor)

    def get(self, row_col: Tuple) -> Union[MergeAnchor, MergeReference]:
        return self._references[row_col]

    def size(self, row_col: Tuple) -> Tuple:
        return self._references[row_col].size

    def rect(self, row_col: Tuple) -> Tuple:
        return self._references[row_col].rect

    def merge_cells(self):
        return [k for k, v in self._references.items() if self.is_merge_anchor(k)]


class DataLists(Cacheable):
    """Model for TST.DataList with caching and key generation for new values."""

    def __init__(self, model: object, datalist_name: str, value_attr: str = None):
        self._model = model
        self._datalists = {}
        self._value_attr = value_attr
        self._datalist_name = datalist_name

    @cache()
    def add_table(self, table_id: int):
        """Cache a new datalist for a table if not already seen."""
        base_data_store = self._model.objects[table_id].base_data_store
        datalist_id = getattr(base_data_store, self._datalist_name).identifier
        datalist = self._model.objects[datalist_id]

        max_key = 0
        self._datalists[table_id] = {}
        self._datalists[table_id]["by_key"] = {}
        self._datalists[table_id]["by_value"] = {}
        self._datalists[table_id]["datalist"] = datalist
        self._datalists[table_id]["id"] = datalist_id
        for entry in datalist.entries:
            if entry.key > max_key:
                max_key = entry.key
                self._datalists[table_id]["by_key"][entry.key] = entry
                value_key = self.value_key(getattr(entry, self._value_attr))
                self._datalists[table_id]["by_value"][value_key] = entry.key
        self._datalists[table_id]["next_key"] = max_key + 1

    def id(self, table_id: int) -> int:
        self.add_table(table_id)
        return self._datalists[table_id]["id"]

    def lookup_value(self, table_id: int, key: int):
        """Return the an entry in a table's datalist matching a key."""
        self.add_table(table_id)
        return self._datalists[table_id]["by_key"][key]

    def value_key(self, value):
        if hasattr(value, "DESCRIPTOR"):
            return repr(value)
        else:
            return value

    def init(self, table_id: int):
        """Remove all entries from a datalist."""
        self.add_table(table_id)
        self._datalists[table_id]["by_key"] = {}
        self._datalists[table_id]["by_value"] = {}
        self._datalists[table_id]["next_key"] = 1
        self._datalists[table_id]["datalist"].nextListID = 1
        clear_field_container(self._datalists[table_id]["datalist"].entries)

    def lookup_key(self, table_id: int, value) -> int:
        """Return the key associated with a value for a particular table entry.
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
            self._datalists[table_id]["by_value"][value_key] = key
        else:
            value_key = self.value_key(value)
            key = self._datalists[table_id]["by_value"][value_key]
            self._datalists[table_id]["by_key"][key].refcount += 1

        return key


class _NumbersModel(Cacheable):
    """Loads all objects from Numbers document and provides decoding
    methods for other classes in the module to abstract away the
    internal structures of Numbers document data structures.

    Not to be used in application code.
    """

    def __init__(self, filepath: Path):
        if filepath is None:
            filepath = Path(DEFAULT_DOCUMENT)
        self.objects = ObjectStore(filepath)
        self._merge_cells = defaultdict(MergeCells)
        self._row_heights = {}
        self._col_widths = {}
        self._table_formats = DataLists(self, "format_table", "format")
        self._table_styles = DataLists(self, "styleTable", "reference")
        self._table_strings = DataLists(self, "stringTable", "string")
        self._control_specs = DataLists(self, "control_cell_spec_table", "cell_spec")
        self._table_data = {}
        self._styles = None
        self._images = {}
        self._custom_formats = None
        self._custom_format_archives = None
        self._custom_format_ids = None
        self._strokes = {
            "top": defaultdict(),
            "right": defaultdict(),
            "bottom": defaultdict(),
            "left": defaultdict(),
        }

    def save(self, filepath: Path, package: bool) -> None:
        self.objects.save(filepath, package)

    def find_refs(self, ref: str) -> list:
        return self.objects.find_refs(ref)

    def sheet_ids(self):
        return [o.identifier for o in self.objects[DOCUMENT_ID].sheets]

    def sheet_name(self, sheet_id, value=None):
        if value is None:
            if sheet_id not in self.objects:
                return None
            else:
                return self.objects[sheet_id].name
        else:
            self.objects[sheet_id].name = value

    def set_table_data(self, table_id: int, data: List):
        self._table_data[table_id] = data

    # Don't cache: new tables can be added at runtime
    def table_ids(self, sheet_id: int) -> list:
        """Return a list of table IDs for a given sheet ID."""
        table_info_ids = self.find_refs("TableInfoArchive")
        return [
            self.objects[t_id].tableModel.identifier
            for t_id in table_info_ids
            if self.objects[t_id].super.parent.identifier == sheet_id
        ]

    # Don't cache: new tables can be added at runtime
    def table_info_id(self, table_id: int) -> int:
        """Return the TableInfoArchive ID for a given table ID."""
        ids = [
            x
            for x in self.objects.find_refs("TableInfoArchive")
            if self.objects[x].tableModel.identifier == table_id
        ]
        return ids[0]

    @cache()
    def row_storage_map(self, table_id):
        # The base data store contains a reference to rowHeaders.buckets
        # which is an ordered list that matches the storage buffers, but
        # identifies which row a storage buffer belongs to (empty rows have
        # no storage buffers). Each bucket is:
        #
        #  {
        #      "hiding_state": 0,
        #      "index": 0,
        #      "number_of_cells": 3,
        #      "size": 0.0
        #  },
        row_bucket_map = {i: None for i in range(self.objects[table_id].number_of_rows)}
        bds = self.objects[table_id].base_data_store
        bucket_ids = [x.identifier for x in bds.rowHeaders.buckets]
        idx = 0
        for bucket_id in bucket_ids:
            for header in self.objects[bucket_id].headers:
                row_bucket_map[header.index] = idx
                idx += 1
        return row_bucket_map

    def number_of_rows(self, table_id, num_rows=None):
        if num_rows is not None:
            self.objects[table_id].number_of_rows = num_rows
        return self.objects[table_id].number_of_rows

    def number_of_columns(self, table_id, num_cols=None):
        if num_cols is not None:
            self.objects[table_id].number_of_columns = num_cols
        return self.objects[table_id].number_of_columns

    @cache()
    def col_storage_map(self, table_id: int):
        # The base data store contains a reference to columnHeaders
        # which is an ordered list that identfies which offset to use
        # to index storage buffers for each column.
        #
        #  {
        #      "hiding_state": 0,
        #      "index": 0,
        #      "number_of_cells": 3,
        #      "size": 0.0
        #  },
        col_bucket_map = {i: None for i in range(self.objects[table_id].number_of_columns)}
        bds = self.objects[table_id].base_data_store
        buckets = self.objects[bds.columnHeaders.identifier].headers
        for i, bucket in enumerate(buckets):
            col_bucket_map[bucket.index] = i
        return col_bucket_map

    def table_name(self, table_id, value=None):
        if value is None:
            return self.objects[table_id].table_name
        else:
            self.objects[table_id].table_name = value

    def table_name_enabled(self, table_id: int, enabled: bool = None):
        if enabled is not None:
            self.objects[table_id].table_name_enabled = enabled
        else:
            return self.objects[table_id].table_name_enabled

    @cache()
    def table_tiles(self, table_id):
        bds = self.objects[table_id].base_data_store
        return [self.objects[t.tile.identifier] for t in bds.tiles.tiles]

    @cache(num_args=0)
    def custom_format_map(self):
        custom_format_list_id = self.objects[DOCUMENT_ID].super.custom_format_list.identifier
        custom_format_list = self.objects[custom_format_list_id]
        return {
            NumbersUUID(u).hex: custom_format_list.custom_formats[i]
            for i, u in enumerate(custom_format_list.uuids)
        }

    @cache(num_args=2)
    def table_format(self, table_id: int, key: int) -> str:
        """Return the format associated with a format ID for a particular table."""
        return self._table_formats.lookup_value(table_id, key).format

    @cache(num_args=3)
    def format_archive(self, table_id: int, format_type: FormattingType, format: Formatting):
        """Create a table format from a Formatting spec and return the table format ID"""
        attrs = {x: getattr(format, x) for x in ALLOWED_FORMATTING_PARAMETERS[format_type]}
        attrs["format_type"] = FORMAT_TYPE_MAP[format_type]

        format = TSKArchives.FormatStructArchive(**attrs)
        return self._table_formats.lookup_key(table_id, format)

    def cell_popup_model(self, parent_id: int, format: Formatting):
        tsce_items = [{"cell_value_type": "NIL_TYPE"}]
        for item in format.popup_values:
            if isinstance(item, str):
                tsce_items.append(
                    {
                        "cell_value_type": "STRING_TYPE",
                        "string_value": {
                            "value": item,
                            "format": {"format_type": FormatType.TEXT},
                        },
                    }
                )
            else:
                tsce_items.append(
                    {
                        "cell_value_type": "NUMBER_TYPE",
                        "number_value": {
                            "value": item,
                            "format": {"format_type": FormatType.DECIMAL},
                        },
                    }
                )
        popup_menu_id, _ = self.objects.create_object_from_dict(
            f"Index/Tables/DataList-{parent_id}",
            {"tsce_item": tsce_items},
            TSTArchives.PopUpMenuModel,
            True,
        )
        return popup_menu_id

    def control_cell_archive(self, table_id: int, format_type: FormattingType, format: Formatting):
        """Create control cell archive from a Formatting spec and return the table format ID"""
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
                range_control_min=format.minimum,
                range_control_max=format.maximum,
                range_control_inc=format.increment,
            )
        else:  # POPUP
            popup_id = self.cell_popup_model(self._control_specs.id(table_id), format)
            cell_spec = TSTArchives.CellSpecArchive(
                interaction_type=CellInteractionType.POPUP,
                chooser_control_popup_model=TSPMessages.Reference(identifier=popup_id),
                chooser_control_start_w_first=not (format.allow_none),
            )
        return self._control_specs.lookup_key(table_id, cell_spec)

    def add_custom_decimal_format_archive(self, format: CustomFormatting) -> None:
        """Create a custom format from the format spec"""
        integer_format = format.integer_format
        decimal_format = format.decimal_format
        num_integers = format.num_integers
        num_decimals = format.num_decimals
        show_thousands_separator = format.show_thousands_separator

        if num_integers == 0:
            format_string = ""
        elif integer_format == PaddingType.NONE:
            format_string = "#" * num_integers
        else:
            format_string = "0" * num_integers
        if num_integers > 6:
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
            name=format.name,
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
        self.add_custom_format_archive(format, format_archive)

    def add_custom_datetime_format_archive(self, format: CustomFormatting) -> None:
        format_archive = TSKArchives.CustomFormatArchive(
            name=format.name,
            format_type_pre_bnc=FormatType.CUSTOM_DATE,
            format_type=FormatType.CUSTOM_DATE,
            default_format=TSKArchives.FormatStructArchive(
                custom_format_string=format.format,
                format_type=FormatType.CUSTOM_DATE,
            ),
        )
        self.add_custom_format_archive(format, format_archive)

    def add_custom_format_archive(self, format: CustomFormatting, format_archive: object) -> None:
        format_uuid = NumbersUUID().protobuf2
        self._custom_formats[format.name] = format
        self._custom_format_archives[format.name] = format_archive
        self._custom_format_uuids[format.name] = format_uuid

        custom_format_list_id = self.objects[DOCUMENT_ID].super.custom_format_list.identifier
        custom_format_list = self.objects[custom_format_list_id]
        custom_format_list.custom_formats.append(format_archive)
        custom_format_list.uuids.append(format_uuid)

    def custom_format_id(self, table_id: int, format: CustomFormatting) -> int:
        """Look up the custom format and return the format ID for the table"""
        format_type = CUSTOM_FORMAT_TYPE_MAP[format.type]
        format_uuid = self._custom_format_uuids[format.name]
        custom_format = TSKArchives.FormatStructArchive(
            format_type=format_type,
            custom_uid=TSPMessages.UUID(lower=format_uuid.lower, upper=format_uuid.upper),
        )
        return self._table_formats.lookup_key(table_id, custom_format)

    def add_custom_text_format_archive(self, format: CustomFormatting) -> None:
        format_string = format.format.replace("%s", CUSTOM_TEXT_PLACEHOLDER)
        format_archive = TSKArchives.CustomFormatArchive(
            name=format.name,
            format_type_pre_bnc=FormatType.CUSTOM_TEXT,
            format_type=FormatType.CUSTOM_TEXT,
            default_format=TSKArchives.FormatStructArchive(
                custom_format_string=format_string,
                format_type=FormatType.CUSTOM_TEXT,
            ),
        )
        self.add_custom_format_archive(format, format_archive)

    @cache(num_args=2)
    def table_style(self, table_id: int, key: int) -> str:
        """Return the style associated with a style ID for a particular table."""
        style_entry = self._table_styles.lookup_value(table_id, key)
        return self.objects[style_entry.reference.identifier]

    @cache(num_args=2)
    def table_string(self, table_id: int, key: int) -> str:
        """Return the string associated with a string ID for a particular table."""
        return self._table_strings.lookup_value(table_id, key).string

    def init_table_strings(self, table_id: int):
        """Cache table strings reference and delete all existing keys/values."""
        self._table_strings.init(table_id)

    def table_string_key(self, table_id: int, value: str) -> int:
        """Return the key associated with a string for a particular table. If
        the string is not in the strings table, allocate a new entry with the
        next available key.
        """
        return self._table_strings.lookup_key(table_id, value)

    @cache(num_args=0)
    def owner_id_map(self):
        """ "
        Extracts the mapping table from Owner IDs to UUIDs. Returns a
        dictionary mapping the owner ID int to a 128-bit UUID.
        """
        # The TSCE.CalculationEngineArchive contains a list of mapping entries
        # in dependencyTracker.formulaOwnerDependencies from the root level
        # of the protobuf. Each mapping contains a 32-bit style UUID:
        #
        # "owner_id_map": {
        #     "map_entry": [
        #     {
        #         "internal_ownerId": 33,
        #         "owner_id": 0x3cb03f23_c26dda92_1e4bfcc0_8750e563
        #     },
        #
        #
        calc_engine = self.calc_engine()
        if calc_engine is None:
            return {}

        owner_id_map = {}
        for e in calc_engine.dependency_tracker.owner_id_map.map_entry:
            owner_id_map[e.internal_owner_id] = NumbersUUID(e.owner_id).hex
        return owner_id_map

    @cache()
    def table_base_id(self, table_id: int) -> int:
        """ "Finds the UUID of a table."""
        # Look for a TSCE.FormulaOwnerDependenciesArchive objects with the following at the
        # root level of the protobuf:
        #
        #     "base_owner_uid": "6a4a5281-7b06-f5a1-904b-7f9ec784b368"",
        #     "formula_owner_uid": "6a4a5281-7b06-f5a1-904b-7f9ec784b36d"
        #
        # The Table UUID is the TSCE.FormulaOwnerDependenciesArchive whose formula_owner_uid
        # matches the UUID of the haunted_owner of the Table:
        #
        #    "haunted_owner": {
        #        "owner_uid": "6a4a5281-7b06-f5a1-904b-7f9ec784b368""
        #    }
        haunted_owner = NumbersUUID(self.objects[table_id].haunted_owner.owner_uid).hex
        formula_owner_ids = self.find_refs("FormulaOwnerDependenciesArchive")
        for dependency_id in formula_owner_ids:  # pragma: no branch
            obj = self.objects[dependency_id]
            # if obj.owner_kind == OwnerKind.HAUNTED_OWNER:
            if obj.HasField("base_owner_uid") and obj.HasField(
                "formula_owner_uid"
            ):  # pragma: no branch
                base_owner_uid = NumbersUUID(obj.base_owner_uid).hex
                formula_owner_uid = NumbersUUID(obj.formula_owner_uid).hex
                if formula_owner_uid == haunted_owner:
                    return base_owner_uid

    @cache()
    def formula_cell_ranges(self, table_id: int) -> list:
        """Exract all the formula cell ranges for the Table."""
        # https://github.com/masaccio/numbers-parser/blob/main/doc/Numbers.md#formula-ranges
        calc_engine = self.calc_engine()

        table_base_id = self.table_base_id(table_id)
        cell_records = []
        for finfo in calc_engine.dependency_tracker.formula_owner_info:
            if finfo.HasField("cell_dependencies"):  # pragma: no branch
                formula_owner_id = NumbersUUID(finfo.formula_owner_id).hex
                if formula_owner_id == table_base_id:
                    for cell_record in finfo.cell_dependencies.cell_record:
                        if cell_record.contains_a_formula:  # pragma: no branch
                            cell_records.append((cell_record.row, cell_record.column))
        return cell_records

    @cache(num_args=0)
    def calc_engine_id(self):
        """Return the CalculationEngine ID for the current document."""
        ce_id = self.find_refs("CalculationEngineArchive")
        if len(ce_id) == 0:
            return 0
        else:
            return ce_id[0]

    @cache(num_args=0)
    def calc_engine(self):
        """Return the CalculationEngine object for the current document."""
        ce_id = self.calc_engine_id()
        if ce_id == 0:
            return None
        else:
            return self.objects[ce_id]

    @cache()
    def calculate_merge_cell_ranges(self, table_id):
        """Exract all the merge cell ranges for the Table."""
        # https://github.com/masaccio/numbers-parser/blob/main/doc/Numbers.md#merge-ranges
        owner_id_map = self.owner_id_map()
        table_base_id = self.table_base_id(table_id)

        formula_table_ids = self.find_refs("FormulaOwnerDependenciesArchive")
        for formula_id in formula_table_ids:
            dependencies = self.objects[formula_id]
            if dependencies.owner_kind != OwnerKind.MERGE_OWNER:
                continue
            for record in dependencies.range_dependencies.back_dependency:
                to_owner_id = record.internal_range_reference.owner_id
                if owner_id_map[to_owner_id] == table_base_id:
                    record_range = record.internal_range_reference.range
                    row_start = record_range.top_left_row
                    row_end = record_range.bottom_right_row
                    col_start = record_range.top_left_column
                    col_end = record_range.bottom_right_column

                    size = (row_end - row_start + 1, col_end - col_start + 1)
                    for row in range(row_start, row_end + 1):
                        for col in range(col_start, col_end + 1):
                            self._merge_cells[table_id].add_reference(
                                row, col, (row_start, col_start, row_end, col_end)
                            )
                    self._merge_cells[table_id].add_anchor(row_start, col_start, size)

        base_data_store = self.objects[table_id].base_data_store
        if base_data_store.merge_region_map.identifier == 0:
            return

        cell_ranges = self.objects[base_data_store.merge_region_map.identifier]
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
            for row in range(row_start, row_end + 1):
                for col in range(col_start, col_end + 1):
                    self._merge_cells[table_id].add_reference(
                        row, col, (row_start, col_start, row_end, col_end)
                    )
            self._merge_cells[table_id].add_anchor(row_start, col_start, (num_rows, num_columns))

    def merge_cells(self, table_id):
        self.calculate_merge_cell_ranges(table_id)
        return self._merge_cells[table_id]

    def table_id_to_sheet_id(self, table_id: int) -> int:
        for sheet_id in self.sheet_ids():  # pragma: no branch
            if table_id in self.table_ids(sheet_id):
                return sheet_id

    @cache()
    def table_uuids_to_id(self, table_uuid) -> int:
        for sheet_id in self.sheet_ids():  # pragma: no branch
            for table_id in self.table_ids(sheet_id):
                if table_uuid == self.table_base_id(table_id):
                    return table_id

    def node_to_ref(self, this_table_id: int, row: int, col: int, node):
        if node.HasField("AST_cross_table_reference_extra_info"):
            table_uuid = NumbersUUID(node.AST_cross_table_reference_extra_info.table_id).hex
            other_table_id = self.table_uuids_to_id(table_uuid)
            other_table_name = self.table_name(other_table_id)
        else:
            other_table_id = None
            other_table_name = None

        if other_table_id is not None:
            this_sheet_id = self.table_id_to_sheet_id(this_table_id)
            other_sheet_id = self.table_id_to_sheet_id(other_table_id)
            if this_sheet_id != other_sheet_id:
                other_sheet_name = self.sheet_name(other_sheet_id)
                other_table_name = f"{other_sheet_name}::" + other_table_name

        if node.HasField("AST_colon_tract"):
            return self.tract_to_row_col_ref(node, other_table_name, row, col)
        elif node.HasField("AST_row") and not node.HasField("AST_column"):
            return node_to_row_ref(node, other_table_name, row)
        elif node.HasField("AST_column") and not node.HasField("AST_row"):
            return node_to_col_ref(node, other_table_name, col)
        else:
            return node_to_row_col_ref(node, other_table_name, row, col)

    def tract_to_row_col_ref(self, node: object, table_name: str, row: int, col: int) -> str:
        if node.AST_sticky_bits.begin_row_is_absolute:
            row_begin = node.AST_colon_tract.absolute_row[0].range_begin
        else:
            row_begin = row + node.AST_colon_tract.relative_row[0].range_begin

        if node.AST_sticky_bits.end_row_is_absolute:
            row_end = range_end(node.AST_colon_tract.absolute_row[0])
        else:
            row_end = row + range_end(node.AST_colon_tract.relative_row[0])

        if node.AST_sticky_bits.begin_column_is_absolute:
            col_begin = node.AST_colon_tract.absolute_column[0].range_begin
        else:
            col_begin = col + node.AST_colon_tract.relative_column[0].range_begin

        if node.AST_sticky_bits.end_column_is_absolute:
            col_end = range_end(node.AST_colon_tract.absolute_column[0])
        else:
            col_end = col + range_end(node.AST_colon_tract.relative_column[0])

        begin_ref = xl_rowcol_to_cell(
            row_begin,
            col_begin,
            row_abs=node.AST_sticky_bits.begin_row_is_absolute,
            col_abs=node.AST_sticky_bits.begin_column_is_absolute,
        )
        end_ref = xl_rowcol_to_cell(
            row_end,
            col_end,
            row_abs=node.AST_sticky_bits.end_row_is_absolute,
            col_abs=node.AST_sticky_bits.end_column_is_absolute,
        )
        if table_name is not None:
            return f"{table_name}::{begin_ref}:{end_ref}"
        else:
            return f"{begin_ref}:{end_ref}"

    @cache()
    def formula_ast(self, table_id: int):
        bds = self.objects[table_id].base_data_store
        formula_table_id = bds.formula_table.identifier
        formula_table = self.objects[formula_table_id]
        formulas = {}
        for formula in formula_table.entries:
            formulas[formula.key] = formula.formula.AST_node_array.AST_node
        return formulas

    @cache()
    def storage_buffers(self, table_id: int) -> List:
        buffers = []
        for tile in self.table_tiles(table_id):
            if not tile.last_saved_in_BNC:
                raise UnsupportedError("Pre-BNC storage is unsupported")
            for r in tile.rowInfos:
                buffer = get_storage_buffers_for_row(
                    r.cell_storage_buffer,
                    r.cell_offsets,
                    self.number_of_columns(table_id),
                    r.has_wide_offsets,
                )
                buffers.append(buffer)
        return buffers

    @cache(num_args=3)
    def storage_buffer(self, table_id: int, row: int, col: int) -> bytes:
        row_offset = self.row_storage_map(table_id)[row]
        if row_offset is None:
            return None
        storage_buffers = self.storage_buffers(table_id)
        if row_offset >= len(storage_buffers):
            return None
        if col >= len(storage_buffers[row_offset]):
            return None
        return storage_buffers[row_offset][col]

    def recalculate_row_headers(self, table_id: int, data: List):
        base_data_store = self.objects[table_id].base_data_store
        buckets = self.objects[base_data_store.rowHeaders.buckets[0].identifier]
        clear_field_container(buckets.headers)
        for row in range(len(data)):
            if table_id in self._row_heights and row in self._row_heights[table_id]:
                height = self._row_heights[table_id][row]
            else:
                height = 0.0
            header = TSTArchives.HeaderStorageBucket.Header(
                index=row,
                numberOfCells=len(data[row]),
                size=height,
                hidingState=0,
            )
            buckets.headers.append(header)

    def recalculate_column_headers(self, table_id: int, data: List):
        current_column_widths = {}
        for col in range(self.number_of_columns(table_id)):
            current_column_widths[col] = self.col_width(table_id, col)

        base_data_store = self.objects[table_id].base_data_store
        buckets = self.objects[base_data_store.columnHeaders.identifier]
        clear_field_container(buckets.headers)
        # Transpose data to get columns
        col_data = [list(x) for x in zip(*data)]

        for col, cells in enumerate(col_data):
            num_rows = len(cells) - sum([isinstance(x, MergedCell) for x in cells])
            if table_id in self._col_widths and col in self._col_widths[table_id]:
                width = self._col_widths[table_id][col]
            else:
                width = current_column_widths[col]
            header = TSTArchives.HeaderStorageBucket.Header(
                index=col, numberOfCells=num_rows, size=width, hidingState=0
            )
            buckets.headers.append(header)

    def recalculate_merged_cells(self, table_id: int):
        merge_cells = self.merge_cells(table_id)

        merge_map_id, merge_map = self.objects.create_object_from_dict(
            "CalculationEngine", {}, TSTArchives.MergeRegionMapArchive
        )

        merge_cells = self.merge_cells(table_id)
        for row_col in merge_cells.merge_cells():
            size = merge_cells.size(row_col)
            cell_id = TSTArchives.CellID(packedData=(row_col[1] << 16 | row_col[0]))
            table_size = TSTArchives.TableSize(packedData=(size[1] << 16 | size[0]))
            cell_range = TSTArchives.CellRange(origin=cell_id, size=table_size)
            merge_map.cell_range.append(cell_range)

        base_data_store = self.objects[table_id].base_data_store
        base_data_store.merge_region_map.CopyFrom(TSPMessages.Reference(identifier=merge_map_id))

    def recalculate_row_info(
        self, table_id: int, data: List, tile_row_offset: int, row: int
    ) -> TSTArchives.TileRowInfo:
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
    def metadata_component(self, reference: Union[str, int] = None) -> int:
        """Return the ID of an object in the document metadata given it's name or ID."""
        component_map = {c.identifier: c for c in self.objects[PACKAGE_ID].components}
        if isinstance(reference, str):
            component_ids = [
                id for id, c in component_map.items() if c.preferred_locator == reference
            ]
        else:
            component_ids = [id for id, c in component_map.items() if c.identifier == reference]
        return component_map[component_ids[0]]

    def add_component_metadata(self, object_id: int, parent: str, locator: str):
        """Add a new ComponentInfo record to the parent object in the document metadata."""
        locator = locator.format(object_id)
        preferred_locator = re.sub(r"\-\d+.*", "", locator)
        component_info = TSPArchiveMessages.ComponentInfo(
            identifier=object_id,
            locator=locator,
            preferred_locator=preferred_locator,
            is_stored_outside_object_archive=False,
            document_read_version=[2, 0, 0],
            document_write_version=[2, 0, 0],
            save_token=1,
        )
        self.objects[PACKAGE_ID].components.append(component_info)
        self.add_component_reference(object_id, parent)

    def add_component_reference(
        self,
        object_id: int,
        location: str = None,
        parent_id: int = None,
        is_weak: bool = False,
    ):
        """Add an external reference to an object in a metadata component."""
        component = self.metadata_component(location or parent_id)
        if parent_id is not None:
            params = {"object_identifier": object_id, "component_identifier": parent_id}
        else:
            params = {"component_identifier": object_id}
        if is_weak:
            params["is_weak"] = True
        component.external_references.append(
            TSPArchiveMessages.ComponentExternalReference(**params)
        )

    def recalculate_table_data(self, table_id: int, data: List):
        table_model = self.objects[table_id]
        table_model.number_of_rows = len(data)
        table_model.number_of_columns = len(data[0])

        self.init_table_strings(table_id)
        self.recalculate_row_headers(table_id, data)
        self.recalculate_column_headers(table_id, data)
        self.recalculate_merged_cells(table_id)
        self.update_paragraph_styles()
        self.update_cell_styles(table_id, data)

        table_model.ClearField("base_column_row_uids")

        tile_idx = 0
        max_tile_idx = len(data) >> 8
        base_data_store = self.objects[table_id].base_data_store
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
                "Index/Tables/Tile-{}", tile_dict, TSTArchives.Tile
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

    def create_string_table(self):
        table_strings_id, table_strings = self.objects.create_object_from_dict(
            "Index/Tables/DataList-{}",
            {"listType": TSTArchives.TableDataList.ListType.STRING, "nextListID": 1},
            TSTArchives.TableDataList,
        )
        self.add_component_metadata(table_strings_id, "CalculationEngine", "Tables/DataList-{}")
        return table_strings_id, table_strings

    def table_height(self, table_id: int) -> int:
        """Return the height of a table in points."""
        table_model = self.objects[table_id]
        bds = self.objects[table_id].base_data_store
        buckets = self.objects[bds.rowHeaders.buckets[0].identifier].headers

        height = 0.0
        for i, row in self.row_storage_map(table_id).items():
            if table_id in self._row_heights and i in self._row_heights[table_id]:
                height += self._row_heights[table_id][i]
            elif row is not None and buckets[i].size != 0.0:
                height += buckets[i].size
            else:
                height += table_model.default_row_height
        return round(height)

    def row_height(self, table_id: int, row: int, height: int = None) -> int:
        if height is not None:
            if table_id not in self._row_heights:
                self._row_heights[table_id] = {}
            self._row_heights[table_id][row] = height
            return height

        if table_id in self._row_heights and row in self._row_heights[table_id]:
            return self._row_heights[table_id][row]

        table_model = self.objects[table_id]
        bds = self.objects[table_id].base_data_store
        bucket_id = bds.rowHeaders.buckets[0].identifier
        buckets = self.objects[bucket_id].headers
        bucket_map = {x.index: x for x in buckets}
        if row in bucket_map and bucket_map[row].size != 0.0:
            return round(bucket_map[row].size)
        else:
            return round(table_model.default_row_height)

    def table_width(self, table_id: int) -> int:
        """Return the width of a table in points."""
        table_model = self.objects[table_id]
        bds = self.objects[table_id].base_data_store
        buckets = self.objects[bds.columnHeaders.identifier].headers

        width = 0.0
        for i, col in self.col_storage_map(table_id).items():
            if table_id in self._col_widths and i in self._col_widths[table_id]:
                width += self._col_widths[table_id][i]
            elif col is not None and buckets[i].size != 0.0:
                width += buckets[i].size
            else:
                width += table_model.default_column_width
        return round(width)

    def col_width(self, table_id: int, col: int, width: int = None) -> int:
        if width is not None:
            if table_id not in self._col_widths:
                self._col_widths[table_id] = {}
            self._col_widths[table_id][col] = width
            return width

        if table_id in self._col_widths and col in self._col_widths[table_id]:
            return self._col_widths[table_id][col]

        table_model = self.objects[table_id]
        bds = self.objects[table_id].base_data_store
        bucket_id = bds.columnHeaders.identifier
        buckets = self.objects[bucket_id].headers
        bucket_map = {x.index: x for x in buckets}
        if col in bucket_map and bucket_map[col].size != 0.0:
            return round(bucket_map[col].size)
        else:
            return round(table_model.default_column_width)

    def num_header_rows(self, table_id: int, num_headers: int = None) -> int:
        """Return/set the number of header rows."""
        table_model = self.objects[table_id]
        if num_headers is not None:
            table_model.number_of_header_rows = num_headers
        return table_model.number_of_header_rows

    def num_header_cols(self, table_id: int, num_headers: int = None) -> int:
        """Return/set the number of header columns."""
        table_model = self.objects[table_id]
        if num_headers is not None:
            table_model.number_of_header_columns = num_headers
        return table_model.number_of_header_columns

    def table_coordinates(self, table_id: int) -> Tuple[float]:
        table_info = self.objects[self.table_info_id(table_id)]
        return (
            table_info.super.geometry.position.x,
            table_info.super.geometry.position.y,
        )

    def is_a_pivot_table(self, table_id: int) -> bool:
        """Table is a pivot table."""
        table_info = self.objects[self.table_info_id(table_id)]
        return table_info.is_a_pivot_table

    def last_table_offset(self, sheet_id):
        """Y offset of the last table in a sheet."""
        table_id = self.table_ids(sheet_id)[-1]
        y_offset = [
            self.objects[self.table_info_id(x)].super.geometry.position.y
            for x in self.table_ids(sheet_id)
            if x == table_id
        ][0]

        return self.table_height(table_id) + y_offset

    def create_drawable(self, sheet_id: int, x: float, y: float) -> object:
        """Create a DrawableArchive for a new table in a sheet."""
        table_x = x if x is not None else 0.0
        table_y = y if y is not None else self.last_table_offset(sheet_id) + DEFAULT_TABLE_OFFSET
        return TSDArchives.DrawableArchive(
            parent=TSPMessages.Reference(identifier=sheet_id),
            geometry=TSDArchives.GeometryArchive(
                angle=0.0,
                flags=3,
                position=TSPMessages.Point(x=table_x, y=table_y),
                size=TSPMessages.Size(height=231.0, width=494.0),
            ),
        )

    def add_table(  # noqa: PLR0913
        self,
        sheet_id: int,
        table_name: str,
        from_table_id: int,
        x: float,
        y: float,
        num_rows: int,
        num_cols: int,
        number_of_header_rows=1,
        number_of_header_columns=1,
    ) -> int:
        from_table = self.objects[from_table_id]

        table_strings_id, table_strings = self.create_string_table()

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
        # Supresses Numbers assertions for tables sharing the same data
        table_model.category_owner.identifier = 0

        column_headers_id, column_headers = self.objects.create_object_from_dict(
            "Index/Tables/HeaderStorageBucket-{}",
            {"bucketHashFunction": 1},
            TSTArchives.HeaderStorageBucket,
        )
        self.add_component_metadata(
            column_headers_id, "CalculationEngine", "Tables/HeaderStorageBucket-{}"
        )

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
            formula_table_id, "CalculationEngine", "Tables/TableDataList-{}"
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
                    tile_size=DEFAULT_TILE_SIZE, should_use_wide_rows=True
                ),
                **data_store_refs,
            )
        )

        row_headers_id, _ = self.objects.create_object_from_dict(
            "Index/Tables/HeaderStorageBucket-{}",
            {"bucketHashFunction": 1},
            TSTArchives.HeaderStorageBucket,
        )

        self.add_component_metadata(
            row_headers_id, "CalculationEngine", "Tables/HeaderStorageBucket-{}"
        )
        table_model.base_data_store.rowHeaders.buckets.append(
            TSPMessages.Reference(identifier=row_headers_id)
        )

        data = [
            [Cell._empty_cell(table_model_id, row, col, self) for col in range(num_cols)]
            for row in range(num_rows)
        ]
        self.recalculate_table_data(table_model_id, data)

        table_info_id, table_info = self.objects.create_object_from_dict(
            "CalculationEngine",
            {},
            TSTArchives.TableInfoArchive,
        )
        table_info.tableModel.MergeFrom(TSPMessages.Reference(identifier=table_model_id))
        table_info.super.MergeFrom(self.create_drawable(sheet_id, x, y))
        self.add_component_reference(table_info_id, "Document", self.calc_engine_id())

        self.add_formula_owner(
            table_info_id,
            num_rows,
            num_cols,
            number_of_header_rows,
            number_of_header_columns,
        )

        self.objects[sheet_id].drawable_infos.append(
            TSPMessages.Reference(identifier=table_info_id)
        )
        return table_model_id

    def add_formula_owner(  # noqa: PLR0913
        self,
        table_info_id: int,
        num_rows: int,
        num_cols: int,
        number_of_header_rows: int,
        number_of_header_columns: int,
    ):
        """Create a FormulaOwnerDependenciesArchive that references a TableInfoArchive
        so that cross-references to cells in this table will work.
        """
        formula_owner_uuid = NumbersUUID()
        calc_engine = self.calc_engine()
        owner_id_map = calc_engine.dependency_tracker.owner_id_map.map_entry
        next_owner_id = max([x.internal_owner_id for x in owner_id_map]) + 1
        formula_deps_id, formula_deps = self.objects.create_object_from_dict(
            "CalculationEngine",
            {
                "formula_owner_uid": formula_owner_uuid.dict2,
                "internal_formula_owner_id": next_owner_id,
                "owner_kind": 1,
                "cell_dependencies": {},
                "range_dependencies": {},
                "volatile_dependencies": {
                    "volatile_time_cells": {},
                    "volatile_random_cells": {},
                    "volatile_locale_cells": {},
                    "volatile_sheet_table_name_cells": {},
                    "volatile_remote_data_cells": {},
                    "volatile_geometry_cell_refs": {},
                },
                "spanning_column_dependencies": {
                    "total_range_for_table": {
                        "top_left_column": 0,
                        "top_left_row": 0,
                        "bottom_right_column": num_cols - 1,
                        "bottom_right_row": num_cols - 1,
                    },
                    "body_range_for_table": {
                        "top_left_column": number_of_header_columns,
                        "top_left_row": number_of_header_rows,
                        "bottom_right_column": num_cols - 1,
                        "bottom_right_row": num_cols - 1,
                    },
                },
                "spanning_row_dependencies": {
                    "total_range_for_table": {
                        "top_left_column": 0,
                        "top_left_row": 0,
                        "bottom_right_column": num_cols - 1,
                        "bottom_right_row": num_cols - 1,
                    },
                    "body_range_for_table": {
                        "top_left_column": number_of_header_columns,
                        "top_left_row": number_of_header_rows,
                        "bottom_right_column": num_cols - 1,
                        "bottom_right_row": num_cols - 1,
                    },
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
            TSPMessages.Reference(identifier=formula_deps_id)
        )
        owner_id_map.append(
            TSCEArchives.OwnerIDMapArchive.OwnerIDMapArchiveEntry(
                internal_owner_id=next_owner_id, owner_id=formula_owner_uuid.protobuf4
            )
        )

    def add_sheet(self, sheet_name: str) -> int:
        """Add a new sheet with a copy of a table from another sheet."""
        sheet_id, _ = self.objects.create_object_from_dict(
            "Document", {"name": sheet_name}, TNArchives.SheetArchive
        )

        self.add_component_reference(sheet_id, "CalculationEngine", DOCUMENT_ID, is_weak=True)

        self.objects[DOCUMENT_ID].sheets.append(TSPMessages.Reference(identifier=sheet_id))

        return sheet_id

    @property
    def styles(self):
        if self._styles is None:
            self._styles = self.available_paragraph_styles()
        return self._styles

    @cache(num_args=0)
    def available_paragraph_styles(self) -> Dict[str, Style]:
        theme_id = self.objects[DOCUMENT_ID].theme.identifier
        presets = find_extension(self.objects[theme_id].super, "paragraph_style_presets")
        presets_map = {
            self.objects[x.identifier].super.name: {
                "id": x.identifier,
                "obj": self.objects[x.identifier],
            }
            for x in presets
        }
        styles = {
            k: Style(
                alignment=Alignment(
                    HorizontalJustification(v["obj"].para_properties.alignment),
                    VerticalJustification(0),
                ),
                font_color=self.cell_font_color(v["obj"]),
                font_size=self.cell_font_size(v["obj"]),
                font_name=self.cell_font_name(v["obj"]),
                bold=self.cell_is_bold(v["obj"]),
                italic=self.cell_is_italic(v["obj"]),
                underline=self.cell_is_underline(v["obj"]),
                strikethrough=self.cell_is_strikethrough(v["obj"]),
                name=self.cell_style_name(v["obj"]),
                _text_style_obj_id=v["id"],
            )
            for k, v in presets_map.items()
        }
        for style in styles.values():
            # Override __setattr__ behaviour for builtin styles
            style.__dict__["_update_text_style"] = False
            style.__dict__["_update_cell_style"] = False
        return styles

    def add_paragraph_style(self, style: Style) -> int:
        if style.underline:
            underline = CharacterStyle.UnderlineType.kSingleUnderline
        else:
            underline = CharacterStyle.UnderlineType.kNoUnderline
        if style.strikethrough:
            strikethru = CharacterStyle.StrikethruType.kSingleStrikethru
        else:
            strikethru = CharacterStyle.StrikethruType.kNoStrikethru

        style_id_name = "numbers-parser-" + style.name.lower().replace(" ", "-")
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
                    "bold": style.bold,
                    "italic": style.italic,
                    "underline": underline,
                    "strikethru": strikethru,
                    "font_size": style.font_size,
                    "font_name": FONT_FAMILY_TO_NAME[style.font_name],
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
        stylesheet_id = self.objects[DOCUMENT_ID].stylesheet.identifier
        para_style.super.stylesheet.MergeFrom(TSPMessages.Reference(identifier=stylesheet_id))
        self.objects[stylesheet_id].styles.append(TSPMessages.Reference(identifier=para_style_id))
        self.objects[stylesheet_id].identifier_to_style_map.append(
            TSSArchives.StylesheetArchive.IdentifiedStyleEntry(
                identifier=style_id_name,
                style=TSPMessages.Reference(identifier=para_style_id),
            )
        )

        theme_id = self.objects[DOCUMENT_ID].theme.identifier
        presets = find_extension(self.objects[theme_id].super, "paragraph_style_presets")
        presets.append(TSPMessages.Reference(identifier=para_style_id))
        self._styles[style.name] = style
        return para_style_id

    def update_paragraph_style(self, style: Style):
        if style.underline:
            underline = CharacterStyle.UnderlineType.kSingleUnderline
        else:
            underline = CharacterStyle.UnderlineType.kNoUnderline
        if style.strikethrough:
            strikethru = CharacterStyle.StrikethruType.kSingleStrikethru
        else:
            strikethru = CharacterStyle.StrikethruType.kNoStrikethru
        style_obj = self.objects[style._text_style_obj_id]
        style_obj.char_properties.font_color.r = style.font_color.r / 255
        style_obj.char_properties.font_color.g = style.font_color.g / 255
        style_obj.char_properties.font_color.b = style.font_color.b / 255
        style_obj.char_properties.bold = style.bold
        style_obj.char_properties.italic = style.italic
        style_obj.char_properties.underline = underline
        style_obj.char_properties.strikethru = strikethru
        style_obj.char_properties.font_size = style.font_size
        style_obj.char_properties.font_name = FONT_FAMILY_TO_NAME[style.font_name]
        style_obj.char_properties.tsd_fill.color.r = style.font_color.r / 255
        style_obj.char_properties.tsd_fill.color.g = style.font_color.g / 255
        style_obj.char_properties.tsd_fill.color.b = style.font_color.b / 255
        style_obj.para_properties.alignment = style.alignment.horizontal
        style_obj.para_properties.first_line_indent = style.first_indent
        style_obj.para_properties.left_indent = style.left_indent
        style_obj.para_properties.right_indent = style.right_indent

    def update_paragraph_styles(self):
        """Create new paragraph style archives for any new styles that
        have been created for this document.
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

    def update_cell_styles(self, table_id: int, data: List):
        """Create new cell style archives for any cells whose styles
        have changes that require a cell style.
        """
        cell_styles = {}
        for _, cells in enumerate(data):
            for _, cell in enumerate(cells):
                if cell._style is not None and cell._style._update_cell_style:
                    fingerprint = (
                        str(cell.style.alignment.vertical)
                        + str(cell.style.first_indent)
                        + str(cell.style.left_indent)
                        + str(cell.style.right_indent)
                        + str(cell.style.text_inset)
                        + str(cell.style.text_wrap)
                    )
                    if cell._style.bg_color is not None:
                        fingerprint = fingerprint + (
                            str(cell.style.bg_color.r)
                            + str(cell.style.bg_color.g)
                            + str(cell.style.bg_color.b)
                        )
                    if cell._style.bg_image is not None:
                        fingerprint += cell._style.bg_image.filename
                    if fingerprint not in cell_styles:
                        cell_styles[fingerprint] = self.add_cell_style(cell._style)
                    cell._style._cell_style_obj_id = cell_styles[fingerprint]

    def add_cell_style(self, style: Style) -> int:
        if style.bg_image is not None:
            digest = sha1(style.bg_image.data).digest()
            if digest in self._images:
                image_id = self._images[digest]
            else:
                datas = self.objects[PACKAGE_ID].datas
                image_id = self.next_image_identifier()
                datas.append(
                    TSPArchiveMessages.DataInfo(
                        identifier=image_id,
                        digest=digest,
                        preferred_file_name=style.bg_image.filename,
                        file_name=style.bg_image.filename,
                        materialized_length=len(style.bg_image.data),
                    )
                )
                self._images[digest] = image_id
            color_attrs = {
                "cell_fill": {
                    "image": {
                        "technique": "ScaleToFill",
                        "imagedata": {"identifier": image_id},
                        "interpretsUntaggedImageDataAsGeneric": False,
                    }
                }
            }
        elif style.bg_color is not None:
            color_attrs = {
                "cell_fill": {
                    "color": {
                        "model": "rgb",
                        "r": style.bg_color.r / 255,
                        "g": style.bg_color.g / 255,
                        "b": style.bg_color.b / 255,
                        "a": 1.0,
                        "rgbspace": "srgb",
                    }
                }
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

        stylesheet_id = self.objects[DOCUMENT_ID].stylesheet.identifier
        cell_style.super.stylesheet.MergeFrom(TSPMessages.Reference(identifier=stylesheet_id))
        self.objects[stylesheet_id].styles.append(TSPMessages.Reference(identifier=cell_style_id))
        self.objects[stylesheet_id].identifier_to_style_map.append(
            TSSArchives.StylesheetArchive.IdentifiedStyleEntry(
                identifier=style_id_name,
                style=TSPMessages.Reference(identifier=cell_style_id),
            )
        )

        return cell_style_id

    def text_style_object_id(self, cell: Cell) -> int:
        if cell._text_style_id is None:
            return None
        entry = self._table_styles.lookup_value(cell._table_id, cell._text_style_id)
        return entry.reference.identifier

    def cell_style_object_id(self, cell: Cell) -> int:
        if cell._cell_style_id is None:
            return None
        entry = self._table_styles.lookup_value(cell._table_id, cell._cell_style_id)
        return entry.reference.identifier

    def custom_style_name(self) -> str:
        """Find custom styles in the current document and return the next
        highest numbered style.
        """
        stylesheet_id = self.objects[DOCUMENT_ID].stylesheet.identifier
        current_styles = self.styles.keys()
        custom_styles = [x for x in current_styles if re.fullmatch(r"Custom Style \d+", x)]
        for style_entry in self.objects[stylesheet_id].identifier_to_style_map:
            style_id = style_entry.style.identifier
            style_name = getattr(self.objects[style_id].super, "name", "")
            if re.fullmatch(r"Custom Style \d+", style_name):
                custom_styles.append(style_name)

        if len(custom_styles) > 0:
            offset = len("Custom Style ")
            custom_style_ids = [int(x[offset:]) for x in custom_styles]
            return "Custom Style " + str(custom_style_ids[-1] + 1)
        else:
            return "Custom Style 1"

    @property
    def custom_formats(self) -> Dict[str, CustomFormatting]:
        if self._custom_formats is None:
            custom_format_list_id = self.objects[DOCUMENT_ID].super.custom_format_list.identifier
            custom_formats = self.objects[custom_format_list_id].custom_formats
            custom_format_names = [x.name for x in custom_formats]
            custom_format_uuids = [x for x in self.objects[custom_format_list_id].uuids]
            self._custom_formats = {}
            self._custom_format_archives = {}
            self._custom_format_uuids = {}
            for i, format_name in enumerate(custom_format_names):
                self._custom_formats[format_name] = CustomFormatting.from_archive(custom_formats[i])
                self._custom_format_archives[format_name] = custom_formats[i]
                self._custom_format_uuids[format_name] = custom_format_uuids[i]

        return self._custom_formats

    def custom_format_name(self) -> str:
        """Find custom formats in the current document and return the next
        highest numbered format.
        """
        current_formats = self.custom_formats.keys()
        if "Custom Format" not in current_formats:
            return "Custom Format"
        current_formats = [
            m.group(1) for x in current_formats if (m := re.fullmatch(r"Custom Format (\d+)", x))
        ]
        if len(current_formats) > 0:
            last_id = int(current_formats[-1])
            return f"Custom Format {last_id + 1}"
        else:
            return "Custom Format 1"

    @cache()
    def table_formulas(self, table_id: int):
        return TableFormulas(self, table_id)

    @cache(num_args=2)
    def table_rich_text(self, table_id: int, string_key: int) -> Dict:
        """Extract bullets and hyperlinks from a rich text data cell."""
        # The table model base data store contains a richTextTable field
        # which is a reference to a TST.TableDataList. The TableDataList
        # has a list of payloads in a field called entries. This will be
        # empty if there is no rich text, i.e. text contents are plaintext.
        #
        # "entries": [
        #     { "key": 1,
        #       "refcount": 1,
        #       "richTextPayload": { "identifier": "2035264" }
        #     },
        #     ...
        #
        # entries[n].richTextPayload.identifier is a reference to a
        # TST.RichTextPayloadArchive that contains a field called storage
        # that itself is a reference to a TSWP.StorageArchive that contains
        # the actual paragraph data:
        #
        # "tableParaStyle": {
        #     "entries": [
        #         { "characterIndex": 0, "object": { "identifier": "1566948" } },
        #         { "characterIndex": 6 },
        #         { "characterIndex": 12 }
        #     ]
        # },
        # "text": [ "Lorem\nipsum\ndolor" ]
        #
        # The bullet character is stored in a TSWP.ListStyleArchive. Each bullet
        # paragraph can have its own reference to a list style or, if none is
        # defined, the previous bullet character is used. All StorageArchives
        # reference a ListStyleArchive but not all those ListStyleArchives have
        # a string with a new bullet character
        bds = self.objects[table_id].base_data_store
        rich_text_table = self.objects[bds.rich_text_table.identifier]
        for entry in rich_text_table.entries:  # pragma: no branch
            if string_key == entry.key:
                payload = self.objects[entry.rich_text_payload.identifier]
                payload_storage = self.objects[payload.storage.identifier]
                smartfield_entries = payload_storage.table_smartfield.entries
                cell_text = payload_storage.text[0]

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
                for i, offset in enumerate(offsets):
                    if i == len(offsets) - 1:
                        bullets.append(cell_text[offset:])
                    else:
                        # Remove the last character (always newline)
                        bullets.append(cell_text[offset : offsets[i + 1] - 1])

                    # Re-use last style if there is none defined for this bullet
                    if i < len(table_list_styles):
                        table_list_style = table_list_styles[i]

                    bullet_style = self.objects[table_list_style.object.identifier]
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

    def cell_text_style(self, cell: Cell) -> object:
        """Return the text style object for the cell or, if none
        is defined, the default header, footer or body style.
        """
        if cell._text_style_id is not None:
            return self.table_style(cell._table_id, cell._text_style_id)

        table_model = self.objects[cell._table_id]
        if cell.row in range(table_model.number_of_header_rows):
            return self.objects[table_model.header_row_text_style.identifier]
        elif cell.col in range(table_model.number_of_header_columns):
            return self.objects[table_model.header_column_text_style.identifier]
        elif table_model.number_of_footer_rows > 0:
            start_row_num = table_model.number_of_rows - table_model.number_of_footer_rows
            end_row_num = start_row_num + table_model.number_of_footer_rows
            if cell.row in range(start_row_num, end_row_num):
                return self.objects[table_model.footer_row_text_style.identifier]
        return self.objects[table_model.body_text_style.identifier]

    def cell_alignment(self, cell: Cell) -> Alignment:
        style = self.cell_text_style(cell)
        horizontal = HorizontalJustification(self.para_property(style, "alignment"))

        if cell._cell_style_id is None:
            vertical = VerticalJustification.TOP
        else:
            style = self.table_style(cell._table_id, cell._cell_style_id)
            vertical = VerticalJustification(self.cell_property(style, "vertical_alignment"))
        return Alignment(horizontal, vertical)

    def cell_bg_color(self, cell: Cell) -> Union[Tuple, List[Tuple]]:
        if cell._cell_style_id is None:
            return None

        style = self.table_style(cell._table_id, cell._cell_style_id)
        cell_properties = style.cell_properties.cell_fill

        if cell_properties.HasField("color"):
            return rgb(cell_properties.color)
        elif cell_properties.HasField("gradient"):
            return [(rgb(s.color)) for s in cell_properties.gradient.stops]

    def char_property(self, style: object, field: str):
        """Return a char_property field from a style if present
        in the style, or from the parent if not.
        """
        if not style.char_properties.HasField(field):
            parent = self.objects[style.super.parent.identifier]
            return getattr(parent.char_properties, field)
        else:
            return getattr(style.char_properties, field)

    def para_property(self, style: object, field: str) -> float:
        """Return a para_property field from a style if present
        in the style, or from the parent if not.
        """
        if not style.para_properties.HasField(field):
            parent = self.objects[style.super.parent.identifier]
            return getattr(parent.para_properties, field)
        else:
            return getattr(style.para_properties, field)

    def cell_property(self, style: object, field: str) -> float:
        """Return a cell_property field from a style if present
        in the style, or from the parent if not.
        """
        if not style.cell_properties.HasField(field):
            parent = self.objects[style.super.parent.identifier]
            return getattr(parent.cell_properties, field)
        else:
            return getattr(style.cell_properties, field)

    def cell_is_bold(self, obj: Union[Cell, object]) -> bool:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.char_property(style, "bold")

    def cell_is_italic(self, obj: Union[Cell, object]) -> bool:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.char_property(style, "italic")

    def cell_is_underline(self, obj: Union[Cell, object]) -> bool:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        underline = self.char_property(style, "underline")
        return underline != CharacterStyle.UnderlineType.kNoUnderline

    def cell_is_strikethrough(self, obj: Union[Cell, object]) -> bool:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        strikethru = self.char_property(style, "strikethru")
        return strikethru != CharacterStyle.StrikethruType.kNoStrikethru

    def cell_style_name(self, obj: Union[Cell, object]) -> bool:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return style.super.name

    def cell_font_color(self, obj: Union[Cell, object]) -> Tuple:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return rgb(self.char_property(style, "font_color"))

    def cell_font_size(self, obj: Union[Cell, object]) -> float:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.char_property(style, "font_size")

    def cell_font_name(self, obj: Union[Cell, object]) -> str:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        font_name = self.char_property(style, "font_name")
        return FONT_NAME_TO_FAMILY[font_name]

    def cell_first_indent(self, obj: Union[Cell, object]) -> float:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.para_property(style, "first_line_indent")

    def cell_left_indent(self, obj: Union[Cell, object]) -> float:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.para_property(style, "left_indent")

    def cell_right_indent(self, obj: Union[Cell, object]) -> float:
        style = self.cell_text_style(obj) if isinstance(obj, Cell) else obj
        return self.para_property(style, "right_indent")

    def cell_text_inset(self, cell: Cell) -> float:
        if cell._cell_style_id is None:
            return DEFAULT_TEXT_INSET
        else:
            style = self.table_style(cell._table_id, cell._cell_style_id)
            padding = self.cell_property(style, "padding")
            # Padding is always identical (only one UI setting)
            return padding.left

    def cell_text_wrap(self, cell: Cell) -> float:
        if cell._cell_style_id is None:
            return DEFAULT_TEXT_WRAP
        else:
            style = self.table_style(cell._table_id, cell._cell_style_id)
            return self.cell_property(style, "text_wrap")

    def stroke_type(self, stroke_run: object) -> str:
        """Return the stroke type for a stroke run."""
        stroke_type = stroke_run.stroke.pattern.type
        if stroke_type == StrokePattern.StrokePatternType.TSDSolidPattern:
            return "solid"
        elif stroke_type == StrokePattern.StrokePatternType.TSDPattern:
            if stroke_run.stroke.pattern.pattern[0] < 1.0:
                return "dots"
            else:
                return "dashes"
        else:
            return "none"

    def cell_for_stroke(self, table_id: int, side: str, row: int, col: int) -> object:
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
            if side in ["top", "left"]:
                return cell
        else:
            return cell
        return None

    def set_cell_border(  # noqa: PLR0913
        self, table_id: int, row: int, col: int, side: str, border_value: Border
    ):
        """Set the 2 borders adjacent to a stroke if within the table range."""
        if side == "top":
            if (cell := self.cell_for_stroke(table_id, "top", row, col)) is not None:
                cell._border.top = border_value
            if (cell := self.cell_for_stroke(table_id, "bottom", row - 1, col)) is not None:
                cell._border.bottom = border_value
        elif side == "right":
            if (cell := self.cell_for_stroke(table_id, "right", row, col)) is not None:
                cell._border.right = border_value
            if (cell := self.cell_for_stroke(table_id, "left", row, col + 1)) is not None:
                cell._border.left = border_value
        elif side == "bottom":
            if (cell := self.cell_for_stroke(table_id, "bottom", row, col)) is not None:
                cell._border.bottom = border_value
            if (cell := self.cell_for_stroke(table_id, "top", row + 1, col)) is not None:
                cell._border.top = border_value
        else:  # left border
            if (cell := self.cell_for_stroke(table_id, "left", row, col)) is not None:
                cell._border.left = border_value
            if (cell := self.cell_for_stroke(table_id, "right", row, col - 1)) is not None:
                cell._border.right = border_value

    def extract_strokes_in_layers(self, table_id: int, layer_ids: List, side: str):
        for layer_id in layer_ids:
            stroke_layer = self.objects[layer_id.identifier]
            for stroke_run in stroke_layer.stroke_runs:
                border_value = Border(
                    width=round(stroke_run.stroke.width, 2),
                    color=rgb(stroke_run.stroke.color),
                    style=self.stroke_type(stroke_run),
                    _order=stroke_run.order,
                )
                if side in ["top", "bottom"]:
                    start_row = stroke_layer.row_column_index
                    start_column = stroke_run.origin
                    for col in range(start_column, start_column + stroke_run.length):
                        self.set_cell_border(table_id, start_row, col, side, border_value)
                else:
                    start_row = stroke_run.origin
                    start_column = stroke_layer.row_column_index
                    for row in range(start_row, start_row + stroke_run.length):
                        self.set_cell_border(table_id, row, start_column, side, border_value)

    @cache()
    def extract_strokes(self, table_id: int):
        table_obj = self.objects[table_id]
        sidecar_obj = self.objects[table_obj.stroke_sidecar.identifier]
        self.extract_strokes_in_layers(table_id, sidecar_obj.top_row_stroke_layers, "top")
        self.extract_strokes_in_layers(table_id, sidecar_obj.left_column_stroke_layers, "left")
        self.extract_strokes_in_layers(table_id, sidecar_obj.right_column_stroke_layers, "right")
        self.extract_strokes_in_layers(table_id, sidecar_obj.bottom_row_stroke_layers, "bottom")

    def create_stroke(self, origin: int, length: int, border_value: Border):
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
            order=border_value._order,
            stroke=TSDArchives.StrokeArchive(
                color=color,
                width=width,
                cap=line_cap,
                join=line_join,
                miter_limit=4.0,
                pattern=pattern,
            ),
        )

    def add_stroke(  # noqa: PLR0913
        self,
        table_id: int,
        row: int,
        col: int,
        side: str,
        border_value: Border,
        length: int,
    ):
        table_obj = self.objects[table_id]
        sidecar_obj = self.objects[table_obj.stroke_sidecar.identifier]
        sidecar_obj.max_order += 1
        sidecar_obj.row_count = table_obj.number_of_rows
        sidecar_obj.column_count = table_obj.number_of_columns
        border_value._order = sidecar_obj.max_order

        if side == "top":
            layer_ids = sidecar_obj.top_row_stroke_layers
            row_column_index = row
            origin = col
        elif side == "right":
            layer_ids = sidecar_obj.right_column_stroke_layers
            row_column_index = col
            origin = row
        elif side == "bottom":
            layer_ids = sidecar_obj.bottom_row_stroke_layers
            row_column_index = row
            origin = col
        else:  # left border
            layer_ids = sidecar_obj.left_column_stroke_layers
            row_column_index = col
            origin = row

        stroke_layer = None
        for layer_id in layer_ids:
            if self.objects[layer_id.identifier].row_column_index == row_column_index:
                stroke_layer = self.objects[layer_id.identifier]
        if stroke_layer is not None:
            stroke_patched = False
            for stroke_run in stroke_layer.stroke_runs:
                stroke_start = stroke_run.origin
                stroke_end = stroke_run.origin + stroke_run.length
                stroke_range = range(stroke_start, stroke_end)
                if origin <= stroke_start and (origin + length) >= stroke_end:
                    # New stroke overwrites all of existing stroke
                    stroke_run.CopyFrom(self.create_stroke(origin, length, border_value))
                    stroke_patched = True
                elif origin == stroke_start and length < stroke_run.length:
                    # New stroke writes to start of existing stroke
                    stroke_run.origin = origin + length
                    stroke_run.length = stroke_run.length - length
                elif origin in stroke_range and (origin + length) == stroke_end:
                    # New stoke writes to end of existing stroke
                    stroke_run.length = stroke_run.length - length
                elif origin in stroke_range and (origin + length) in stroke_range:
                    # New stroke in middle of existing stroke
                    stroke_run.length = origin - stroke_start
                    stroke_layer.stroke_runs.append(
                        TSTArchives.StrokeLayerArchive.StrokeRunArchive()
                    )
                    stroke_layer.stroke_runs[-1].CopyFrom(stroke_run)
                    stroke_layer.stroke_runs[-1].origin = origin + length
                    stroke_layer.stroke_runs[-1].length = stroke_end - origin - length
            if not stroke_patched:
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
            raise IndexError(f"{filename}: image already exists in document")
        self.objects.file_store[stored_filename] = data

    def next_image_identifier(self):
        """Return the next available ID in the list of images in the document."""
        datas = self.objects[PACKAGE_ID].datas
        image_ids = [x.identifier for x in datas]
        # datas never appears to be an empty list (default themes include images)
        return max(image_ids) + 1


def rgb(obj) -> RGB:
    """Convert a TSPArchives.Color into an RGB tuple."""
    return RGB(round(obj.r * 255), round(obj.g * 255), round(obj.b * 255))


def range_end(obj):
    """Select end range for a IndexSetArchive.IndexSetEntry."""
    if obj.HasField("range_end"):
        return obj.range_end
    else:
        return obj.range_begin


def formatted_number(number_type, index):
    """Returns the numbered index bullet formatted for different types."""
    bullet_char = BULLET_PREFIXES[number_type]
    bullet_char += BULLET_CONVERSION[number_type](index)
    bullet_char += BULLET_SUFFIXES[number_type]

    return bullet_char


def node_to_col_ref(node: object, table_name: str, col: int) -> str:
    col = node.AST_column.column if node.AST_column.absolute else col + node.AST_column.column

    col_name = xl_col_to_name(col, node.AST_column.absolute)
    if table_name is not None:
        return f"{table_name}::{col_name}"
    else:
        return col_name


def node_to_row_ref(node: object, table_name: str, row: int) -> str:
    row = node.AST_row.row if node.AST_row.absolute else row + node.AST_row.row

    row_name = f"${row+1}" if node.AST_row.absolute else f"{row+1}"
    if table_name is not None:
        return f"{table_name}::{row_name}:{row_name}"
    else:
        return f"{row_name}:{row_name}"


def node_to_row_col_ref(node: object, table_name: str, row: int, col: int) -> str:
    row = node.AST_row.row if node.AST_row.absolute else row + node.AST_row.row
    col = node.AST_column.column if node.AST_column.absolute else col + node.AST_column.column

    ref = xl_rowcol_to_cell(
        row,
        col,
        row_abs=node.AST_row.absolute,
        col_abs=node.AST_column.absolute,
    )
    if table_name is not None:
        return f"{table_name}::{ref}"
    else:
        return ref


def get_storage_buffers_for_row(
    storage_buffer: bytes, offsets: list, num_cols: int, has_wide_offsets: bool
) -> List[bytes]:
    """Extract storage buffers for each cell in a table row.

    Args:
        storage_buffer:  cell_storage_buffer or cell_storage_buffer for a table row
        offsets: 16-bit cell offsets for a table row
        num_cols: number of columns in a table row
        has_wide_offsets: use 4-byte offsets rather than 1-byte offset

    Returns:
         data: list of bytes for each cell in a row, or None if empty
    """
    offsets = array("h", offsets).tolist()
    if has_wide_offsets:
        offsets = [o * 4 for o in offsets]

    data = []
    for col in range(num_cols):
        if col >= len(offsets):
            break

        start = offsets[col]
        if start < 0:
            data.append(None)
            continue

        if col == (len(offsets) - 1):
            end = len(storage_buffer)
        else:
            end = None
            # Find next positive offset
            for i, x in enumerate(offsets[col + 1 :]):
                if x >= 0:
                    end = offsets[col + i + 1]
                    break
            if end is None:
                end = len(storage_buffer)
        data.append(storage_buffer[start:end])

    return data


def clear_field_container(obj):
    """Remove all entries from a protobuf RepeatedCompositeFieldContainer
    in a portable fashion.
    """
    while len(obj) > 0:
        _ = obj.pop()


def field_references(obj: object) -> dict:
    """Return a dict of all fields in an object that are references to other objects."""
    return {
        x[0].name: {"identifier": getattr(obj, x[0].name).identifier}
        for x in obj.ListFields()
        if type(getattr(obj, x[0].name)) == TSPMessages.Reference
    }
