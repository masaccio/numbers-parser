from array import array
from datetime import datetime, timedelta
from functools import lru_cache
from roman import toRoman
from struct import unpack
from typing import Dict, List

from numbers_parser.containers import ObjectStore
from numbers_parser.cell import xl_rowcol_to_cell, xl_col_to_name
from numbers_parser.exceptions import UnsupportedError
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.generated import TSWPArchives_pb2 as TSWPArchives
from numbers_parser.formula import TableFormulas


class CellValue:
    def __init__(self, _type):
        self.type = _type
        self.value = None
        self.rich = None
        self.text = None
        self.ieee = None
        self.date = None
        self.bullets = None


class NumbersModel:
    """
    Loads all objects from Numbers document and provides decoding
    methods for other classes in the module to abstract away the
    internal structures of Numbers document data structures.

    Not to be used in application code.
    """

    def __init__(self, filename):
        self.objects = ObjectStore(filename)
        self._row_headers = {}
        self._row_columns = {}
        self._table_strings = {}
        self._table_tiles = {}
        self._tables_in_sheet = {}
        self._table_base_ids = {}

    def find_refs(self, ref: str) -> list:
        return self.objects.find_refs(ref)

    def sheet_ids(self):
        return [o.identifier for o in self.objects[1].sheets]

    def sheet_name(self, sheet_id):
        return self.objects[sheet_id].name

    @lru_cache(maxsize=None)
    def table_ids(self, sheet_id):
        table_info_ids = self.find_refs("TableInfoArchive")
        return [
            self.objects[t_id].tableModel.identifier
            for t_id in table_info_ids
            if self.objects[t_id].super.parent.identifier == sheet_id
        ]

    @lru_cache(maxsize=None)
    def row_storage_map(self, table_id):
        # The base data store contains a reference to rowHeaders.buckets
        # which is an ordered list that matches the storage buffers, but
        # identifies which row a storage buffer belongs to (empty rows have
        # no storage buffers). Each bucket is:
        #
        #  {
        #      "hidingState": 0,
        #      "index": 0,
        #      "numberOfCells": 3,
        #      "size": 0.0
        #  },
        row_bucket_map = {i: None for i in range(self.objects[table_id].number_of_rows)}
        bds = self.objects[table_id].base_data_store
        buckets = self.objects[bds.rowHeaders.buckets[0].identifier].headers
        for i, bucket in enumerate(buckets):
            row_bucket_map[bucket.index] = i
        return row_bucket_map

    def number_of_rows(self, table_id):
        return self.objects[table_id].number_of_rows

    def number_of_columns(self, table_id):
        return self.objects[table_id].number_of_columns

    def table_name(self, table_id):
        return self.objects[table_id].table_name

    @lru_cache(maxsize=None)
    def table_tiles(self, table_id):
        bds = self.objects[table_id].base_data_store
        return [self.objects[t.tile.identifier] for t in bds.tiles.tiles]

    @lru_cache(maxsize=None)
    def table_string(self, table_id, key):
        bds = self.objects[table_id].base_data_store
        strings_id = bds.stringTable.identifier
        for x in self.objects[strings_id].entries:  # pragma: no branch
            if x.key == key:
                return x.string

    @lru_cache(maxsize=None)
    def owner_id_map(self):
        """ "
        Extracts the mapping table from Owner IDs to UUIDs. Returns a
        dictionary mapping the owner ID int to a 128-bit UUID.
        """
        # The TSCE.CalculationEngineArchive contains a list of mapping entries
        # in dependencyTracker.formulaOwnerDependencies from the root level
        # of the protobuf. Each mapping contains a 32-bit style UUID:
        #
        # "ownerIdMap": {
        #     "mapEntry": [
        #     {
        #         "internalOwnerId": 33,
        #         "ownerId": {
        #             "uuidW0": "0x8750e563, "uuidW1": "0x1e4bfcc0,
        #             "uuidW2": "0xc26dda92, "uuidW3": "0x3cb03f23
        #         }
        #     },
        #
        #
        ce_id = self.find_refs("CalculationEngineArchive")[0]
        calc_engine = self.objects[ce_id]
        owner_id_map = {}
        for e in calc_engine.dependency_tracker.owner_id_map.map_entry:
            owner_id_map[e.internal_owner_id] = uuid(e.owner_id)
        return owner_id_map

    @lru_cache(maxsize=None)
    def table_base_id(self, table_id: int) -> int:
        """ "Finds the UUID of a table"""
        # Look for a TSCE.FormulaOwnerDependenciesArchive objects with the following at the
        # root level of the protobuf:
        #
        #     "baseOwnerUid": {
        #         "lower": "0x0c4ebfb1d9676393",
        #         "upper": "0xf9ad9f35d33aba96"
        #     }
        #      "formulaOwnerUid": {
        #         "lower": "0x0c4ebfb1d96763b6",
        #         "upper": "0xf9ad9f35d33aba96"
        #     },
        #
        # The Table UUID is the TSCE.FormulaOwnerDependenciesArchive whose formulaOwnerUid
        # matches the UUID of the hauntedOwner of the Table:
        #
        #    "hauntedOwner": {
        #        "ownerUid": {
        #            "lower": "0x0c4ebfb1d96763b6",
        #            "upper": "0xf9ad9f35d33aba96"
        #        }
        #    }
        haunted_owner = uuid(self.objects[table_id].haunted_owner.owner_uid)
        formula_owner_ids = self.find_refs("FormulaOwnerDependenciesArchive")
        for dependency_id in formula_owner_ids:  # pragma: no branch
            obj = self.objects[dependency_id]
            if obj.HasField("base_owner_uid") and obj.HasField(
                "formula_owner_uid"
            ):  # pragma: no branch
                base_owner_uid = uuid(obj.base_owner_uid)
                formula_owner_uid = uuid(obj.formula_owner_uid)
                if formula_owner_uid == haunted_owner:
                    return base_owner_uid

    @lru_cache(maxsize=None)
    def formula_cell_ranges(self, table_id: int) -> list:
        """Exract all the formula cell ranges for the Table."""
        # The TSCE.CalculationEngineArchive contains formulaOwnerInfo records
        # inside the dependencyTracker. Each of these has a cellDependencies
        # dictionary and some contain cellRecords that contain formula references:
        #
        # "cellDependencies": {
        #     "cellRecord": [
        #         {
        #             "column": 0,
        #             "containsAFormula": true,
        #             "edges": { "packedEdgeWithoutOwner": [ 16777216 ] },
        #             "row": 1
        #         },
        #         {
        #             "column": 1,
        #             "containsAFormula": true,
        #             "edges": { "packedEdgeWithoutOwner": [16777216 ] },
        #             "row": 1
        #         }
        #     ]
        # }
        cell_records = []
        table_base_id = self.table_base_id(table_id)
        ce_id = self.find_refs("CalculationEngineArchive")[0]
        for finfo in self.objects[ce_id].dependency_tracker.formula_owner_info:
            if finfo.HasField("cell_dependencies"):
                formula_owner_id = uuid(finfo.formula_owner_id)
                if formula_owner_id == table_base_id:
                    for cell_record in finfo.cell_dependencies.cell_record:
                        if cell_record.contains_a_formula:
                            cell_records.append((cell_record.row, cell_record.column))
        return cell_records

    @lru_cache(maxsize=None)
    def merge_cell_ranges(self, table_id):
        """Exract all the merge cell ranges for the Table."""
        # Merge ranges are stored in a number of structures, but the simplest is
        # a TSCE.RangePrecedentsTileArchive which exists for each Table in the
        # Document. These archives contain a fromToRange list which has the merge
        # ranges associated with an Owner ID
        #
        # The Owner IDs need to be extracted from the Calculation Engine using
        # UUID matching (this seems fragile, but no other mechanism has been found)
        #
        #  "fromToRange": [
        #      {
        #          "fromCoord": {
        #          "column": 0,
        #          "row": 0
        #      },
        #      "refersToRect": {
        #          "origin": {
        #              "column": 0,
        #              "row": 0
        #          },
        #          "size": {
        #              "numColumns": 2
        #          }
        #      }
        #  ],
        #  "toOwnerId": 1
        #
        owner_id_map = self.owner_id_map()
        table_base_id = self.table_base_id(table_id)

        merge_cells = {}
        range_table_ids = self.find_refs("RangePrecedentsTileArchive")
        for range_id in range_table_ids:
            o = self.objects[range_id]
            to_owner_id = o.to_owner_id
            if owner_id_map[to_owner_id] == table_base_id:
                for from_to_range in o.from_to_range:
                    rect = from_to_range.refers_to_rect
                    row_start = rect.origin.row
                    row_end = row_start + rect.size.num_rows - 1
                    col_start = rect.origin.column
                    col_end = col_start + rect.size.num_columns - 1
                    for row_num in range(row_start, row_end + 1):
                        for col_num in range(col_start, col_end + 1):
                            merge_cells[(row_num, col_num)] = {
                                "merge_type": "ref",
                                "rect": (row_start, col_start, row_end, col_end),
                                "size": (rect.size.num_rows, rect.size.num_columns),
                            }
                    merge_cells[(row_start, col_start)]["merge_type"] = "source"
        return merge_cells

    @lru_cache(maxsize=None)
    def table_uuids_to_id(self, table_uuid):
        for t_id in self.find_refs("TableInfoArchive"):  # pragma: no branch
            table_model_id = self.objects[t_id].tableModel.identifier
            if table_uuid == self.table_base_id(table_model_id):
                return table_model_id

    def node_to_ref(self, row_num: int, col_num: int, node):
        table_name = None
        if node.HasField("AST_cross_table_reference_extra_info"):
            table_uuid = uuid(node.AST_cross_table_reference_extra_info.table_id)
            table_id = self.table_uuids_to_id(table_uuid)
            table_name = self.table_name(table_id)

        if node.HasField("AST_column") and not node.HasField("AST_row"):
            return node_to_col_ref(node, table_name, col_num)
        else:
            return node_to_row_col_ref(node, table_name, row_num, col_num)

    @lru_cache(maxsize=None)
    def formula_ast(self, table_id: int):
        bds = self.objects[table_id].base_data_store
        formula_table_id = bds.formula_table.identifier
        formula_table = self.objects[formula_table_id]
        formulas = {}
        for formula in formula_table.entries:
            formulas[formula.key] = formula.formula.AST_node_array.AST_node
        return formulas

    @lru_cache(maxsize=None)
    def storage_buffers_pre_bnc(self, table_id: int) -> List:
        row_infos = []
        for tile in self.table_tiles(table_id):
            row_infos += tile.rowInfos

        return [
            get_storage_buffers_for_row(
                r.cell_storage_buffer_pre_bnc,
                r.cell_offsets_pre_bnc,
                self.number_of_columns(table_id),
                r.has_wide_offsets,
            )
            for r in row_infos
        ]

    @lru_cache(maxsize=None)
    def storage_buffers(self, table_id: int) -> List:
        row_infos = []
        for tile in self.table_tiles(table_id):
            row_infos += tile.rowInfos

        return [
            get_storage_buffers_for_row(
                r.cell_storage_buffer,
                r.cell_offsets,
                self.number_of_columns(table_id),
                r.has_wide_offsets,
            )
            for r in row_infos
        ]

    @lru_cache(maxsize=None)
    def storage_buffer(self, table_id: int, row_num: int, col_num: int) -> bytes:
        row_offset = self.row_storage_map(table_id)[row_num]
        if row_offset is None:
            return None
        try:
            storage_buffers = self.storage_buffers(table_id)
            return storage_buffers[row_offset][col_num]
        except IndexError:
            return None

    @lru_cache(maxsize=None)
    def storage_buffer_pre_bnc(
        self, table_id: int, row_num: int, col_num: int
    ) -> bytes:
        row_offset = self.row_storage_map(table_id)[row_num]
        if row_offset is None:
            return None
        try:
            storage_buffers_pre_bnc = self.storage_buffers_pre_bnc(table_id)
            return storage_buffers_pre_bnc[row_offset][col_num]
        except IndexError:
            return None

    @lru_cache(maxsize=None)
    def table_formulas(self, table_id: int):
        return TableFormulas(self, table_id)

    @lru_cache(maxsize=None)
    def table_cell_decode(self, table_id: int, row_num: int, col_num: int) -> Dict:
        storage_buffer = self.storage_buffer(table_id, row_num, col_num)
        storage_buffer_pre_bnc = self.storage_buffer_pre_bnc(table_id, row_num, col_num)

        if storage_buffer is None:
            return None

        cell_type = storage_buffer[1]
        if storage_buffer_pre_bnc is not None:
            cell_value = storage_buffer_value(storage_buffer_pre_bnc, cell_type)
        else:
            cell_value = CellValue(cell_type)

        if cell_value.type == TSTArchives.numberCellType or cell_value.type == 10:
            if storage_buffer_pre_bnc is None:
                cell_value.value = 0.0
            else:
                cell_value.value = cell_value.ieee
            cell_value.type = TSTArchives.numberCellType
        elif cell_value.type == TSTArchives.textCellType:
            if cell_value.text is None:
                cell_value.text = unpack("<i", storage_buffer[12:16])[0]
            cell_value.value = self.table_string(table_id, cell_value.text)
        elif cell_value.type == TSTArchives.dateCellType:
            if storage_buffer_pre_bnc is None:
                cell_value.value = datetime(2001, 1, 1)
            else:
                cell_value.value = datetime(2001, 1, 1) + timedelta(
                    seconds=cell_value.date
                )
        elif cell_value.type == TSTArchives.boolCellType:
            cell_value.value = cell_value.ieee > 0.0
        elif cell_value.type == TSTArchives.durationCellType:
            cell_value.value = cell_value.ieee
        elif cell_value.type == TSTArchives.automaticCellType:
            cell_value.bullets = self.table_bullets(table_id, cell_value.rich)

        return cell_value

    @lru_cache(maxsize=None)
    def table_bullets(self, table_id: int, string_key: int) -> Dict:
        """
        Extract bullets from a rich text data cell.
        Returns None if the cell is not rich text
        """
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
        for entry in rich_text_table.entries:
            if string_key == entry.key:
                payload = self.objects[entry.rich_text_payload.identifier]
                payload_storage = self.objects[payload.storage.identifier]
                payload_entries = payload_storage.table_para_style.entries
                table_list_styles = payload_storage.table_list_style.entries
                offsets = [e.character_index for e in payload_entries]

                cell_text = payload_storage.text[0]
                bullets = []
                bullet_chars = []
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
                        bullet_char = ""

                    bullet_chars.append(bullet_char)

                return {
                    "text": cell_text,
                    "bullets": bullets,
                    "bullet_chars": bullet_chars,
                }
        return None

    @lru_cache(maxsize=None)
    def table_cell_formula_decode(
        self, table_id: int, row_num: int, col_num: int, cell_type: int
    ):
        storage_buffer = self.storage_buffer(table_id, row_num, col_num)
        if not self.table_formulas(table_id).is_formula(row_num, col_num):
            return None

        # TODO: Review this decoding mod versus storage_buffer_value()
        mod_offset = (
            unpack("<h", storage_buffer[6:8])[0] == 1
            or unpack("<h", storage_buffer[6:8])[0] == 8
        )

        if cell_type == TSTArchives.formulaErrorCellType:
            formula_key = unpack("<h", storage_buffer[-4:-2])[0]
        elif (
            cell_type == TSTArchives.textCellType
            or cell_type == TSTArchives.durationCellType
        ) and mod_offset:
            formula_key = unpack("<h", storage_buffer[-16:-14])[0]
        else:
            formula_key = unpack("<h", storage_buffer[-12:-10])[0]

        return formula_key


def formatted_number(number_type, index):
    """Returns the numbered index bullet formatted for different types"""
    if number_type == TSWPArchives.ListStyleArchive.kNumericDecimal:
        bullet_char = str(index + 1) + "."
    elif number_type == TSWPArchives.ListStyleArchive.kNumericDoubleParen:
        bullet_char = "(" + str(index + 1) + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kNumericRightParen:
        bullet_char = str(index + 1) + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kRomanUpperDecimal:
        bullet_char = toRoman(index + 1) + "."
    elif number_type == TSWPArchives.ListStyleArchive.kRomanUpperDoubleParen:
        bullet_char = "(" + toRoman(index + 1) + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kRomanUpperRightParen:
        bullet_char = toRoman(index + 1) + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kRomanLowerDecimal:
        bullet_char = toRoman(index + 1).lower() + "."
    elif number_type == TSWPArchives.ListStyleArchive.kRomanLowerDoubleParen:
        bullet_char = "(" + toRoman(index + 1).lower() + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kRomanLowerRightParen:
        bullet_char = toRoman(index + 1).lower() + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kAlphaUpperDecimal:
        bullet_char = chr(index + 65) + "."
    elif number_type == TSWPArchives.ListStyleArchive.kAlphaUpperDoubleParen:
        bullet_char = "(" + chr(index + 65) + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kAlphaUpperRightParen:
        bullet_char = chr(index + 65) + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kAlphaLowerDecimal:
        bullet_char = chr(index + 97) + "."
    elif number_type == TSWPArchives.ListStyleArchive.kAlphaLowerDoubleParen:
        bullet_char = "(" + chr(index + 97) + ")"
    elif number_type == TSWPArchives.ListStyleArchive.kAlphaLowerRightParen:
        bullet_char = chr(index + 97) + ")"
    else:
        bullet_char = ""
    return bullet_char


def node_to_col_ref(node: object, table_name: str, col_num: int) -> str:
    if node.AST_column.absolute:
        col = node.AST_column.column
    else:
        col = col_num + node.AST_column.column

    col_name = xl_col_to_name(col, node.AST_column.absolute)
    if table_name is not None:
        return f"{table_name}::{col_name}"
    else:
        return col_name


def node_to_row_col_ref(
    node: object, table_name: str, row_num: int, col_num: int
) -> str:
    if node.AST_row.absolute:
        row = node.AST_row.row
    else:
        row = row_num + node.AST_row.row
    if node.AST_column.absolute:
        col = node.AST_column.column
    else:
        col = col_num + node.AST_column.column

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


def uuid(ref: dict) -> int:
    """
    Extract storage buffers for each cell in a table row

    Args:
        ref: Google protobuf containing either four 32-bit IDs or two 64-bit IDs

    Returns:
        uuid: 128-bit UUID

    Raises:
        UnsupportedError: object does not include expected UUID fields
    """
    try:
        if hasattr(ref, "upper"):
            uuid = ref.upper << 64 | ref.lower
        else:
            uuid = (
                (ref.uuid_w3 << 96)
                | (ref.uuid_w2 << 64)
                | (ref.uuid_w1 << 32)
                | ref.uuid_w0
            )
    except AttributeError:
        raise UnsupportedError(f"Unsupported UUID structure: {ref}")  # pragma: no cover

    return uuid


def storage_buffer_value(buffer, cell_type):
    """Decode data values from a storage buffer based on the type of data represented"""
    cell_value = CellValue(cell_type)
    flags = unpack("<i", buffer[4:8])[0]
    data_offset = 12 + bin(flags & 0x0D8E).count("1") * 4
    if (flags & 0x200) > 0:
        cell_value.rich = unpack("<i", buffer[data_offset : data_offset + 4])[0]
        data_offset = data_offset + 4
    data_offset = data_offset + bin(flags & 0x3000).count("1") * 4
    if (flags & 0x010) > 0:
        cell_value.text = unpack("<i", buffer[data_offset : data_offset + 4])[0]
        data_offset = data_offset + 4
    if (flags & 0x020) > 0:
        cell_value.ieee = unpack("<d", buffer[data_offset : data_offset + 8])[0]
        data_offset = data_offset + 8
    if (flags & 0x040) > 0:
        cell_value.date = unpack("<d", buffer[data_offset : data_offset + 8])[0]
        data_offset = data_offset + 8
    return cell_value


def get_storage_buffers_for_row(
    storage_buffer: bytes, offsets: list, num_cols: int, has_wide_offsets: bool
) -> List[bytes]:
    """
    Extract storage buffers for each cell in a table row

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
    for col_num in range(num_cols):
        if col_num >= len(offsets):
            break

        start = offsets[col_num]
        if start < 0:
            data.append(None)
            continue

        if col_num == (len(offsets) - 1):
            end = len(storage_buffer)
        else:
            # Get next offset past current one that is not -1
            # https://stackoverflow.com/questions/19502378/
            idx = next(
                (i for i, x in enumerate(offsets[col_num + 1 :]) if x >= 0), None
            )
            if idx is None:
                end = len(storage_buffer)
            else:
                end = offsets[col_num + idx + 1]
        data.append(storage_buffer[start:end])

    return data
