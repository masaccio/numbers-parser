import array
import struct

from typing import Union, List, Generator
from datetime import datetime, timedelta

from numbers_parser.containers import ItemsList, ObjectStore, NumbersError
from numbers_parser.cell import (
    BoolCell,
    Cell,
    DateCell,
    DurationCell,
    EmptyCell,
    ErrorCell,
    FormulaCell,
    MergedCell,
    NumberCell,
    TextCell,
    xl_cell_to_rowcol,
    xl_range,
)


from numbers_parser.generated import TSTArchives_pb2 as TSTArchives


class UnsupportedError(NumbersError):
    """Raised for unsupported file format features"""

    pass


class Document:
    def __init__(self, filename):
        self._object_store = ObjectStore(filename)

    def sheets(self):
        refs = [o.identifier for o in self._object_store[1].sheets]
        self._sheets = Sheets(self._object_store, refs)
        return self._sheets


class Sheets(ItemsList):
    def __init__(self, object_store, refs):
        super(Sheets, self).__init__(object_store, refs, Sheet)


class Sheet:
    def __init__(self, object_store, sheet_id):
        self._sheet_id = sheet_id
        self._object_store = object_store
        self._sheet = object_store[sheet_id]
        self.name = self._sheet.name

    def tables(self):
        table_ids = self._object_store.find_refs("TableInfoArchive")
        table_refs = [
            self._object_store[table_id].tableModel.identifier
            for table_id in table_ids
            if self._object_store[table_id].super.parent.identifier == self._sheet_id
        ]
        self._tables = Tables(self._object_store, table_refs)

        return self._tables


class Tables(ItemsList):
    def __init__(self, object_store, refs):
        super(Tables, self).__init__(object_store, refs, Table)


class Table:
    def __init__(self, object_store, table_id):
        super(Table, self).__init__()
        self._object_store = object_store
        self._table = object_store[table_id]
        self._table_id = table_id
        self.name = self._table.table_name

        bds = self._table.base_data_store
        self._row_headers = [
            h.numberOfCells for h in object_store[bds.rowHeaders.buckets[0].identifier].headers
        ]
        self._column_headers = [
            h.numberOfCells for h in object_store[bds.columnHeaders.identifier].headers
        ]
        self._tile_ids = [t.tile.identifier for t in bds.tiles.tiles]
        self._table_data = None
        self._table_cells = None
        self._merge_cells = None

    @property
    def data(self):
        if self._table_data is None:
            table_data = []
            for row in self.cells:
                table_data.append([cell.value for cell in row])
            self._table_data = table_data

        return self._table_data

    @property
    def cells(self) -> list:
        if self._table_cells is None:
            self._table_cells = self._extract_table_cells()
        return self._table_cells

    @property
    def merge_ranges(self) -> list:
        if self._merge_cells is None:
            self._merge_cells = _extract_merge_cells(self._object_store, self._table)

        ranges = []
        for row_num, cols in self._merge_cells.items():
            for col_num in cols:
                ranges.append(xl_range(*self._merge_cells[row_num][col_num]["rect"]))
        return sorted(set(list(ranges)))

    def _extract_table_cells(self):
        if self._merge_cells is None:
            self._merge_cells = _extract_merge_cells(self._object_store, self._table)

        row_infos = []
        for tile_id in self._tile_ids:
            row_infos += self._object_store[tile_id].rowInfos

        storage_version = self._object_store[tile_id].storage_version
        if storage_version != 5:
            raise UnsupportedError(  # pragma: no cover
                f"Unsupported row info version {storage_version}"
            )

        storage_buffers = [
            _extract_cell_data(
                r.cell_storage_buffer, r.cell_offsets, self.num_cols, r.has_wide_offsets
            )
            for r in row_infos
        ]
        storage_buffers_pre_bnc = [
            _extract_cell_data(
                r.cell_storage_buffer_pre_bnc,
                r.cell_offsets_pre_bnc,
                self.num_cols,
                r.has_wide_offsets,
            )
            for r in row_infos
        ]

        table_cells = []
        num_rows = sum([self._object_store[t].numrows for t in self._tile_ids])
        for row_num in range(num_rows):
            row = []
            for col_num in range(self.num_cols):
                if col_num < len(storage_buffers[row_num]):
                    storage_buffer = storage_buffers[row_num][col_num]
                else:
                    # Rest of row is empty cells
                    row.append(EmptyCell(row_num, col_num))
                    continue

                if col_num < len(storage_buffers_pre_bnc[row_num]):
                    storage_buffer_pre_bnc = storage_buffers_pre_bnc[row_num][col_num]
                else:
                    storage_buffer_pre_bnc = None

                cell = None
                if storage_buffer is None:
                    try:
                        cell = MergedCell(*self._merge_cells[row_num][col_num]["rect"])
                    except KeyError:
                        cell = EmptyCell(row_num, col_num)
                    cell.size = None
                    row.append(cell)
                else:
                    cell_type = storage_buffer[1]
                    if cell_type == TSTArchives.emptyCellValueType:
                        cell = EmptyCell(row_num, col_num)
                    elif cell_type == TSTArchives.numberCellType:
                        if storage_buffer_pre_bnc is None:
                            cell_value = 0.0
                        else:
                            cell_value = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                        cell = NumberCell(row_num, col_num, cell_value)
                    elif cell_type == TSTArchives.textCellType:
                        key = struct.unpack("<i", storage_buffer[12:16])[0]
                        cell = TextCell(row_num, col_num, self._table_string(key))
                    elif cell_type == TSTArchives.dateCellType:
                        if storage_buffer_pre_bnc is None:
                            cell_value = datetime(2001, 1, 1)
                        else:
                            seconds = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                            cell_value = datetime(2001, 1, 1) + timedelta(seconds=seconds)
                        cell = DateCell(row_num, col_num, cell_value)
                    elif cell_type == TSTArchives.boolCellType:
                        d = struct.unpack("<d", storage_buffer[12:20])[0]
                        cell = BoolCell(row_num, col_num, d > 0.0)
                    elif cell_type == TSTArchives.durationCellType:
                        cell_value = struct.unpack("<d", storage_buffer[12:20])[0]
                        cell = DurationCell(row_num, col_num, timedelta(days=cell_value))
                    elif cell_type == TSTArchives.formulaErrorCellType:
                        cell = ErrorCell(row_num, col_num)
                    elif cell_type == 9:
                        cell = FormulaCell(row_num, col_num)
                    elif cell_type == 10:
                        # TODO: Is this realy a number cell (experiments suggest so)
                        if storage_buffer_pre_bnc is None:
                            cell_value = 0.0
                        else:
                            cell_value = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                        cell = NumberCell(row_num, col_num, cell_value)
                    else:
                        raise UnsupportedError(  # pragma: no cover
                            f"Unsupport cell type {cell_type} @{self.name}:({row_num},{col_num})"
                        )

                    try:
                        if self._merge_cells[row_num][col_num]["merge_type"] == "source":
                            cell.is_merged = True
                            cell.size = self._merge_cells[row_num][col_num]["size"]
                    except KeyError:
                        cell.is_merged = False
                        cell.size = (1, 1)
                    row.append(cell)
            table_cells.append(row)

        return table_cells

    def _table_string(self, key):
        if not hasattr(self, "_table_strings"):
            strings_id = self._table.base_data_store.stringTable.identifier
            self._table_strings = {x.key: x.string for x in self._object_store[strings_id].entries}
        return self._table_strings[key]

    @property
    def num_rows(self) -> int:
        """Number of rows in the table"""
        return len(self._row_headers)

    @property
    def num_cols(self) -> int:
        """Number of columns in the table"""
        return len(self._column_headers)

    def cell(self, *args) -> Union[Cell, MergedCell]:
        if type(args[0]) == str:
            (row_num, col_num) = xl_cell_to_rowcol(args[0])
        elif len(args) != 2:
            raise IndexError("invalid cell reference " + str(args))
        else:
            (row_num, col_num) = args

        cells = self.cells
        if row_num >= len(cells):
            raise IndexError(f"row {row_num} out of range")
        if col_num >= len(cells[row_num]):
            raise IndexError(f"coumn {col_num} out of range")
        return cells[row_num][col_num]

    def iter_rows(
        self,
        min_row: int = None,
        max_row: int = None,
        min_col: int = None,
        max_col: int = None,
        values_only: bool = False,
    ) -> Generator[tuple, None, None]:
        """
        Produces cells from a table, by row. Specify the iteration range using
        the indexes of the rows and columns.
        Args:
            min_row: smallest row index (zero indexed)
            max_row: largest row (zero indexed)
            min_col: smallest row index (zero indexed)
            max_col: largest row (zero indexed)
            values_only: return cell values rather than Cell objects
        Returns:
            generator: tuple of cells
        Raises:
            IndexError: row or column values are out of range for the table
        """
        min_row = min_row or 0
        max_row = max_row or self.num_rows - 1
        min_col = min_col or 0
        max_col = max_col or self.num_cols - 1

        if min_row < 0:
            raise IndexError(f"row {min_row} out of range")
        if max_row > self.num_rows:
            raise IndexError(f"row {max_row} out of range")
        if min_col < 0:
            raise IndexError(f"column {min_col} out of range")
        if max_col > self.num_cols:
            raise IndexError(f"column {max_col} out of range")

        cells = self.cells
        for row_num in range(min_row, max_row + 1):
            if values_only:
                yield tuple(cell.value or None for cell in cells[row_num][min_col : max_col + 1])
            else:
                yield tuple(cells[row_num][min_col : max_col + 1])

    def iter_cols(
        self,
        min_col: int = None,
        max_col: int = None,
        min_row: int = None,
        max_row: int = None,
        values_only: bool = False,
    ) -> Generator[tuple, None, None]:
        """
        Produces cells from a table, by column. Specify the iteration range using
        the indexes of the rows and columns.
        Args:
            min_col: smallest row index (zero indexed)
            max_col: largest row (zero indexed)
            min_row: smallest row index (zero indexed)
            max_row: largest row (zero indexed)
            values_only: return cell values rather than Cell objects
        Returns:
            generator: tuple of cells
        Raises:
            IndexError: row or column values are out of range for the table
        """
        min_row = min_row or 0
        max_row = max_row or self.num_rows - 1
        min_col = min_col or 0
        max_col = max_col or self.num_cols - 1

        if min_row < 0:
            raise IndexError(f"row {min_row} out of range")
        if max_row > self.num_rows:
            raise IndexError(f"row {max_row} out of range")
        if min_col < 0:
            raise IndexError(f"column {min_col} out of range")
        if max_col > self.num_cols:
            raise IndexError(f"column {max_col} out of range")

        cells = self.cells
        for col_num in range(min_col, max_col + 1):
            if values_only:
                yield tuple(row[col_num].value for row in cells[min_row : max_row + 1])
            else:
                yield tuple(row[col_num] for row in cells[min_row : max_row + 1])


def _extract_cell_data(
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
    offsets = array.array("h", offsets).tolist()
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
            idx = next((i for i, x in enumerate(offsets[col_num + 1 :]) if x >= 0), None)
            if idx is None:
                end = len(storage_buffer)
            else:
                end = offsets[col_num + idx + 1]
        data.append(storage_buffer[start:end])

    return data


def _uuid(ref: dict) -> int:
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
        uuid = ref.upper << 64 | ref.lower
    except AttributeError:
        try:
            uuid = (ref.uuid_w3 << 96) | (ref.uuid_w2 << 64) | (ref.uuid_w1 << 32) | ref.uuid_w0
        except AttributeError:
            raise UnsupportedError(f"Unsupported UUID structure: {ref}")  # pragma: no cover

    return uuid


def _extract_merge_cells(object_store: ObjectStore, table: Table):
    """Exract all the merge cell ranges for the the Table."""
    # Merge ranges are stored in a number of structures, but the simplest is
    # a TSCE.RangePrecedentsTileArchive which exists for each Table in the
    # Document. These archives contain a fromToRange list which has the merge
    # ranges associated with an Owner ID
    #
    # The Owner IDs need to be extracted from teh Calculation Engine using
    # UUID matching (this seems fragile, but no other mechanism has been found)
    #
    # "fromToRange": [
    #     {
    #         "fromCoord": {
    #         "column": 0,
    #         "row": 0
    #     },
    #     "refersToRect": {
    #         "origin": {
    #             "column": 0,
    #             "row": 0
    #         },
    #         "size": {
    #             "numColumns": 2
    #         }
    #     }
    # ],
    # "toOwnerId": 1
    #
    owner_id_map = _extract_owner_id_map(object_store)
    table_base_id = _extract_table_base_id(object_store, table)

    merge_cells = {}
    range_table_ids = object_store.find_refs("RangePrecedentsTileArchive")
    for range_id in range_table_ids:
        o = object_store[range_id]
        to_owner_id = o.to_owner_id
        if owner_id_map[to_owner_id] == table_base_id:
            for from_to_range in o.from_to_range:
                rect = from_to_range.refers_to_rect
                row_start = rect.origin.row
                row_end = row_start + rect.size.num_rows - 1
                col_start = rect.origin.column
                col_end = col_start + rect.size.num_columns - 1
                for row_num in range(row_start, row_end + 1):
                    if row_num not in merge_cells:
                        merge_cells[row_num] = {}
                    for col_num in range(col_start, col_end + 1):
                        merge_cells[row_num][col_num] = {
                            "merge_type": "ref",
                            "rect": (row_start, col_start, row_end, col_end),
                            "size": (rect.size.num_rows, rect.size.num_columns),
                        }
                merge_cells[row_start][col_start]["merge_type"] = "source"
    return merge_cells


def _extract_owner_id_map(object_store: ObjectStore):
    """"
    Extracts the mapping table from Owner IDs to UUIDs. Returns a
    dictionary mapping the owner ID int to a 128-bit UUID.
    """
    # The TSCE.CalculationEngineArchive contains a list of mapping entries
    # in odependencyTracker.formulaOwnerDependencies from the root level
    # of the protobuf. Each mapping contains a 32-bit style UUID:
    #
    # "ownerIdMap": {
    #     "mapEntry": [
    #     {
    #         "internalOwnerId": 33,
    #         "ownerId": {
    #             "uuidW0": 8750e563, "uuidW1": 1e4bfcc0,
    #             "uuidW2": c26dda92, "uuidW3": 3cb03f23
    #         }
    #     },
    #
    #
    calculation_engine_id = object_store.find_refs("CalculationEngineArchive")[0]
    owner_id_map = {}
    for e in object_store[calculation_engine_id].dependency_tracker.owner_id_map.map_entry:
        owner_id_map[e.internal_owner_id] = _uuid(e.owner_id)
    return owner_id_map


def _extract_table_base_id(object_store: ObjectStore, table: Table):
    """"Finds the UUID of a table"""
    # Look for a TSCE.FormulaOwnerDependenciesArchive objects with the following at the
    # root level of the protobuf:
    #
    #     "baseOwnerUid": {
    #         "lower": 0c4ebfb1d9676393",
    #         "upper": f9ad9f35d33aba96"
    #     }
    #     "formulaOwnerUid": {
    #         "lower": 0c4ebfb1d96763b6",
    #         "upper": f9ad9f35d33aba96"
    #     },
    #
    # The Table UUID is the TSCE.FormulaOwnerDependenciesArchive whose formulaOwnerUid
    # matches the UUID of the hauntedOwner of the Table:
    #
    #    "hauntedOwner": {
    #        "ownerUid": {
    #            "lower": 0c4ebfb1d96763b6",
    #            "upper": f9ad9f35d33aba96"
    #        }
    #    }
    haunted_owner = _uuid(table.haunted_owner.owner_uid)
    for dependency_id in object_store.find_refs("FormulaOwnerDependenciesArchive"):
        obj = object_store[dependency_id]
        if obj.HasField("base_owner_uid") and obj.HasField("formula_owner_uid"):
            base_owner_uid = _uuid(obj.base_owner_uid)
            formula_owner_uid = _uuid(obj.formula_owner_uid)
            if formula_owner_uid == haunted_owner:
                return base_owner_uid
