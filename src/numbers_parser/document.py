import array
import struct
import binascii

from typing import Union, List, Generator
from datetime import datetime, timedelta

from numbers_parser.containers import ItemsList, ObjectStore, NumbersError
from numbers_parser.formula import TableFormulas, get_merge_cell_ranges

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
        self.name = self._table.table_name

        bds = self._table.base_data_store
        self._row_headers = [
            h.numberOfCells for h in object_store[bds.rowHeaders.buckets[0].identifier].headers
        ]
        self._column_headers = [
            h.numberOfCells for h in object_store[bds.columnHeaders.identifier].headers
        ]
        self._tiles = [object_store[t.tile.identifier] for t in bds.tiles.tiles]
        self._table_data = None
        self._table_cells = None
        self._merge_cells = None

    @property
    def data(self) -> list:
        """The data property is deprecated: use rows(values_only=True) instead"""
        return self.rows(values_only=True)

    @property
    def cells(self) -> list:
        """The cells property is deprecated: use rows(values_only=False) instead"""
        return self.rows(values_only=False)

    def rows(self, values_only: bool = False) -> list:
        """
        Return all rows of cells for the Table.

        Args:
            values_only: if True, return cell values instead of Cell objects

        Returns:
            rows: list of rows; each row is a list of Cell objects
        """
        if self._table_cells is None:
            self._table_cells = self._extract_table_cells()
        if values_only:
            row_values = []
            for row_cells in self._table_cells:
                row_values.append([cell.value for cell in row_cells])
            return row_values
        else:
            return self._table_cells

    @property
    def merge_ranges(self) -> list:
        if self._merge_cells is None:
            self._merge_cells = get_merge_cell_ranges(self._object_store, self._table)

        ranges = []
        for row_num, cols in self._merge_cells.items():
            for col_num in cols:
                ranges.append(xl_range(*self._merge_cells[row_num][col_num]["rect"]))
        return sorted(set(list(ranges)))

    def _extract_table_cells(self):
        if self._merge_cells is None:
            self._merge_cells = get_merge_cell_ranges(self._object_store, self._table)

        row_infos = []
        for tile in self._tiles:
            row_infos += tile.rowInfos

        storage_version = self._tiles[0].storage_version

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

        table_formulas = TableFormulas(self._object_store, self._table)

        table_cells = []
        num_rows = sum([t.numrows for t in self._tiles])
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
                        string_key = struct.unpack("<i", storage_buffer[12:16])[0]
                        cell = TextCell(row_num, col_num, self._table_string(string_key))
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
                    elif cell_type == TSTArchives.currencyCellValueType:
                        cell_value = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
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

                    if table_formulas.is_formula(row_num, col_num):
                        try:
                            if table_formulas.is_error(row_num, col_num):
                                formula_key = struct.unpack("<h", storage_buffer[12:14])[0]
                            if cell_type == TSTArchives.numberCellType:  # 2
                                formula_key = struct.unpack("<h", storage_buffer[28:30])[0]
                            elif cell_type == TSTArchives.textCellType:  # 3
                                if len(storage_buffer) > 28:
                                    formula_key = struct.unpack("<h", storage_buffer[20:22])[0]
                                else:
                                    formula_key = struct.unpack("<h", storage_buffer[24:26])[0]
                            elif cell_type == TSTArchives.dateCellType:  # 5
                                formula_key = struct.unpack("<h", storage_buffer[24:26])[0]
                            elif cell_type == TSTArchives.boolCellType:  # 6
                                formula_key = struct.unpack("<h", storage_buffer[24:26])[0]
                            elif cell_type == TSTArchives.durationCellType:  # 7
                                formula_key = struct.unpack("<h", storage_buffer[24:26])[0]
                            elif cell_type == TSTArchives.formulaErrorCellType:  # 8
                                formula_key = struct.unpack("<h", storage_buffer[12:14])[0]
                            elif cell_type == TSTArchives.currencyCellValueType:  # 9
                                formula_key = struct.unpack("<h", storage_buffer[32:34])[0]
                            elif cell_type == 10:
                                # Might be [36:38]
                                formula_key = struct.unpack("<h", storage_buffer[32:34])[0]
                            else:
                                raise UnsupportedError(  # pragma: no cover
                                    f"Unsupported formula cell type '{cell_type}' @{self.name}:({row_num},{col_num})"
                                )

                            # print(f"@{row_num},{col_num}: type={cell_type}, key={formula_key}")
                            # print("\tstorage_buffer =", binascii.hexlify(storage_buffer, sep=":"))
                            ast = table_formulas.ast(formula_key)
                            ast["row"] = row_num
                            ast["column"] = col_num
                            cell.add_formula(ast)
                        except KeyError:
                            raise UnsupportedError(  # pragma: no cover
                                f"Formula not found @{self.name}:({row_num},{col_num})"
                            )
                        except struct.error:
                            raise UnsupportedError(  # pragma: no cover
                                f"Unsupported formula ref @{self.name}:({row_num},{col_num})"
                            )

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
