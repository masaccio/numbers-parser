import array
import struct
import logging
import warnings

from typing import Union, List, Generator
from datetime import datetime, timedelta

from numbers_parser.containers import ItemsList, ObjectStore
from numbers_parser.exceptions import UnsupportedError
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

        bds = self._table.base_data_store
        self._row_headers = [
            h.numberOfCells
            for h in object_store[bds.rowHeaders.buckets[0].identifier].headers
        ]
        self._column_headers = [
            h.numberOfCells for h in object_store[bds.columnHeaders.identifier].headers
        ]
        self._tiles = [object_store[t.tile.identifier] for t in bds.tiles.tiles]

    @property
    def name(self):
        return self._table.table_name

    @property
    def data(self) -> list:
        """The data property will be deprecated; use rows instead"""
        warnings.warn(
            "data property will be deprecated in 3.0; use rows(values_only=True) instead",
            PendingDeprecationWarning,
        )
        return self.rows(values_only=True)

    @property
    def cells(self) -> list:
        """The cells property will be deprecated: use rows instead"""
        warnings.warn(
            "cells property will be deprecated in 3.0; use rows instead",
            PendingDeprecationWarning,
        )
        return self.rows(values_only=False)

    def rows(self, values_only: bool = False) -> list:
        """
        Return all rows of cells for the Table.

        Args:
            values_only: if True, return cell values instead of Cell objects

        Returns:
            rows: list of rows; each row is a list of Cell objects
        """
        if values_only:
            row_values = []
            for row_cells in self._table_cells:
                row_values.append([cell.value for cell in row_cells])
            return row_values
        else:
            return self._table_cells

    @property
    def merge_ranges(self) -> list:
        if not hasattr(self, "_merge_cells"):
            self._merge_cells = get_merge_cell_ranges(self._object_store, self._table)

        ranges = [xl_range(*r["rect"]) for r in self._merge_cells.values()]
        return sorted(set(list(ranges)))

    def _storage_buffer(self, row_num: int, col_num: int, bnc: bool = False) -> bytes:
        if not hasattr(self, "_storage_buffers") or not hasattr(
            self, "_storage_buffers_pre_bnc"
        ):
            row_infos = []
            for tile in self._tiles:
                row_infos += tile.rowInfos

            storage_version = self._tiles[0].storage_version
            if storage_version != 5:
                raise UnsupportedError(  # pragma: no cover
                    f"Unsupported row info version {storage_version}"
                )

            if bnc:
                self._storage_buffers_pre_bnc = [
                    get_storage_buffers_for_row(
                        r.cell_storage_buffer_pre_bnc,
                        r.cell_offsets_pre_bnc,
                        self.num_cols,
                        r.has_wide_offsets,
                    )
                    for r in row_infos
                ]
            else:
                self._storage_buffers = [
                    get_storage_buffers_for_row(
                        r.cell_storage_buffer,
                        r.cell_offsets,
                        self.num_cols,
                        r.has_wide_offsets,
                    )
                    for r in row_infos
                ]

        try:
            if bnc:
                buffer = self._storage_buffers_pre_bnc[row_num][col_num]
            else:
                buffer = self._storage_buffers[row_num][col_num]
        except IndexError:
            return None
        return buffer

    @property
    def _merge_cells(self):
        if not hasattr(self, "__merge_cells"):
            self.__merge_cells = get_merge_cell_ranges(self._object_store, self._table)
        return self.__merge_cells

    @property
    def _table_cells(self):
        if hasattr(self, "__table_cells"):
            return self.__table_cells

        table_formulas = TableFormulas(self._object_store, self._table)

        self.__table_cells = []
        num_rows = sum([t.numrows for t in self._tiles])
        for row_num in range(num_rows):
            row = []
            for col_num in range(self.num_cols):
                storage_buffer = self._storage_buffer(row_num, col_num)
                storage_buffer_pre_bnc = self._storage_buffer(
                    row_num, col_num, bnc=True
                )

                cell = None
                if storage_buffer is None:
                    row_col = (row_num, col_num)
                    if (
                        row_col in self._merge_cells
                        and self._merge_cells[row_col]["merge_type"] == "ref"
                    ):
                        cell = MergedCell(*self._merge_cells[row_col]["rect"])
                    else:
                        cell = EmptyCell(row_num, col_num)
                    cell.size = None
                    row.append(cell)
                else:
                    cell_type = storage_buffer[1]

                    logging.debug(
                        "%s@[%d,%d]: cell_type=%d, storage_buffer=%s",
                        self.name,
                        row_num,
                        col_num,
                        cell_type,
                        ":".join(["{0:02x}".format(b) for b in storage_buffer]),
                    )

                    if cell_type == TSTArchives.emptyCellValueType:
                        cell = EmptyCell(row_num, col_num)
                    elif cell_type == TSTArchives.numberCellType:
                        if storage_buffer_pre_bnc is None:
                            cell_value = 0.0
                        else:
                            cell_value = struct.unpack(
                                "<d", storage_buffer_pre_bnc[-12:-4]
                            )[0]
                        cell = NumberCell(row_num, col_num, cell_value)
                    elif cell_type == TSTArchives.textCellType:
                        string_key = struct.unpack("<i", storage_buffer[12:16])[0]
                        cell = TextCell(
                            row_num, col_num, self._table_strings[string_key]
                        )
                    elif cell_type == TSTArchives.dateCellType:
                        if storage_buffer_pre_bnc is None:
                            cell_value = datetime(2001, 1, 1)
                        else:
                            seconds = struct.unpack(
                                "<d", storage_buffer_pre_bnc[-12:-4]
                            )[0]
                            cell_value = datetime(2001, 1, 1) + timedelta(
                                seconds=seconds
                            )
                        cell = DateCell(row_num, col_num, cell_value)
                    elif cell_type == TSTArchives.boolCellType:
                        d = struct.unpack("<d", storage_buffer[12:20])[0]
                        cell = BoolCell(row_num, col_num, d > 0.0)
                    elif cell_type == TSTArchives.durationCellType:
                        cell_value = struct.unpack("<d", storage_buffer[12:20])[0]
                        cell = DurationCell(
                            row_num, col_num, timedelta(days=cell_value)
                        )
                    elif cell_type == TSTArchives.formulaErrorCellType:
                        cell = ErrorCell(row_num, col_num)
                    elif cell_type == TSTArchives.currencyCellValueType:
                        cell_value = struct.unpack(
                            "<d", storage_buffer_pre_bnc[-12:-4]
                        )[0]
                        cell = FormulaCell(row_num, col_num)
                    elif cell_type == 10:
                        # TODO: Is this realy a number cell (experiments suggest so)
                        if storage_buffer_pre_bnc is None:
                            cell_value = 0.0
                        else:
                            cell_value = struct.unpack(
                                "<d", storage_buffer_pre_bnc[-12:-4]
                            )[0]
                        cell = NumberCell(row_num, col_num, cell_value)
                    else:
                        raise UnsupportedError(  # pragma: no cover
                            f"Unsupport cell type {cell_type} @{self.name}:({row_num},{col_num})"
                        )

                    row_col = (row_num, col_num)
                    if (
                        row_col in self._merge_cells
                        and self._merge_cells[row_col]["merge_type"] == "source"
                    ):
                        cell.is_merged = True
                        cell.size = self._merge_cells[row_col]["size"]
                    else:
                        cell.is_merged = False
                        cell.size = (1, 1)

                    if table_formulas.is_formula(row_num, col_num):
                        try:
                            if table_formulas.is_error(row_num, col_num):
                                formula_key = struct.unpack(
                                    "<h", storage_buffer[12:14]
                                )[0]
                            if cell_type == TSTArchives.numberCellType:  # 2
                                formula_key = struct.unpack(
                                    "<h", storage_buffer[28:30]
                                )[0]
                            elif cell_type == TSTArchives.textCellType:  # 3
                                if len(storage_buffer) > 28:
                                    formula_key = struct.unpack(
                                        "<h", storage_buffer[20:22]
                                    )[0]
                                else:
                                    formula_key = struct.unpack(
                                        "<h", storage_buffer[24:26]
                                    )[0]
                            elif cell_type == TSTArchives.dateCellType:  # 5
                                formula_key = struct.unpack(
                                    "<h", storage_buffer[24:26]
                                )[0]
                            elif cell_type == TSTArchives.boolCellType:  # 6
                                formula_key = struct.unpack(
                                    "<h", storage_buffer[24:26]
                                )[0]
                            elif cell_type == TSTArchives.durationCellType:  # 7
                                formula_key = struct.unpack(
                                    "<h", storage_buffer[24:26]
                                )[0]
                            elif cell_type == TSTArchives.formulaErrorCellType:  # 8
                                formula_key = struct.unpack(
                                    "<h", storage_buffer[12:14]
                                )[0]
                            elif cell_type == TSTArchives.currencyCellValueType:  # 9
                                formula_key = struct.unpack(
                                    "<h", storage_buffer[32:34]
                                )[0]
                            elif cell_type == 10:
                                # Might be [36:38]
                                formula_key = struct.unpack(
                                    "<h", storage_buffer[32:34]
                                )[0]
                            else:
                                raise UnsupportedError(  # pragma: no cover
                                    f"Unsupported formula cell type '{cell_type}' @{self.name}:({row_num},{col_num})"
                                )

                            logging.debug(
                                "@[%d,%d]: formula_key=%d, storage_buffer=%s",
                                row_num,
                                col_num,
                                formula_key,
                                ":".join(["{0:02x}".format(b) for b in storage_buffer]),
                            )

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
            self.__table_cells.append(row)

        return self.__table_cells

    @property
    def _table_strings(self):
        if not hasattr(self, "__table_strings"):
            strings_id = self._table.base_data_store.stringTable.identifier
            self.__table_strings = {
                x.key: x.string for x in self._object_store[strings_id].entries
            }
        return self.__table_strings

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

        rows = self.rows()
        if row_num >= len(rows):
            raise IndexError(f"row {row_num} out of range")
        if col_num >= len(rows[row_num]):
            raise IndexError(f"coumn {col_num} out of range")
        return rows[row_num][col_num]

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

        rows = self.rows()
        for row_num in range(min_row, max_row + 1):
            if values_only:
                yield tuple(
                    cell.value or None for cell in rows[row_num][min_col : max_col + 1]
                )
            else:
                yield tuple(rows[row_num][min_col : max_col + 1])

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

        rows = self.rows()
        for col_num in range(min_col, max_col + 1):
            if values_only:
                yield tuple(row[col_num].value for row in rows[min_row : max_row + 1])
            else:
                yield tuple(row[col_num] for row in rows[min_row : max_row + 1])


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
            idx = next(
                (i for i, x in enumerate(offsets[col_num + 1 :]) if x >= 0), None
            )
            if idx is None:
                end = len(storage_buffer)
            else:
                end = offsets[col_num + idx + 1]
        data.append(storage_buffer[start:end])

    return data
