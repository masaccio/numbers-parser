from functools import lru_cache

from numbers_parser.containers import ObjectStore
from numbers_parser.cell import xl_rowcol_to_cell
from numbers_parser.exceptions import UnsupportedError


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

    @lru_cache(maxsize=None)
    def table_ids(self, sheet_id):
        table_info_ids = self.find_refs("TableInfoArchive")
        return [
            self.objects[t_id].tableModel.identifier
            for t_id in table_info_ids
            if self.objects[t_id].super.parent.identifier == sheet_id
        ]

    @lru_cache(maxsize=None)
    def table_row_headers(self, table_id):
        bds = self.objects[table_id].base_data_store
        return [
            h.numberOfCells
            for h in self.objects[bds.rowHeaders.buckets[0].identifier].headers
        ]

    @lru_cache(maxsize=None)
    def table_row_columns(self, table_id):
        bds = self.objects[table_id].base_data_store
        return [
            h.numberOfCells for h in self.objects[bds.columnHeaders.identifier].headers
        ]

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
        for x in self.objects[strings_id].entries:
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
        for dependency_id in self.find_refs("FormulaOwnerDependenciesArchive"):
            obj = self.objects[dependency_id]
            if obj.HasField("base_owner_uid") and obj.HasField("formula_owner_uid"):
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
    def error_cell_ranges(self, table_id: int) -> list:
        """Exract all the formula error cell ranges for the Table."""
        cell_errors = {}
        table_base_id = self.table_base_id(table_id)
        ce_id = self.find_refs("CalculationEngineArchive")[0]
        for finfo in self.objects[ce_id].dependency_tracker.formula_owner_info:
            if finfo.HasField("cell_dependencies"):
                formula_owner_id = uuid(finfo.formula_owner_id)
                if formula_owner_id == table_base_id:
                    for cell_error in finfo.cell_errors.errors:
                        cell_errors[
                            (cell_error.coordinate.row, cell_error.coordinate.column)
                        ] = cell_error.error_flavor
        return cell_errors

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

    def node_to_cell_ref(self, row_num: int, col_num: int, node):
        table_name = None
        if node.HasField("AST_cross_table_reference_extra_info"):
            table_ids = self.find_refs("TableInfoArchive")
            tables = [
                self.objects[self.objects[table_id].tableModel.identifier]
                for table_id in table_ids
            ]
            tables_uuids = [self.table_base_id(t) for t in tables]
            table_uuid = uuid(node.AST_cross_table_reference_extra_info.table_id)
            table_name = tables[tables_uuids.index(table_uuid)].table_name
        else:
            pass
        if node.AST_column.absolute:
            row = node.AST_row.row
        else:
            row = row_num + node.AST_row.row
        if node.AST_column.absolute:
            col = node.AST_column.column
        else:
            col = col_num + node.AST_column.column
        try:
            ref = xl_rowcol_to_cell(
                row,
                col,
                row_abs=node.AST_column.absolute,
                col_abs=node.AST_column.absolute,
            )
            if table_name is not None:
                return f"{table_name}::{ref}"
            else:
                return ref
        except IndexError:
            return f"INVALID[{row_num},{col_num}]"

    @lru_cache(maxsize=None)
    def formula_ast(self, table_id: int):
        bds = self.objects[table_id].base_data_store
        formula_table_id = bds.formula_table.identifier
        formula_table = self.objects[formula_table_id]
        formulas = {}
        for formula in formula_table.entries:
            formulas[formula.key] = formula.formula.AST_node_array.AST_node
        return formulas


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
