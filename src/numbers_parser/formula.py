from numbers_parser.containers import ObjectStore
from numbers_parser.exceptions import UnsupportedError

# from numbers_parser.document import Table
from numbers_parser.cell import FormulaCell

from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives


class TableFormulas:
    def __init__(self, object_store: ObjectStore, table):
        self._object_store = object_store
        self._table = table
        self._formulas = get_formula_ast(object_store, table)
        self._error_cells = get_error_cell_ranges(object_store, table)

    def is_formula(self, row, col):
        if not (hasattr(self, "_formula_cells")):
            self._formula_cells = get_formula_cell_ranges(self._object_store, self._table)
        return (row, col) in self._formula_cells

    def is_error(self, row, col):
        if not (hasattr(self, "_error_cells")):
            self._error_cells = get_error_cell_ranges(self._object_store, self._table)
        return (row, col) in self._error_cells

    def ast(self, formula_key):
        if not (hasattr(self, "_ast")):
            self._formula_asts = get_formula_ast(self._object_store, self._table)
        return self._formula_asts[formula_key]


def get_formula_ast(object_store, table):
    formula_type_lookup = {
        k: v.name
        for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()
    }
    formula_table_id = table.base_data_store.formula_table.identifier
    formula_table = object_store[formula_table_id]
    formulas = {}
    for formula in formula_table.entries:
        ast_nodes = []
        for node in formula.formula.AST_node_array.AST_node:
            node_type = formula_type_lookup[node.AST_node_type]
            if node_type == "FUNCTION_NODE":
                ast_nodes.append(
                    {
                        "type": node_type,
                        "node_index": node.AST_function_node_index,
                        "num_args": node.AST_function_node_numArgs,
                    }
                )
            elif node_type == "NUMBER_NODE":
                ast_nodes.append(
                    {
                        "type": node_type,
                        "decimal_high": node.AST_number_node_decimal_high,
                        "decimal_low": node.AST_number_node_decimal_low,
                        "number": node.AST_number_node_number,
                    }
                )
            elif node_type == "CELL_REFERENCE_NODE":
                ast_nodes.append(
                    {
                        "type": node_type,
                        "row": node.AST_row.row,
                        "column": node.AST_column.column,
                        "row_is_abs": node.AST_column.absolute,
                        "col_is_abs": node.AST_column.absolute,
                    }
                )
            elif node_type == "STRING_NODE":
                ast_nodes.append(
                    {
                        "type": node_type,
                        "string": node.AST_string_node_string,
                    }
                )
            elif node_type == "PREPEND_WHITESPACE_NODE":
                ast_nodes.append(
                    {
                        "type": node_type,
                        "whitespace": node.AST_whitespace,
                    }
                )
            elif any(
                [
                    node_type == "ADDITION_NODE",
                    node_type == "BEGIN_EMBEDDED_NODE_ARRAY",
                    node_type == "BOOLEAN_NODE",
                    node_type == "COLON_NODE",
                    node_type == "CONCATENATION_NODE",
                    node_type == "DIVISION_NODE",
                    node_type == "END_THUNK_NODE",
                    node_type == "EQUAL_TO_NODE",
                    node_type == "GREATER_THAN_NODE",
                    node_type == "MULTIPLICATION_NODE",
                    node_type == "NOT_EQUAL_TO_NODE",
                    node_type == "SUBTRACTION_NODE",
                ]
            ):
                ast_nodes.append({"type": node_type})
            else:
                raise UnsupportedError(f"Unsupported formula type {node_type}")  # pragma: no cover

            #  TODO: deal with error cells

        formulas[formula.key] = {"ast": ast_nodes}
    return formulas


def get_owner_id_map(object_store: ObjectStore):
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
    calculation_engine_id = object_store.find_refs("CalculationEngineArchive")[0]
    owner_id_map = {}
    for e in object_store[calculation_engine_id].dependency_tracker.owner_id_map.map_entry:
        owner_id_map[e.internal_owner_id] = get_uuid(e.owner_id)
    return owner_id_map


def get_table_base_id(object_store: ObjectStore, table):
    """"Finds the UUID of a table"""
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
    haunted_owner = get_uuid(table.haunted_owner.owner_uid)
    for dependency_id in object_store.find_refs("FormulaOwnerDependenciesArchive"):
        obj = object_store[dependency_id]
        if obj.HasField("base_owner_uid") and obj.HasField("formula_owner_uid"):
            base_owner_uid = get_uuid(obj.base_owner_uid)
            formula_owner_uid = get_uuid(obj.formula_owner_uid)
            if formula_owner_uid == haunted_owner:
                return base_owner_uid


def get_formula_cell_ranges(object_store: ObjectStore, table):
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
    table_base_id = get_table_base_id(object_store, table)
    calculation_engine_id = object_store.find_refs("CalculationEngineArchive")[0]
    for finfo in object_store[calculation_engine_id].dependency_tracker.formula_owner_info:
        if finfo.HasField("cell_dependencies"):
            formula_owner_id = get_uuid(finfo.formula_owner_id)
            if formula_owner_id == table_base_id:
                for cell_record in finfo.cell_dependencies.cell_record:
                    if cell_record.contains_a_formula:
                        cell_records.append((cell_record.row, cell_record.column))
    return cell_records


def get_error_cell_ranges(object_store: ObjectStore, table):
    """Exract all the formula error cell ranges for the Table."""
    cell_errors = {}
    table_base_id = get_table_base_id(object_store, table)
    calculation_engine_id = object_store.find_refs("CalculationEngineArchive")[0]
    for finfo in object_store[calculation_engine_id].dependency_tracker.formula_owner_info:
        if finfo.HasField("cell_dependencies"):
            formula_owner_id = get_uuid(finfo.formula_owner_id)
            if formula_owner_id == table_base_id:
                for cell_error in finfo.cell_errors.errors:
                    cell_errors[
                        (cell_error.coordinate.row, cell_error.coordinate.column)
                    ] = cell_error.error_flavor
    return cell_errors


def get_merge_cell_ranges(object_store: ObjectStore, table):
    """Exract all the merge cell ranges for the Table."""
    # Merge ranges are stored in a number of structures, but the simplest is
    # a TSCE.RangePrecedentsTileArchive which exists for each Table in the
    # Document. These archives contain a fromToRange list which has the merge
    # ranges associated with an Owner ID
    #
    # The Owner IDs need to be extracted from teh Calculation Engine using
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
    owner_id_map = get_owner_id_map(object_store)
    table_base_id = get_table_base_id(object_store, table)

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
                    for col_num in range(col_start, col_end + 1):
                        merge_cells[(row_num, col_num)] = {
                            "merge_type": "ref",
                            "rect": (row_start, col_start, row_end, col_end),
                            "size": (rect.size.num_rows, rect.size.num_columns),
                        }
                merge_cells[(row_start, col_start)]["merge_type"] = "source"
    return merge_cells


def get_uuid(ref: dict) -> int:
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
