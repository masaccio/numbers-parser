# Numbers file format

[SheetsJS has documented](https://oss.sheetjs.com/notes/iwa/#hyperlinks) some of the core file format details. These notes are additional information tracking other discoveries.

## Cell formats

Bytes offets 6 and 7 include other values which are important to Numbers

| Offset | Notes |
| ------ | ----- |
| 0      | Storage version |
| 1      | Currency cells are type=10 which is not an enum identified in the protofus protos |
| 2-5    | Unused |
| 6      | 0x01 is set when there is a number format ID |
|        | 0x02 is set when there is a currency format ID |
|        | 0x04 is set when there is a duration format ID |
|        | 0x08 is set when there is a date format ID |
|        | 0x20 is set when there is a bool format ID |
|        | 0x80 is set when there is a string ID |
| 7      | Some cells have 0x80 for this byte. Unclear why |

This is then followed by the flags as described in the SheetsJS docs.

## Owner IDs

The single `TSCE.CalculationEngineArchive` contains a number of owner ID mappings:

* `formula_owner_id`: there are a large number of these (one per formula?) where the lower 16 bits of the 128-bit UUID varies by formula ID but the rest of the UUID is constant.
* `conditional_style_formula_owner_id`: one per table.
* `base_owner_uid`: one per table.
* `nrm_owner_uid`: one per document.

`TST.HeaderNameMgrArchive` maps the following UUIDs:

* `nrm_owner_uid`: TBD but this UUID matches the `formula_owner_uid` in one of the many `TSCE.FormulaOwnerDependenciesArchive` archives.
* `owner_uid`:
* `per_tables`: a list of `TST.HeaderNameMgrArchive.PerTableArchive` archives that contains references to `table_uid`, `header_row_uids` and `header_column_uids`

`TSCE.CalculationEngineArchive` contains the following:

* `dependency_tracker` (a `.TSCE.DependencyTrackerArchive):
  * `formula_owner_dependencies`: a list of references to `TSCE.FormulaOwnerDependenciesArchive` archives that contain information about where formulas are located and a mapping between a `base_owner_uid` and `formula_owner_id`
  * `formula_owner_info`: a list of `.TSCE.FormulaOwnerInfoArchive` archives that contain information about where formulas are located and refer to a `formula_owner_id`.
  * `owner_id_map`: a list of `TSCE.OwnerIDMapArchive` archives that maps a `internal_owner_id` number to a UUID, which is one of the various `conditional_style_formula_owner_id` or `formula_owner_id` UUIDs:

    ```json
    { "internal_ownerId": 33, "owner_id": 0x3cb03f23_c26dda92_1e4bfcc0_8750e563 },
    ```

The (`TST.TableModelArchive`) for each table contains a number of UUIDs:

* `haunted_owner`: matches the `formula_owner_id` of one `TSCE.FormulaOwnerDependenciesArchive` which itself will contain a `base_owner_uid` which is referred to as the owner of many objects. This can be used to escalish a link between a specific table and a `base_owner_id`. The `base_owner_uid` is used in a number of ways:
  * Resolving cross-table formula references (see `_NumbersModel.node_to_ref()` for how this is used in `numbers-parser`)
  * Locating all cells that contain formulas (see `_NumbersModel.formula_cell_ranges()`)
  * Naming all merge cell ranges in a from/to notation (see `_NumbersModel.calculate_merge_cell_ranges()`)
* `merge_owner`:

## Merge ranges

Merge ranges are stored in a number of structures, but the simplest is a `TSCE.RangePrecedentsTileArchive` which exists for each Table in the Document. These archives contain a `from_to_range`` list which has the merge ranges associated with an integer owner ID key.

 The Owner IDs need to be extracted from the Calculation Engine using the `owner_id_map` described in the [UUID maps of `TSCE.CalculationEngineArchive`](#owner-ids)

``` json
 {
   "from_to_range": [
      { "from_coord": { "column": 0, "row": 0 },
        "refers_to_rect": {
            "origin": {"column": 0,  "row": 0 },
            "size": { "num_columns": 2 } } }
  ],
  "to_owner_id": 1
}
```

##  Formula ranges

 The `TSCE.CalculationEngineArchive` contains `formula_owner_info` records inside `dependency_tracker`. Each of these has a list of  `cell_dependencies` that can be used to identify which cells contain formulas:

 ``` json
"cell_dependencies": {
    "cell_record": [
        {
            "column": 0,
            "contains_a_formula": true,
            "edges": { "packed_edge_without_owner": [ 0x1000000 ] },
            "row": 1
        },
        {
            "column": 1,
            "contains_a_formula": true,
            "edges": { "packed_edge_without_owner": [0x1000000 ] },
            "row": 1
        }
    ]
}
```

## Creating a Formula Owner for a new Table Model

Each table has a `TSCE.FormulaOwnerDependenciesArchive` with `owner_kind` 1 which identifies a TableModel. For the first table in a document, the `tiled_cell_dependencies` is empty but subsequent sheet tables have a reference to `cell_record_tiles` in them (TODO: what happens for multiple tables in the first sheet). `spanning_column_dependencies` and `spanning_row_dependencies` reflect the size of the table body and headers:

``` json
{
  "formula_owner_uid": "0bfb3e8b-e953-37a8-a845-34053047a144",
  "internal_formula_owner_id": 0,
  "owner_kind": 1,
  "cell_dependencies": {},
  "range_dependencies": {},
  "volatile_dependencies": { "volatile_time_cells": {}, "volatile_random_cells": {}, 
    "volatile_locale_cells": {}, "volatile_sheet_table_name_cells": {}, 
    "volatile_remote_data_cells": {}, "volatile_geometry_cell_refs": {}
  },
  "spanning_column_dependencies": {
    "total_range_for_table": { "top_left_column": 0, "top_left_row": 0, "bottom_right_column": 2, "bottom_right_row": 3 },
    "body_range_for_table": { "top_left_column": 1, "top_left_row": 1, "bottom_right_column": 2, "bottom_right_row": 3 }
  },
  "spanning_row_dependencies": {
    "total_range_for_table": { "top_left_column": 0, "top_left_row": 0, "bottom_right_column": 2, "bottom_right_row": 3 },
    "body_range_for_table": { "top_left_column": 1, "top_left_row": 1, "bottom_right_column": 2, "bottom_right_row": 3 }
  },
  "whole_owner_dependencies": { "dependent_cells": {} },
  "cell_errors": {},
  "formula_owner": { "identifier": "905976" },
  "tiled_cell_dependencies": { "cell_record_tiles": [ { "identifier": "906191" } ] },
  "uuid_references": {},
  "tiled_range_dependencies": {},
}
```

Specific example files:

* `empty-1.numbers`:
  * Sheet 1 has 1 header row, 1 header column, 1 column and 2 rows of content
  * `formula_owner_uid` is `6a4a5281-7b06-f5a1-904b-7f9ec784b368`
  * `formula_owner` points to `TST.TableInfoArchive` 904712
  * `internal_formula_owner_id` is 6
* `empty-2.numbers`:
  * Sheet 2 has 1 header row, 1 header column, 2 column2 and 3 rows of content
  * `formula_owner_uid` is `0bfb3e8b-e953-37a8-a845-34053047a144`
  * `internal_formula_owner_id` is 0
  * `formula_owner` points to `TST.TableInfoArchive` 905976
  * `tiled_cell_dependencies.cell_record_tiles` points to a `TSCE.CellRecordTileArchive` with identifier 906191
* `empty-3.numbers`:
  * Sheet 3 has 1 header row, 1 header column, 3 column2 and 4 rows of content
  * `formula_owner_uid` is `2dc400b4-257e-1892-7b4a-8d539034f423`
  * `internal_formula_owner_id` is 1
  * `formula_owner` points to `TST.TableInfoArchive` 906522
  * `tiled_cell_dependencies.cell_record_tiles` points to `TSCE.CellRecordTileArchive` with identifier 906768

For each of the `formula_owner_uid` and `internal_formula_owner_id` pairs is added to the `owner_id_map` in the `TSCE.CalculationEngineArchive`.

For `TSCE.FormulaOwnerDependenciesArchive` with a non-empty `tiled_cell_dependencies`, the referenced `TSCE.CellRecordTileArchive` is is simply:

``` json:
{
    "internal_owner_id": 0,
    "tile_column_begin": 0,
    "tile_row_begin": 0,
}
```

Where the `internal_owner_id` is the `internal_formula_owner_id` from the associated `TSCE.FormulaOwnerDependenciesArchive`.

## Header Map

The single `TST.HeaderNameMgrTileArchive` describes the relationships between column and row headers by their UUIDs. Each row or column has a `TST.HeaderNameMgrTileArchive.NameFragmentArchive`:

``` json
{    
    "name_fragment": "header4",
    "name_precedent": { "column": 4, "row": 1 },   
    "uses_of_name_fragment": {
    "owner_entries": [
        { "owner_uid": "0xfffe000f_b02bd298_8447808f_21eb01ed",
          "coord_set": {
            "column_entries": [{ "column": "0xe5e7a592_dca4a21c_5d41b28a_8aa852e2",
                                 "row_set": [ "0x19332f73_d13d5d45_07165d5b_9d909765" ]
                               }]
            } 
        }    
    ]    
}
```

Each row and column coordinate has a pair of UUIDs for the row and column which match the UUIDs listed in the following archive fields:

* `TST.HeaderNameMgrArchive.per_tables[].header_row_uids`
* `TST.HeaderNameMgrArchive.per_tables[].header_column_uids`
* `TST.ColumnRowUIDMapArchive.sorted_column_uids`
* `TST.ColumnRowUIDMapArchive.sorted_row_uids`
