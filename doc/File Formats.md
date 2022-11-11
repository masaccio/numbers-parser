# Numbers file format

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
  "from_to_range": [
      { "from_coord": { "column": 0, "row": 0 },
        "refers_to_rect": {
            "origin": {"column": 0,  "row": 0 },
            "size": { "num_columns": 2 }
        }
      }
  ],
  "to_owner_id": 1
```

##  Formula ranges

 The `TSCE.CalculationEngineArchive` contains `formula_owner_info` records inside `dependency_tracker`. Each of these has a list of  `cell_dependencies` that can be used to identify which cells contain formulas:

 ``` json
"cell_dependencies": {
    "cell_record": [
        {
            "column": 0,
            "contains_a_formula": true,
            "edges": { "packed_edge_without_owner": [ 16777216 ] },
            "row": 1
        },
        {
            "column": 1,
            "contains_a_formula": true,
            "edges": { "packed_edge_without_owner": [16777216 ] },
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

## Random unformed notes

Adding a new sheet to a document generates a new set of nodes in the `contained_tracked_reference` of a `TSCE.TrackedReferenceStoreArchive`:

``` json
{ "ast": { "AST_node": [ {
        "AST_node_type": "CELL_REFERENCE_NODE",
        "AST_column": { "column": 1, "absolute": true },
        "AST_row": { "row": 0, "absolute": true },
        "AST_cross_table_reference_extra_info": { "table_id": "fffe000f-b02b-d298-8447-808f21eb01ed" } } ]
  },
  "formula_id": 9
},
{ "ast": { "AST_node": [ {
        "AST_node_type": "CELL_REFERENCE_NODE",
        "AST_column": { "column": 0, "absolute": true },
        "AST_row": { "row": 1, "absolute": true },
        "AST_cross_table_reference_extra_info": { "table_id": "fffe000f-b02b-d298-8447-808f21eb01ed" } } ] },
  "formula_id": 10
},
{ "ast": { "AST_node": [ {
        "AST_node_type": "CELL_REFERENCE_NODE",
        "AST_column": { "column": 0, "absolute": true },
        "AST_row": { "row": 2, "absolute": true },
        "AST_cross_table_reference_extra_info": { "table_id": "fffe000f-b02b-d298-8447-808f21eb01ed" } } ] },
  "formula_id": 11
},
{ "ast": { "AST_node": [ {
        "AST_node_type": "CELL_REFERENCE_NODE",
        "AST_column": { "column": 0, "absolute": true },
        "AST_row": { "row": 3, "absolute": true },
        "AST_cross_table_reference_extra_info": { "table_id": "fffe000f-b02b-d298-8447-808f21eb01ed" } } ] },
  "formula_id": 12
}
```

This archive is referenced in a `TSCE.NamedReferenceManagerArchive`:

``` json
"reference_tracker": { "identifier": "905007" }
```

And this is in turn referenced inside the document's `TSCE.CalculationEngineArchive`:

``` json
"named_reference_manager": { "identifier": "904994" }
```

The `table_id` referred to above matches a `base_owner_uid` in a number of `TSCE.FormulaOwnerDependenciesArchive`. Each of these has a different `internal_formula_owner_id` and `formula_owner_uid`. The mappign between the `internal_formula_id` and the `formula_owner_uid` is additonally stored in the Calculation Engine's `owner_id_map`:

``` json
"owner_id_map": {
  "map_entry": [
    { "internal_owner_id": 18, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01ed" },
    { "internal_owner_id": 19, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01f1" },
    { "internal_owner_id": 21, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01f8" },
    { "internal_owner_id": 22, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01f0" },
    { "internal_owner_id": 23, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01f7" },
    { "internal_owner_id": 24, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01f2" },
    { "internal_owner_id": 25, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01f3" },
    { "internal_owner_id": 26, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01f5" },
    { "internal_owner_id": 27, "owner_id": "fffe000f-b02b-d298-8447-808f21eb0210" },
    { "internal_owner_id": 28, "owner_id": "fffe000f-b02b-d298-8447-808f21eb01f6" },
  ]
}
```

Calculation Engine's `formula_owner_dependencies` is a list of references to all the `TSCE.FormulaOwnerDependenciesArchive`.

Two of the archives also have non-empty `tiled_cell_dependencies` references to a `TSCE.CellRecordTileArchive`:

``` json
"tiled_cell_dependencies": { "cell_record_tiles": [ { "identifier": "906718" } ] },
"tiled_cell_dependencies": { "cell_record_tiles": [ { "identifier": "906719" } ] },
```

These `TSCE.CellRecordTileArchive` archives contain:

``` json
"internal_owner_id": 27,
"tile_column_begin": 0,
"tile_row_begin": 0,
```

And for reference ID 906719:

``` json
"internal_owner_id": 26,
"tile_column_begin": 0,
"tile_row_begin": 0,
"cell_records": [
  { "column": 0, "row": 0, "expanded_edges": { "edge_without_owner_rows": [ 0 ], "edge_without_owner_columns": [ 2 ] } },
  { "column": 3, "row": 0, "expanded_edges": {} },
  { "column": 7, "row": 0, "expanded_edges": { "edge_with_owner_rows": [ 0 ], "edge_with_owner_columns": [ 11 ], "internal_owner_id_for_edge": [ 27 ] } }
],
```

`TST.HeaderNameMgrArchive` contains a list of `per_tables` with an entry for each table in addition to the first one

``` json
"owner_uid": "00000000-0000-0000-0000-00000000029a",
"nrm_owner_uid": "924db583-f350-4a87-5f42-1278be73332e",
"per_tables": [
  {    
    "table_uid": "6a4a5281-7b06-f5a1-904b-7f9ec784b368",
    "per_table_precedent": { "column": 1, "row": 1 },
    "header_row_uids": [ "19332f73-d13d-5d45-0716-5d5b9d909765" ],
    "header_column_uids": [ "20cc67a1-e317-d6c4-1e4d-4e85b897cacb" ]
  },   
  {    
    "table_uid": "e2cf803a-ba85-788f-3f41-0709393e975c",
    "per_table_precedent": { "column": 2, "row": 1 },
    "header_row_uids": [ "19332f73-d13d-5d45-0716-5d5b9d909765" ],
    "header_column_uids": [ "20cc67a1-e317-d6c4-1e4d-4e85b897cacb" ]
  }    
],   
"name_frag_tiles": [ { "identifier": "905504" } ],
```

And a further one the `tiled_cell_dependencies` plus:

``` json
"cell_dependencies": { "cell_record": [
    { "column": 0, "row": 0, "expanded_edges": { "edge_without_owner_rows"   : [ 0 ], "edge_without_owner_columns": [ 2 ] } },
    { "column": 3, "row": 0, "expanded_edges": {} },
    { "column": 7, "row": 0, "expanded_edges": { "edge_with_owner_rows"      : [  0 ], "edge_with_owner_columns"   : [ 11 ], "internal_owner_id_for_edge": [ 27 ] } } ] },
```

The UUID of the `TSCE.TrackedReferenceStoreArchive` that contains formula ASTs is also the `nrm_owner_uid` in the `TST.HeaderNameMgrArchive` and has an entry in `dependency_tracker.formula_owner_info` where `cell_record` is a list of references to the formula IDs as column numbers"

``` json
{    
  "formula_owner_id": "924db583-f350-4a87-5f42-1278be73332e",
  "cell_dependencies": {
    "cell_record": [
      { "column": 27, "row": 0, "contains_a_formula": true,
        "edges": {"packed_edge_with_owner": [ 0x1010000 ], "internal_owner_id_for_edge": [ 0 ]} },
      { "column": 28, "row": 0, "contains_a_formula": true,
        "edges": {"packed_edge_with_owner": [ 0x1020000 ], "internal_owner_id_for_edge": [ 0 ]} },
      { "column": 33, "row": 0, "contains_a_formula": true,
        "edges": {"packed_edge_with_owner": [ 0x1000001 ], "internal_owner_id_for_edge": [ 0 ]} },
      { "column": 34, "row": 0, "contains_a_formula": true,
        "edges": {"packed_edge_with_owner": [ 0x1000002 ], "internal_owner_id_for_edge": [ 0 ]} },
      { "column": 35, "row": 0, "contains_a_formula": true,
        "edges": {"packed_edge_with_owner": [ 0x1000003 ], "internal_owner_id_for_edge": [ 0 ]} } 
    ]    
  },  
  "range_dependencies": {},
  "volatile_dependencies": {},
  "spanning_column_dependencies": {
    "total_range_for_deleted_table": { "top_left_column": 0x7fff, "top_left_row": 0x7fffffff,
                                       "bottom_right_column": 0x7fff, "bottom_right_row": 0x7fffffff },
    "body_range_for_deleted_table" : { "top_left_column": 0x7fff, "top_left_row": 0x7fffffff,
                                        "bottom_right_column": 0x7fff, "bottom_right_row": 0x7fffffff }
  },
  "spanning_row_dependencies": {
    "total_range_for_deleted_table": { "top_left_column": 0x7fff, "top_left_row": 0x7fffffff,
                                       "bottom_right_column": 0x7fff, "bottom_right_row": 0x7fffffff },
    "body_range_for_deleted_table" : { "top_left_column": 0x7fff, "top_left_row": 0x7fffffff,
                                       "bottom_right_column": 0x7fff, "bottom_right_row": 0x7fffffff }
  },
  "whole_owner_dependencies": {},
  "cell_errors": {}
}
```

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
