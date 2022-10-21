# Numbers file format

##  Debugging notes

empty-1 versus empty-2 differences

* dependency_tracker.formula_owner_info entries:
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0ea"
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b104"
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0e9" (cell_records 0,0; 0,3; 0,7)
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0e7"
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0e6"
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0eb"
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0e4"
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0ec"
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0e5"
  * "formula_owner_id": "0xe07d1144_82dea39d_4f405d4b_bec5b0e1"
  * "formula_owner_id": "0x924db583_f3504a87_5f421278_be73332e" (cell_records 0,27; 0,33; 0; 34)
* owner_id_map.map_entry added in empty-2:
  * 18 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0e1
  * 19 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0e5
  * 20 -> 0x924db583_f3504a87_5f421278_be73332e
  * 21 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0ec
  * 22 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0e4
  * 23 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0eb
  * 24 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0e6
  * 25 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0e7
  * 26 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0e9
  * 27 -> 0xe07d1144_82dea39d_4f405d4b_bec5b104
  * 28 -> 0xe07d1144_82dea39d_4f405d4b_bec5b0ea
* "number_of_formulas": 5 (empty-1) / 11 (empty-2) references to TSCE.FormulaOwnerDependenciesArchive
* Addition of a TSCE.CellRecordTileArchive:
  * internal_owner_id=27 and cell_records 0,0; 0,3; 0,7
  * internal_owner_id=26 and cell_records 0,0; 0,3; 0,7
  * internal_owner_id=18
  * internal_owner_id=20 and cell_records 0,27
  * internal_owner_id=20 and cell_records 0,33; 0; 34
* TSCE.FormulaOwnerDependenciesArchive with formula_owner_uid=0x924db583_f3504a87_5f421278_be73332e (ID 20)
  * same internal_formula_owner_id=20
  * addition of cell_records 0,27; 0,33; 0; 34
  * addition of 2 tiled_cell_dependencies.cell_record_tiles
* Additional CELL_REFERENCE_NODE in TSCE.TrackedReferenceStoreArchive:
  * formula_id: 27 = 0,1
  * formula_id: 33 = 1,0
  * formula_id: 34 = 2,0
* Addition of TST.TableInfoArchive
  * group_by_uuid: 0xe07d1144_82dea39d_4f405d4b_bec5b0e9
  * hidden_states_uuid: 0xe07d1144_82dea39d_4f405d4b_bec5b0e5
  * parent.identifier = 905467 (SheetArchive 'sheet 2')
* Addition of TST.TableModelArchive
  * haunted_owner.owner_id = 0xe07d1144_82dea39d_4f405d4b_bec5b104 (ID 27)
  * group_by_uid: "0xe07d1144_82dea39d_4f405d4b_bec5b0e9 (ID 26)
  * merge_owner.owner_id = 0xe07d1144_82dea39d_4f405d4b_bec5b0e6 (ID 24)
  * conditional_style_formula_owner_id: 0xe07d1144_82dea39d_4f405d4b_bec5b0e4 (ID 22)
* TSCE.TrackedReferenceStoreArchive has uuid=0xe07d1144_82dea39d_4f405d4b_bec5b0e7 (ID 25)
* TST.ColumnRowUIDMapArchive
  * sorted_column_uids=0x20cc67a1_e317d6c4_1e4d4e85_b897cacb, 0xe5e7a592_dca4a21c_5d41b28a_8aa852e2
  * sorted_row_uids=0x19332f73_d13d5d45_07165d5b_9d909765, 0x1aacaf86_25f99bf8_d3600774_36b3ac91, 0x90f16628_dfcccafd_450f82a9_fa649d54

TSCE.TrackedReferenceStoreArchive.contained_tracked_reference:

``` json
{ "ast": { "AST_node": [ {
        "AST_node_type": "CELL_REFERENCE_NODE",
        "AST_column": { "column": 1, "absolute": true },
        "AST_row": { "row": 0, "absolute": true },
        "AST_cross_table_reference_extra_info": { "table_id": "e07d1144-82de-a39d-4f40-5d4bbec5b0e1" } } ] },
  "formula_id": 27
},
{ "ast": { "AST_node": [ {
        "AST_node_type": "CELL_REFERENCE_NODE",
        "AST_column": { "column": 0, "absolute": true },
        "AST_row": { "row": 1, "absolute": true },
        "AST_cross_table_reference_extra_info": { "table_id": "e07d1144-82de-a39d-4f40-5d4bbec5b0e1" } } ] },
  "formula_id": 33
},
{ "ast": { "AST_node": [ {
        "AST_node_type": "CELL_REFERENCE_NODE",
        "AST_column": { "column": 0, "absolute": true },
        "AST_row": { "row": 2, "absolute": true },
        "AST_cross_table_reference_extra_info": { "table_id": "e07d1144-82de-a39d-4f40-5d4bbec5b0e1" } } ] },
  "formula_id": 34
}
```

TSCE.CalculationEngineArchive.owner_id_map.map_entry:

``` json
{ "internal_owner_id": 34, "owner_id": "2bf9df9b-380b-b0b2-684e-15541ca116f9" },
{ "internal_owner_id": 33, "owner_id": "627d5f1b-72c2-50a8-e941-1c5615559a4f" },
{ "internal_owner_id": 27, "owner_id": "e07d1144-82de-a39d-4f40-5d4bbec5b104" },
```

``` json
{    
  "formula_owner_uid": "e07d1144-82de-a39d-4f40-5d4bbec5b104", // matches internal_owner_id=27
  "internal_formula_owner_id": 27,
  "owner_kind": 35,
  "cell_dependencies": {},
  "range_dependencies": {},
  "volatile_dependencies": {
    "volatile_time_cells": {},
    "volatile_random_cells": {},
    "volatile_locale_cells": {},
    "volatile_sheet_table_name_cells": {},
    "volatile_remote_data_cells": {},
    "volatile_geometry_cell_refs": {}
  },   
  "spanning_column_dependencies": {
    "total_range_for_table": {
      "top_left_column": 0x7fff, "top_left_row": 0x7fffffff,
      "bottom_right_column": 0x7fff, "bottom_right_row": 0x7fffffff
    },   
    "body_range_for_table": {
      "top_left_column": 0x7fff, "top_left_row": 0x7fffffff,
      "bottom_right_column": 0x7fff, "bottom_right_row": 0x7fffffff
    }    
  },   
  "spanning_row_dependencies": {
    "total_range_for_table": {
      "top_left_column": 0x7fff, "top_left_row": 0x7fffffff,
      "bottom_right_column": 0x7fff, "bottom_right_row": 0x7fffffff
    },   
    "body_range_for_table": {
      "top_left_column": 0x7fff, "top_left_row": 0x7fffffff,
      "bottom_right_column": 0x7fff, "bottom_right_row": 0x7fffffff
    }    
  },   
  "whole_owner_dependencies": { "dependent_cells": {}},   
  "cell_errors": {},
  "base_owner_uid": "e07d1144-82de-a39d-4f40-5d4bbec5b0e1",
  "tiled_cell_dependencies": { "cell_record_tiles": [ { 
    "identifier": "905737" } ] }, // -> CellRecordTileArchive
  "uuid_references": {},
  "tiled_range_dependencies": {},
  "_pbtype": "TSCE.FormulaOwnerDependenciesArchive"
}
```

``` json
{
  "internal_owner_id": 27,
  "tile_column_begin": 0,
  "tile_row_begin": 0,
  "_pbtype": "TSCE.CellRecordTileArchive"
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
