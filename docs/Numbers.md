# Numbers file format

[SheetsJS has documented](https://oss.sheetjs.com/notes/iwa/) some of the core file format details. These notes are additional information tracking other discoveries. The project also has an [IWA inspector](https://sheetjs.com/tools/iwa-inspector/) which may be more useful than using `unpack-numbers`.

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

## UUID mapping

References between tables for formulas, merge ranges and likely plenty of other features rely on UUIDs that relate to a `FormulaOwnerDependenciesArchive` which includes:

``` proto
message FormulaOwnerDependenciesArchive {
  required .TSP.UUID formula_owner_uid = 1;
  required uint32 internal_formula_owner_id = 2;
  optional uint32 owner_kind = 3 [default = 0];
  /* ... */
  optional .TSP.UUID base_owner_uid = 12;
}
```

The `OwnerKind` appears to be a Numbers-internal enum whose protobuf equivalent is:

``` proto
enum OwnerKind {
  UNKNOWN = 0;
  TABLE_MODEL = 1;
  CHART = 2;
  CONDITIONAL_STYLE_OWNER = 3;
  HIDDEN_STATES_FOR_ROWS = 4;
  MERGE_OWNER = 5;
  SORT_RULE_OWNER = 6;
  LEGACY_NAMED_REFERENCE_MGR = 7;
  CATEGORIES_GROUP_BY = 8;
  CATEGORY_AGGREGATE_OWNER = 9;
  PENCIL_ANNOTATION_OWNER = 10;
  HIDDEN_STATES_FOR_COLUMNS = 11;
  PIVOT_OWNER = 17;
  HAUNTED_OWNER = 35;
  PIVOT_DATA_MODEL = 100;
  HEADER_NAME_MGR = 200;
  PRACTICAL_GROUP_BY_MIN = 205;
  PRACTICAL_GROUP_BY_MAX = 214;
  BITWISE_GROUP_BY_MIN = 215;
  BITWISE_GROUP_BY_MAX = 1305;
  GROUP_BY_FOR_PIVOT_MIN = 205;
  GROUP_BY_FOR_PIVOT_MAX = 1305
}
```

Within the `TableModel` there is a UUID reference in the `haunted_owner` field which maps to a `FormulaOwnerDependenciesArchive` whose `formula_owner_uid` is the same UUID as the `haunted_owner`. Within this `FormulaOwnerDependenciesArchive`, the `base_owner_uid` is the table's UUID that is used within the document for cross-referencing. The formula owner archive containing this mapping will have an `owner_kind=OwnerKind.HAUNTED_OWNER`.

``` proto
message TableModelArchive {
  /* ... */
  optional .TSCE.HauntedOwnerArchive haunted_owner = 84;
}

message HauntedOwnerArchive {
  required .TSP.UUID owner_uid = 1;
}
```


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

## Captions

`TSWP.StorageArchive` contains the text for the caption along with references to paragraph styles. The caption is stored in an array in the `text` property. The `TSWP.StorageArchive` is not referenced in `Metadata.json` aside from in the `object_uuid_map_entries`.

`TSA.CaptionInfoArchive` references the `TSWP.StorageArchive` via a reference called `owned_storage` but through the following structure. Other `super` references are required:

``` proto
message CaptionInfoArchive {
  required .TSWP.ShapeInfoArchive super = 1;
  optional .TSP.Reference placement = 2;
  optional .TSD.CaptionOrTitleKind childInfoKind = 3;
}

message ShapeInfoArchive {
  required .TSD.ShapeArchive super = 1;
  optional .TSP.Reference deprecated_storage = 2 [deprecated = true];
  optional .TSP.Reference text_flow = 3;
  optional .TSP.Reference owned_storage = 4;
  optional bool is_text_box = 6;
}

message ShapeArchive {
  required .TSD.DrawableArchive super = 1;
  optional .TSP.Reference style = 2;
  optional .TSD.PathSourceArchive pathsource = 3;
  optional .TSD.LineEndArchive head_line_end = 4 [deprecated = true];
  optional .TSD.LineEndArchive tail_line_end = 5 [deprecated = true];
  optional float strokePatternOffsetDistance = 6;
}

message StorageArchive {
  enum KindType {
    BODY = 0;
    HEADER = 1;
    FOOTNOTE = 2;
    TEXTBOX = 3;
    NOTE = 4;
    CELL = 5;
    UNCLASSIFIED = 6;
    TABLEOFCONTENTS = 7;
    UNDEFINED = 8;
  }
  optional .TSWP.StorageArchive.KindType kind = 1 [default = TEXTBOX];
  optional .TSP.Reference style_sheet = 2;
  repeated string text = 3;
  optional bool has_itext = 4 [default = false];
  optional bool in_document = 10 [default = false];
  optional .TSWP.ObjectAttributeTable table_para_style = 5;
  optional .TSWP.ParaDataAttributeTable table_para_data = 6;
  optional .TSWP.ObjectAttributeTable table_list_style = 7;
  optional .TSWP.ObjectAttributeTable table_char_style = 8;
  optional .TSWP.ObjectAttributeTable table_attachment = 9;
  optional .TSWP.ObjectAttributeTable table_smartfield = 11;
  optional .TSWP.ObjectAttributeTable table_layout_style = 12;
  optional .TSWP.ParaDataAttributeTable table_para_starts = 14;
  optional .TSWP.ObjectAttributeTable table_bookmark = 15;
  optional .TSWP.ObjectAttributeTable table_footnote = 16;
  optional .TSWP.ObjectAttributeTable table_section = 17;
  optional .TSWP.ObjectAttributeTable table_rubyfield = 18;
  optional .TSWP.StringAttributeTable table_language = 19;
  optional .TSWP.StringAttributeTable table_dictation = 20;
  optional .TSWP.ObjectAttributeTable table_insertion = 21;
  optional .TSWP.ObjectAttributeTable table_deletion = 22;
  optional .TSWP.ObjectAttributeTable table_highlight = 23;
  optional .TSWP.ParaDataAttributeTable table_para_bidi = 24;
  optional .TSWP.OverlappingFieldAttributeTable table_overlapping_highlight = 25;
  optional .TSWP.OverlappingFieldAttributeTable table_pencil_annotation = 26;
  optional .TSWP.ObjectAttributeTable table_tatechuyoko = 27;
  optional .TSWP.ObjectAttributeTable table_drop_cap_style = 28;
}```
