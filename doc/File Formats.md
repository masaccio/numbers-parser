#  Owner IDs

The `TSCE.CalculationEngineArchive` contains a number of owner ID mappings:

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
  * `owner_id_map`: a list of `TSCE.OwnerIDMapArchive` archives that maps a `internal_owner_id` number to a UUID, which is one of the various `conditional_style_formula_owner_id` or `formula_owner_id` UUIDs.

The (`TST.TableModelArchive`) for each table contains a number of UUIDs:

* `haunted_owner`: matches the `formula_owner_id` of one `TSCE.FormulaOwnerDependenciesArchive` which itself will contain a `base_owner_uid` which is referred to as the owner of many objects. This can be used to escalish a link between a specific table and a `base_owner_id`. The `base_owner_uid` is used in a number of ways:
  * Resolving cross-table formula references (see `_NumbersModel.node_to_ref()` for how this is used in `numbers-parser`)
  * Locating all cells that contain formulas (see `_NumbersModel.formula_cell_ranges()`)
  * Naming all merge cell ranges in a from/to notation (see `_NumbersModel.calculate_merge_cell_ranges()`)
* `merge_owner`:
