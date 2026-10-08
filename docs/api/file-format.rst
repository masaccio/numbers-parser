Apple Numbers file format
#########################

This document describes the on-disk format that ``numbers-parser`` reads and
writes, following the bytes from the document container down to individual
cells. The format is proprietary: Apple has not published a complete
specification, and the protobuf schemas in this repository are reverse
engineered definitions, not a promise that every Numbers release uses every
field in the same way.

.. note::

   Always check the sources linked to from here as this part of the documentation is created
   and maintained by AI. The protobuf examples are tested during the docs build, but the
   format descriptions may still be complete nonsense.

The description combines work from multiple sources, some of which are quite old
but remain valuable resources and have been invaluable in creating `numbers-parser`
and this document.

* Sean Patrick O'Brien `iWork file format reader <https://github.com/obriensp/iWorkFileFormat>`__.
* The SheetsJS Project's `documentation of the IWA file format <https://oss.sheetjs.com/notes/iwa/>`__.
* An earlier version of the `Stingray-Reader by S.Lott <https://github.com/slott56/Stingray-Reader/tree/V4_Archive/stingray>`__.
* The author's investigations into Numbers' protobufs.

The main implementation files are :src_root:`src/numbers_parser/iwork.py` (container
and encryption), :src_pkg:`iwafile.py` (IWA framing and protobuf segments),
:src_pkg:`containers.py` (object/file stores), :src_pkg:`model.py` (Numbers object model and
table storage), :src_pkg:`cell.py` (cell interpretation and serialization), and
:src_pkg:`constants.py` (format constants). The extracted proto2 schemas are in
:src_root:`src/protos`. In particular, consult :src_proto:`TSPArchiveMessages.proto`,
:src_proto:`TSPMessages.proto`, :src_proto:`TNArchives.proto`, :src_proto:`TSTArchives.proto`,
:src_proto:`TSCEArchives.proto`, :src_proto:`TSKArchives.proto`, :src_proto:`TSDArchives.proto`,
:src_proto:`TSSArchives.proto`, and :src_proto:`TSWPArchives.proto`. The generated Python
descriptors in :src_root:`src/numbers_parser/generated` are compiled from these
definitions.

Container and file layout
=========================

A ``.numbers`` document is not an XML spreadsheet. It is a ZIP-based package
containing metadata, resources such as images, and ``.iwa`` entries. The
single-file form is a ZIP archive. An iCloud/package form is a directory with
an ``Index.zip`` plus files and directories beside it. Some exported iCloud
documents wrap an ``Index.zip`` inside an outer ZIP. The reader handles nested
ZIP files and package directories.

A typical package has the following broad shape (the exact set of entries
varies by document and Numbers version)::

    Example.numbers/
        Index.zip
            Index/
                Document.iwa
                CalculationEngine.iwa
                DocumentStylesheet.iwa
                ...other IWA components...
        Data/
            ...image and other resource files...
        Metadata/
            Properties.plist
            BuildVersionHistory.plist
            DocumentIdentifier
        preview.jpg
        preview-web.jpg
        preview-micro.jpg

In a single-file document these entries are found inside the ZIP container
rather than as a directory tree. ``Metadata/Properties.plist`` contains the
``fileFormatVersion`` used by the reader to identify the Numbers document
version; ``BuildVersionHistory.plist`` is also expected for normal
documents.

The storage hierarchy is not the spreadsheet hierarchy. ZIP entries group
IWA data into components; the objects inside those IWA files refer to one
another by numeric identifiers. A sheet can refer to a table whose model is
stored in a different component, and a table model can refer to a tile, style,
string list, or formula list elsewhere. ``TSP.Reference`` is the generic
identifier-only edge in that graph (:src_proto:`TSPMessages.proto`); it does not carry
a reliable protobuf type, so the reader must resolve both the identifier and
the expected schema.

Password-protected files
------------------------

Encrypted documents use the ``.iwpv2`` verifier and ``.iwph`` hint entries.
The implementation checks a 104-byte verifier record with version 2 and
format version 1, derives a 16-byte key with PBKDF2-HMAC-SHA1, and verifies a
SHA-256 value in the decrypted verifier block. In the current implementation
the derivation uses the verifier's iteration count (the writer creates it
with 100,000 iterations).

An encrypted IWA blob has a 16-byte initialization vector at the front, a
block-aligned AES-128-CBC ciphertext, and 20 trailing bytes. After decryption,
PKCS#7 padding is removed and a 16-byte plaintext prefix is discarded before
the IWA stream is parsed. The encryption wrapper is separate from Snappy and
protobuf: first decrypt the IWA blob, then process its IWA chunks as below.
:src_root:`src/numbers_parser/iwork.py` contains the exact verifier and stream
handling logic.

IWA framing, Snappy, and protobuf
=================================

An ``.iwa`` file is an IWork Archive stream. It consists of contiguous
compressed chunks. Each chunk starts with a four-byte header: the first byte
is ``0`` and the remaining three bytes encode the compressed payload length
as a little-endian 24-bit integer. The header length is not included in that
length. The payload is Snappy-compressed data. This is the framing used in
the research sources, not a bare Snappy stream with a stream-identifier
chunk. The sources note that observed IWA streams omit the standard Snappy
stream identifier and CRC-32C checksums.

``IWAFile.from_buffer`` reads chunks until the file ends. The current writer
compresses slices of up to 65,536 uncompressed bytes and emits the same
four-byte framing header for each. ``IWACompressedChunk`` decompresses and
concatenates chunk data before parsing protobuf segments. A failed Snappy
decompression is allowed to pass through as uncompressed bytes as a
compatibility fallback; this is a reader tolerance, not another documented
cell-storage version. ``is_iwa_file`` checks the chunk framing and declared
lengths, while complete validation happens when the contained messages are
decoded.

The uncompressed chunk data is a sequence of archive segments. Each segment
has this layout::

    varint archive_info_length
    ArchiveInfo protobuf bytes
    payload 0 bytes (MessageInfo 0's length)
    payload 1 bytes (MessageInfo 1's length)
    ...

The length prefix is a protobuf base-128 varint and counts only the
``ArchiveInfo`` message. ``TSP.ArchiveInfo`` in
:src_proto:`TSPArchiveMessages.proto` supplies the segment's object ``identifier`` and
one or more ``message_infos``. Each ``TSP.MessageInfo`` gives a numeric
``type``, a version tuple, the number of payload bytes, and optional
``field_infos``, ``object_references``, and ``data_references``. Payloads
follow the header immediately, concatenated in ``message_infos`` order; the
``length`` on each record separates one payload from the next. Although the
schema allows multiple payloads in one segment, the historic research found
one to be usual. The implementation handles all listed message payloads.

The segment header is defined in :src_proto:`TSPArchiveMessages.proto`:

.. code-block:: protobuf

   message ArchiveInfo {
     optional uint64 identifier = 1;
     repeated .TSP.MessageInfo message_infos = 2;
     optional bool should_merge = 3;
   }

   message MessageInfo {
     required uint32 type = 1;
     repeated uint32 version = 2 [packed = true];
     required uint32 length = 3;
     repeated .TSP.FieldInfo field_infos = 4;
     repeated uint64 object_references = 5 [packed = true];
     repeated uint64 data_references = 6 [packed = true];
   }

Protobuf wire data is not self-describing. ``MessageInfo.type`` is resolved
through the Numbers/common registry extracted from the iWork applications;
the resulting maps are checked into :src_root:`src/numbers_parser/generated/mapping.py`.
The same numeric id can mean a different class in another iWork application.
The schema field types and message definitions live in the ``.proto`` files,
not in the bytes on disk. Likewise, a ``TSP.Reference`` identifies an object
but does not say what kind it points to.

The reference itself contains only the archive identifier
(:src_proto:`TSPMessages.proto`):

.. code-block:: protobuf

   message Reference {
     required uint64 identifier = 1;
     optional int32 deprecated_type = 2;
     optional bool deprecated_is_external = 3;
   }

``ArchiveInfo.identifier`` is the object's numeric identity within the
document. ``MessageInfo.object_references`` and ``data_references`` record
cross-object/resource references; the semantic links are also represented as
typed ``TSP.Reference`` and ``TSP.DataReference`` fields in payload
messages. The schema also permits ``should_merge`` and patch metadata
(``base_message_index``, ``diff_field_path`` and related fields); this is
protobuf-level object patching, not a separate cell encoding. The current
``ProtobufPatch`` support in :src_pkg:`iwafile.py` is intentionally limited.

Reader and writer path
======================

On read, ``IWork.open`` opens a ZIP file or package, reads metadata and
optional encryption state, and stores resources. ``ObjectStore`` identifies
IWA files, decompresses their chunks, parses archive segments using the
generated type registry, and indexes each typed protobuf by its
``ArchiveInfo.identifier``. ``_NumbersModel`` then follows references from
the document and tables, builds cached lookup tables for strings, styles,
formats and formulas, and extracts individual cell records from tile rows.
The public ``Document``, ``Sheet`` and ``Table`` classes project that model
into the higher-level API.

``ObjectStore`` keeps decoded protobuf archives in its object map, keyed by
``ArchiveInfo.identifier``. Its ``store_object`` handler records each typed
archive, and ``__getitem__`` exposes lookup by id; the model calls this
``self.objects`` and follows a protobuf reference with expressions such as
``self.objects[reference.identifier]``. The store's file cache separately
retains IWA blobs and other package files. In the public API,
:src_pkg:`document.py` delegates table, caption, header, and merge operations
to the model; :src_pkg:`cell.py` interprets cell records and asks the model
for styles and borders; :src_pkg:`model.py` resolves those structures through
the object store. For example, caption text follows the table-info caption
reference to a caption archive, then follows its owned-storage reference to
the text archive. These APIs expose interpreted values, not the raw protobuf
objects.

On write, updated cell values are encoded as v5 cell buffers; strings,
formats, styles and formulas are inserted into their relevant table lists.
The model rebuilds tile-row buffers and offsets, updates protobuf lengths
and object references, serializes segments, recompresses IWA chunks, and
writes the ZIP file or package. Non-IWA blobs (for example, image data) are
kept in the file store. :src_pkg:`model.py` and
:src_pkg:`iwafile.py` are the best references for following this
process end to end.

Document and component graph
============================

The top-level Numbers protobuf is ``TN.DocumentArchive``
(:src_proto:`TNArchives.proto`). Its ``sheets`` field points to ``TN.SheetArchive``
objects; it also points to a shared stylesheet, theme, sidebar order,
calculation engine and optional UI/custom-format metadata. The ``super``
field is a common ``TSA.DocumentArchive`` object. The sheet's
``drawable_infos`` references include tables and other sheet drawables.
``TN.SheetArchive.name`` is the user-visible sheet name, while its other
fields describe page orientation, headers and footers, printing, tab style,
visibility, and layout.

The root and sheet messages are defined in :src_proto:`TNArchives.proto`:

.. code-block:: protobuf

   message DocumentArchive {
     repeated .TSP.Reference sheets = 1;
     required .TSA.DocumentArchive super = 8;
     optional .TSP.Reference calculation_engine = 3 [deprecated = true];
     // ...
   }

   message SheetArchive {
     required string name = 1;
     repeated .TSP.Reference drawable_infos = 2;
     // ...
   }

These protobuf fields define the table traversal: the document's repeated
``sheets`` references resolve to sheet objects; each sheet's
``drawable_infos`` can resolve to a ``TableInfoArchive``; and its
``tableModel`` reference resolves to a ``TableModelArchive``. The model embeds
``base_data_store``, which contains the table data and tile/list references
described below. ``TableInfoArchive`` is drawable/view metadata, whereas
``TableModelArchive`` holds the table's persistent UUID, dimensions, name,
default styles, and data model. It also carries sorting, hidden and filtered
rows, merges, categories, pivot tables, and other table capabilities.
``TSP.Reference`` fields such as ``tableModel`` identify separate archive
objects resolved through ``ObjectStore``; ``base_data_store`` is an embedded
``TST.DataStore`` protobuf message, not another archive reference.
The data path continues from ``DataStore.tiles`` to ``TileStorage`` entries
that reference ``Tile`` archives; each tile's ``rowInfos`` contains
``TileRowInfo`` records with the row's cell-storage bytes and offsets.

The schema expresses that split directly
(:src_proto:`TSTArchives.proto`):

.. code-block:: protobuf

   message TableInfoArchive {
     required .TSD.DrawableArchive super = 1;
     required .TSP.Reference tableModel = 2;
     // ...
   }

   message TableModelArchive {
     required .TSP.Reference table_style = 3;
     required .TST.DataStore base_data_store = 4;
     required uint32 number_of_rows = 6;
     required uint32 number_of_columns = 7;
     required string table_name = 8;
     // ...
   }

The reader treats object identifiers 1 and 2 as the document and package
roots (``DOCUMENT_ID`` and ``PACKAGE_ID`` in :src_pkg:`constants.py`). The package
metadata describes component locators and versions, referenced external
components, object-UUID maps, and data resources. On load,
``ObjectStore`` indexes protobuf objects by identifier and also keeps the
original file blobs and the mapping back to their IWA entry. This separation
lets the higher-level ``Document``, ``Sheet`` and ``Table`` APIs work with a
graph of objects without exposing archive placement.

Table storage: headers, tiles, and shared lists
===============================================

The ``TST.DataStore`` inside a table model connects its data structures:

* ``rowHeaders`` and ``columnHeaders`` reference header storage. Header
  entries describe an index, size, hidden state, and optional cell/text
  styles. These records preserve row heights, column widths, hidden rows and
  columns, and header formatting independently of cell values.
* ``tiles`` is a ``TST.TileStorage`` index of tile ids and references.
  ``TST.Tile`` records contain per-row ``TileRowInfo`` entries. Tables are
  commonly tiled in blocks of up to 256 rows; the tile size is explicit in
  the schema and the parser's default is 256.
* ``rowTileTree`` and ``columnTileTree`` map row/column offsets to tile
  indexes. The reader also builds a row-to-storage mapping from the row
  header buckets because rows with no stored cells may have no tile-row
  storage entry.
* Shared tables hold strings, styles, formulas, formats, formula errors,
  rich-text payloads, conditional styles, comments, import warnings, control
  cell specifications, and merge information. The DataStore has separate
  references for these lists; optional lists may be absent.

The ``DataStore`` schema links the tile hierarchy and shared lists
(:src_proto:`TSTArchives.proto`):

.. code-block:: protobuf

   message DataStore {
     required .TST.HeaderStorage rowHeaders = 1;
     required .TSP.Reference columnHeaders = 2;
     required .TST.TileStorage tiles = 3;
     required .TSP.Reference stringTable = 4;
     required .TSP.Reference styleTable = 5;
     required .TSP.Reference formula_table = 6;
     required .TST.TableRBTree rowTileTree = 9;
     required .TSP.Reference format_table_pre_bnc = 11;
     optional .TSP.Reference rich_text_table = 17;
     optional .TSP.Reference control_cell_spec_table = 21;
     optional .TSP.Reference format_table = 22;
     // ...
   }

**DataStore headers**

Header metadata is kept in bucket objects rather than in the tile records.
``DataStore`` points to ``HeaderStorage`` for row buckets and to one
column-header bucket; each header stores its index, size, hidden state, cell
count, and optional cell/text style references
(:src_proto:`TSTArchives.proto`):

.. code-block:: protobuf

   message HeaderStorage {
     required uint32 bucketHashFunction = 1;
     repeated .TSP.Reference buckets = 2;
   }

   message HeaderStorageBucket {
     message Header {
       required uint32 index = 1;
       required float size = 2;
       required uint32 hidingState = 3;
       required uint32 numberOfCells = 4;
       optional .TSP.Reference cell_style = 5;
       optional .TSP.Reference text_style = 6;
     }
     required uint32 bucketHashFunction = 1;
     repeated .TST.HeaderStorageBucket.Header headers = 2;
   }

``_NumbersModel.row_storage_map`` follows the row-bucket references through
``self.objects`` and maps header indexes to corresponding row storage
positions. Empty rows can have header metadata but no tile row. The public
``Table`` header-count and dimension properties in :src_pkg:`document.py`
delegate to model methods, which read or update the table model archive;
header count fields and frozen-header flags are also defined on
``TableModelArchive``. Header style references feed the same style resolution
paths described below.

``TST.TileRowInfo`` contains a ``cell_count``, row index, cell storage bytes,
cell offsets, a storage version, and an optional ``has_wide_offsets`` flag.
The offsets map column positions to the corresponding cell record in the
row's storage byte buffer. A negative offset marks a column with no cell
record. A wide-offset row stores offsets in units of four bytes; the parser
scales those offsets before slicing the byte buffer. Empty columns between
stored cells therefore do not consume a cell record. ``get_storage_buffers_for_row``
uses the next nonnegative offset (or the end of the buffer) to delimit each
cell.

The schema keeps row metadata and the two storage layouts together
(:src_proto:`TSTArchives.proto`):

.. code-block:: protobuf

   message TileRowInfo {
     required uint32 tile_row_index = 1;
     required uint32 cell_count = 2;
     required bytes cell_storage_buffer_pre_bnc = 3;
     required bytes cell_offsets_pre_bnc = 4;
     optional uint32 storage_version = 5;
     optional bytes cell_storage_buffer = 6;
     optional bytes cell_offsets = 7;
     optional bool has_wide_offsets = 8;
   }

The shared values are represented by ``TST.TableDataList`` entries. Each
entry has a numeric ``key`` and ``refcount``; its payload depends on the list
type. The proto enumerates string, format, formula, style, formula-error,
custom-format, choice-list-format, rich-text, conditional-style, comment,
import-warning, and control-cell-spec lists. A string entry contains a
string; other entries may hold a protobuf value or a reference to another
archive. The cell's stored key indexes its table's list, not a document-wide
string table. ``DataLists`` in :src_pkg:`model.py` caches lookups, reuses existing
values, and allocates subsequent keys when writing new entries.

The protobuf declares the list kind separately from each entry's optional
value fields (:src_proto:`TSTArchives.proto`):

.. code-block:: protobuf

   message TableDataList {
     enum ListType {
       STRING = 1;
       FORMAT = 2;
       FORMULA = 3;
       STYLE = 4;
       FORMULA_ERROR = 5;
       RICH_TEXT_PAYLOAD = 8;
       CONTROL_CELL_SPEC = 12;
     }
     message ListEntry {
       required uint32 key = 1;
       required uint32 refcount = 2;
       optional string string = 3;
       optional .TSP.Reference reference = 4;
       optional .TSCE.FormulaArchive formula = 5;
       optional .TSK.FormatStructArchive format = 6;
       optional .TSP.Reference rich_text_payload = 9;
       optional .TST.CellSpecArchive cell_spec = 12;
     }
     required .TST.TableDataList.ListType listType = 1;
     required uint32 nextListID = 2;
     repeated .TST.TableDataList.ListEntry entries = 3;
   }

In the implementation, ``_NumbersModel`` creates a ``DataLists`` cache for
the table's string, style, formula, format, and control-spec references.
Lookups use both the table id and list key: ``table_string`` resolves a text
cell's string id, ``table_style`` resolves a style reference, and
``formula_ast`` follows formula entries. When writing, ``lookup_key`` reuses
an equal value or appends a ``ListEntry`` with the next key; reusing a value
increments its ``refcount``. Rich-text payloads and other list kinds are also
decoded through their specific model paths; not every ``ListType`` has a
``DataLists`` cache.

The plain string table and rich-text table are distinct. A plain text cell
looks up its string id in ``stringTable``. Rich text uses a rich-text payload
reference and ``TSWP.StorageArchive`` text runs, with character-indexed
attribute tables for paragraph and character styling. Hyperlinks are attached
to rich-text fragments (for example a ``TSWP.HyperlinkFieldArchive``), not
stored as an independent cell-level URL. This differs from formats that have
a hyperlink property on each cell.

The rich-text list is the DataStore's ``rich_text_table`` reference. A list
entry points to a rich-text payload, which points to its storage; the
storage's smart-field table points to hyperlink objects at character offsets.
The earlier ``TableDataList`` excerpt shows the entry's
``rich_text_payload = 9`` field. The following abbreviated protobuf excerpts
(``// ...`` marks omitted fields) describe that payload, its storage, and
hyperlink attributes. The same chain is discussed in
:src_root:`docs/api/sheetsjs.md`; schema definitions are in
:src_proto:`TSTArchives.proto` and :src_proto:`TSWPArchives.proto`:

.. code-block:: protobuf

   message RichTextPayloadArchive {
     required .TSP.Reference storage = 1;
   }

   message StorageArchive {
     repeated string text = 3;
     optional .TSWP.ObjectAttributeTable table_smartfield = 11;
     // ...
   }

   message ObjectAttributeTable {
     message ObjectAttribute {
       required uint32 character_index = 1;
       optional .TSP.Reference object = 2;
     }
     repeated .TSWP.ObjectAttributeTable.ObjectAttribute entries = 1;
   }

   message HyperlinkFieldArchive {
     optional string url_ref = 2;
   }

``_NumbersModel.table_rich_text`` follows those references through
``self.objects`` and recognizes ``HyperlinkFieldArchive`` objects. Each
``character_index`` begins a linked text run that ends at the next smart-field
entry or at the end of the text; the model returns the run text together with
``url_ref``. This is a rich-text extraction path, not a separate URL field on
the compact cell record.

Cell storage: v5 binary records
===============================

The cell protobuf ``TST.Cell`` describes the logical value and formatting,
but table tiles also store a compact binary record per cell. The current
``numbers-parser`` cell reader and writer support the v5 layout (called BNC
or post-BNC in the research sources). :src_pkg:`model.py` rejects a tile that is
not marked as saved in BNC form, and :src_pkg:`cell.py` rejects a cell record whose
first byte is not 5. The pre-v5 storage buffers are present in the schema for
compatibility, but are not decoded by this implementation and are omitted
here.

Every v5 cell record begins with a 12-byte header::

    byte 0       storage version (5)
    byte 1       storage cell type
    bytes 2-5    format-specific bytes
    bytes 6-7    auxiliary bits (preserved as extras by the reader)
    bytes 8-11   little-endian 32-bit presence mask
    bytes 12...  value fields and flagged fields, in mask order

The header and numeric values are little-endian. The presence mask is the
most important field: fields after byte 12 appear in ascending flag order
only when their bit is set. The reader starts at byte 12, reads the value
field indicated by bits 0-3, then reads optional identifiers and indexes for
bits 4 onward. Integer list keys and identifiers are signed 32-bit values in
the buffer. Date and double fields use 8-byte IEEE-754 values. The v5
Decimal128 payload is 16 bytes.

The field mask has these meanings in the v5 layout (:src_proto:`TSTArchives.proto`
and the SheetsJS research describe the storage map; :src_pkg:`cell.py` implements
the fields currently consumed):

.. list-table::
   :header-rows: 1
   :widths: 12 28 18 42

   * - Mask
     - Field
     - Size
     - Meaning
   * - ``0x000001``
     - Decimal128 value
     - 16 bytes
     - Decimal coefficient/exponent payload for a numeric or currency cell.
   * - ``0x000002``
     - Double value
     - 8 bytes
     - IEEE-754 value used for booleans and durations.
   * - ``0x000004``
     - Date/time value
     - 8 bytes
     - Seconds relative to the Numbers reference epoch.
   * - ``0x000008``
     - String id
     - 4 bytes
     - Key into the table's plain string list.
   * - ``0x000010``
     - Rich-text id
     - 4 bytes
     - Key/reference into the rich-text list.
   * - ``0x000020``
     - Cell-style id
     - 4 bytes
     - Key into the table style list.
   * - ``0x000040``
     - Text-style id
     - 4 bytes
     - Text style associated with the cell.
   * - ``0x000080``
     - Conditional-style id
     - 4 bytes
     - Conditional style reference.
   * - ``0x000100``
     - Conditional-style applied-rule id
     - 4 bytes
     - Identifies the applied conditional-format rule.
   * - ``0x000200``
     - Formula id
     - 4 bytes
     - Key into the table's formula list.
   * - ``0x000400``
     - Control-cell-spec id
     - 4 bytes
     - Key for a checkbox, rating, slider, stepper, or popup control.
   * - ``0x000800``
     - Formula-error id
     - 4 bytes
     - Index into formula error data.
   * - ``0x001000``
     - Suggested cell-format kind
     - 4 bytes
     - Format-kind metadata.
   * - ``0x002000``
     - Number-format id
     - 4 bytes
     - Index into number-format data.
   * - ``0x004000``
     - Currency-format id
     - 4 bytes
     - Index into currency-format data.
   * - ``0x008000``
     - Date-format id
     - 4 bytes
     - Index into date-format data.
   * - ``0x010000``
     - Duration-format id
     - 4 bytes
     - Index into duration-format data.
   * - ``0x020000``
     - Text-format id
     - 4 bytes
     - Index into text-format data.
   * - ``0x040000``
     - Boolean-format id
     - 4 bytes
     - Index into boolean-format data.
   * - ``0x080000``
     - Comment-storage id
     - 4 bytes
     - Comment metadata; the current cell decoder skips this field.
   * - ``0x100000``
     - Import-warning-set id
     - 4 bytes
     - Import/compatibility warnings; the current cell decoder skips this field.

The auxiliary word in bytes 6-7 is separate from the mask at bytes 8-11.
This is a little-endian 16-bit value. Numbers uses these to determine how a cell
should be rendered when it is formatted as automatic. The encoding is:

.. list-table::
   :header-rows: 1
   :widths: 14 48

   * - Bit
     - Hint
   * - ``0x0001`` 
     - Number format id is present
   * - ``0x0002`` 
     - Currency format id is present
   * - ``0x0004`` 
     - Duration format id is present
   * - ``0x0008`` 
     - Date format id is present
   * - ``0x0020`` 
     - Boolean format id is present
   * - ``0x0080`` 
     - String/text format id is present
   * - ``0x0800`` 
     - Currency-format-related hint
   * - ``0x8000`` 
     - Formula-related hint

These mappings are tentative: the writer's associations are based on a
decision-tree classifier trained on available Numbers documents, not a
complete specification. In particular, research has observed ``0x80`` in
byte 7 but has not independently established its meaning. The auxiliary word
is not the presence mask and is not used by the cell reader to locate fields.

Numbers can infer formats for automatically formatted input (for example,
interpreting ``3 3/4`` as a fraction or displaying ``3.14`` to two decimal
places). ``numbers-parser`` does not try to reproduce that inference or create
Automatic-formatted cells from ordinary values; it relies on native Python
types instead. When reading and writing cells, it preserves the auxiliary word
from storage bytes 6-7 verbatim, including hints on Automatic-formatted cells.

Cell type byte and value interpretation
---------------------------------------

The byte at offset 1 is the storage cell type. Do not confuse it with the
``TST.CellValueType`` enum used in protobuf messages, or the public
``numbers_parser.constants.CellType`` enum: those are separate namespaces.
The proto storage enum in :src_proto:`TSTArchives.proto` defines generic/empty,
span, number, text, formula, date, boolean, duration, formula error, and
automatic (rich-text) kinds. The implementation also recognizes storage
type 10 as currency. The type is used together with the mask to interpret
the value fields:

.. list-table::
   :header-rows: 1
   :widths: 14 30 56

   * - Stored type
     - Logical value
     - Decoding
   * - 0
     - Empty
     - No value payload; an empty storage record is 12 bytes.
   * - 1
     - Span
     - Schema value for a span cell; not a normal logical cell value.
   * - 2
     - Number
     - Decimal128 payload is decoded to the Python numeric value.
   * - 3
     - Text
     - String id resolves through the table's shared plain-string list.
   * - 4
     - Formula
     - Schema value for a formula cell; in ordinary table data the formula
       id accompanies the cell's cached value type.
   * - 5
     - Date/time
     - Seconds from 2001-01-01 are added to the reference epoch.
   * - 6
     - Boolean
     - The double value is true when greater than zero, otherwise false.
   * - 7
     - Duration
     - The double value is interpreted as seconds and converted to a duration.
   * - 8
     - Formula error
     - Error id resolves through the table's formula-error data.
   * - 9
     - Automatic / rich text
     - Rich-text id resolves through the table's rich-text list.
   * - 10
     - Currency
     - Decimal128 value interpreted as currency; type 10 is a parser constant.

:src_proto:`TSTArchives.proto` uses a different numeric enum for
``CellValueType``: empty 0, number 1, string 2, provided 3, date 4, boolean
5, duration 6, error 7, rich text 8, and currency 9. It also has a storage
``CellType`` enum where number is 2, text 3, formula 4, date 5, boolean 6,
duration 7, error 8, and automatic 9. The raw v5 storage byte follows the
storage-type convention, including the extra currency type handled in code.

Numbers, dates and durations
----------------------------

Numeric and currency cells use a 128-bit decimal payload, not an IEEE double.
``_unpack_decimal128`` in :src_pkg:`cell.py` extracts a signed integer mantissa and
biased base-10 exponent using ``DECIMAL128_BIAS`` from :src_pkg:`constants.py`;
the API currently exposes the resulting number as a Python numeric value.
When writing, ``_pack_decimal128`` builds the 16-byte representation and
numbers are rounded to the library's configured significant-digit limit
before serialization. The separate double bit is used for boolean values
and durations.

Numbers dates are represented as seconds relative to 2001-01-01, defined as
``EPOCH`` in :src_pkg:`constants.py`. Date records carry an eight-byte double
seconds value. Durations also carry seconds as a double, but are elapsed
intervals, not dates relative to the epoch. The format id controls how these
values are displayed. Number, currency, date, duration, text, and boolean
format ids are separate optional fields in v5 and resolve through the
table's format list.

The ``CellType`` enum exported by :src_pkg:`constants.py` is an API-level
classification, not a transcription of the storage byte:
``EMPTY=1``, ``NUMBER=2``, ``TEXT=3``, ``DATE=4``, ``BOOL=5``,
``DURATION=6``, ``ERROR=7``, ``RICH_TEXT=8``, ``CURRENCY=101``, and
``MERGED=102``. The type ids come from the library's normalized model; in
particular, its currency and merged values are not literal storage type ids.

The writer emits 12-byte headers and v5 type/value payloads for supported
empty, number, currency, text, date, boolean, duration, rich-text and error
cells. It adds each optional field after the base value in mask order. Merged
cells are represented by merge metadata rather than a normal cell value
record.

Formulas, formats, styles and controls
======================================

A cell's cached result lives in the v5 cell record. A formula is stored
separately: the formula-id field indexes the table's formula list, whose
entry contains a ``TSCE.FormulaArchive`` (see :src_proto:`TSCEArchives.proto`). A
formula archive contains an abstract syntax tree, host row/column and table
identity, translation flags, and related formula metadata. The AST is made
of typed nodes for literals, operators, functions, and cell/range
references. References can be local or cross-table and may use row/column
UUIDs; these stable identities matter when rows or columns move. The parser
uses current cell values rather than recalculating formula expressions.
:src_pkg:`cell.py` resolves a formula id and its coordinates, while :src_pkg:`formula.py`
handles formula rendering and references.

Formula AST decoding
--------------------

The table's ``formula_table`` is a ``TableDataList`` of ``FORMULA`` entries.
Each cell's formula id selects an entry by key, and that entry embeds a
``TSCE.FormulaArchive``. Its ``AST_node_array`` is an ordered sequence of
typed nodes rather than formula text:

.. code-block:: protobuf

   message FormulaArchive {
     required .TSCE.ASTNodeArrayArchive AST_node_array = 1;
     optional uint32 host_column = 2;
     optional uint32 host_row = 3;
     optional .TSP.UUID host_table_uid = 7;
     optional .TSP.UUID host_column_uid = 8;
     optional .TSP.UUID host_row_uid = 9;
   }

   message ASTNodeArrayArchive {
     enum ASTNodeType {
       ADDITION_NODE = 1;
       FUNCTION_NODE = 16;
       NUMBER_NODE = 17;
       STRING_NODE = 19;
       LOCAL_CELL_REFERENCE_NODE = 27;
       CROSS_TABLE_CELL_REFERENCE_NODE = 28;
       // Other operators, values, and reference node types are defined here.
     }
     message ASTNodeArchive {
       required .TSCE.ASTNodeArrayArchive.ASTNodeType AST_node_type = 1;
       optional uint32 AST_function_node_index = 2;
       optional uint32 AST_function_node_numArgs = 3;
       optional double AST_number_node_number = 4;
       optional string AST_string_node_string = 6;
       optional .TSCE.ASTNodeArrayArchive.ASTLocalCellReferenceNodeArchive
         AST_local_cell_reference_node_reference = 15;
       optional .TSCE.ASTNodeArrayArchive.ASTCrossTableReferenceExtraInfoArchive
         AST_cross_table_reference_extra_info = 28;
     }
     repeated .TSCE.ASTNodeArrayArchive.ASTNodeArchive AST_node = 1;
   }

``_NumbersModel.formula_ast`` follows the table model's ``base_data_store`` to
the formula list and indexes the embedded AST node sequence by ``entry.key``.
``Cell.formula`` passes its formula id and row/column to ``TableFormulas``.
That renderer visits the stored nodes and feeds operands and operators to a
stack: literal nodes push their values, while operators and functions pop
their arguments and push a rendered expression. Function ids are translated
through the generated function map; unsupported node or function ids produce
``UnsupportedWarning`` rather than being evaluated. A missing formula key is
also reported as unsupported. The result is a formula string for the API,
not a recalculated cell value.

Cell and range reference nodes are resolved by ``_NumbersModel.node_to_ref``.
Local row and column coordinates combine their absolute/relative flags with
the formula cell's location. Cross-table nodes carry a table UUID; the model
maps that UUID back to a table before building the reference. UUID-based
coordinate and range nodes preserve stable row/column identities and sticky
absolute-reference flags where those node forms are present.

Formatting is also indirect. The v5 cell record can carry style ids,
conditional-style data, format ids and a control-spec id. Table list entries
map keys to the actual style, format, formula, or control data. A cell style
can reference ``TST.CellStyleArchive`` and its cell properties; text
properties use the TSWP text/style messages. The table model supplies default
body, header-row, header-column and footer styles, with per-cell entries
providing overrides. The document-level stylesheet and theme provide shared
style definitions and theme context.

Style information is a graph of style archives. The table's ``styleTable``
maps a cell's stored style key to an archive reference; cell and text styles
then inherit unset properties from their parent style
(:src_proto:`TSSArchives.proto`, :src_proto:`TSTArchives.proto`, and
:src_proto:`TSTStylePropertyArchiving.proto`):

.. code-block:: protobuf

   message StyleArchive {
     optional string name = 1;
     optional string style_identifier = 2;
     optional .TSP.Reference parent = 3;
     optional .TSP.Reference stylesheet = 5;
   }

   message CellStyleArchive {
     required .TSS.StyleArchive super = 1;
     optional .TST.CellStylePropertiesArchive cell_properties = 11;
   }

   message CellStylePropertiesArchive {
     optional .TSD.FillArchive cell_fill = 1;
     optional bool text_wrap = 3;
     optional .TSD.StrokeArchive top_stroke = 10;
     optional .TSD.StrokeArchive right_stroke = 11;
     optional .TSD.StrokeArchive bottom_stroke = 12;
     optional .TSD.StrokeArchive left_stroke = 13;
   }

``_NumbersModel.table_style`` follows the style-list entry's reference through
``self.objects``. ``cell_text_style`` chooses a cell-specific text-style key,
or falls back to the table's header-row, header-column, footer-row, or body
default. Its ``char_property``, ``para_property`` and ``cell_property``
helpers retrieve explicitly set values and otherwise follow the style's
parent. :src_pkg:`cell.py` exposes the resolved properties through the cell's
lazy ``style`` object.

Cell borders also occur in ``CellStylePropertiesArchive``, but the grid
strokes reported by ``Cell.border`` are extracted from a table stroke sidecar.
The model follows the table's sidecar and layer references, then converts
ordered stroke runs to border values (:src_proto:`TSTArchives.proto`):

.. code-block:: protobuf

   message TableModelArchive {
     optional .TSP.Reference stroke_sidecar = 49;
     // ...
   }

   message StrokeSidecarArchive {
     repeated .TSP.Reference left_column_stroke_layers = 4;
     repeated .TSP.Reference right_column_stroke_layers = 5;
     repeated .TSP.Reference top_row_stroke_layers = 6;
     repeated .TSP.Reference bottom_row_stroke_layers = 7;
   }

   message StrokeLayerArchive {
     message StrokeRunArchive {
       optional int32 origin = 1;
       optional uint32 length = 2;
       optional .TSD.StrokeArchive stroke = 3;
       optional uint32 order = 4;
     }
     optional uint32 row_column_index = 1;
     repeated .TST.StrokeLayerArchive.StrokeRunArchive stroke_runs = 2;
   }

Each run's orientation comes from its containing sidecar list; ``origin`` and
``length`` locate the run along a row or column, while the ``TSD.StrokeArchive``
provides the visual stroke. ``extract_strokes`` sorts runs by order, resolves
their layers through ``self.objects``, and sets matching edges on adjacent
cells. Setting borders reverses this process by consolidating equal edges into
stroke runs. The separate ``CellBorderArchive`` schema is also present for
logical cell messages, but is not the source of the table-grid sidecar strokes
used by the current border accessor.

:src_pkg:`constants.py` names the API and format enumerations for standard formats
(base, currency, date/time, fraction, number, percentage, scientific, text,
checkbox, rating, duration and custom formats), interactive controls (popup,
rating, slider, stepper and tickbox), negative-number style, fraction
accuracy, and duration units. These enum values describe properties and
formats, not the type byte or field mask in the cell buffer.

Controls have their own ``TST.CellSpecArchive``. Its interaction kind and
optional range limits, increment, or popup-model reference provide the
behavior behind the cell's control-spec id. ``FormattingType`` and related
maps in :src_pkg:`constants.py` connect API formatting choices to protobuf format
archives. Custom formats also have a document-level custom-format list and
UUIDs, while table entries refer to the corresponding custom format.

The numeric ``FormatType`` codes in :src_pkg:`constants.py` are: boolean 1, decimal
256, currency 257, percent 258, scientific 259, text 260, date 261, fraction
262, checkbox 263, rating 267, duration 268, base 269, custom number 270,
custom text 271, custom date 272, and custom currency 274. These protobuf
format kinds are distinct from ``FormattingType`` (the library's operation
categories) and ``CustomFormattingType`` (number 101, datetime 102, text
103). ``ControlFormattingType`` covers base 1, currency 2, fraction 4,
number 5, percentage 6, and scientific 7.

Other values used while interpreting formats and controls include:

* ``DurationStyle``: compact 0, short 1, long 2.
* ``DurationUnits``: none 0, week 1, day 2, hour 4, minute 8, second 16,
  millisecond 32 (the unit values are bit flags).
* ``NegativeNumberStyle``: minus 0, red 1, parentheses 2,
  red-and-parentheses 3.
* ``FractionAccuracy``: up to three, two, or one denominator digits use
  ``0xFFFFFFFD``, ``0xFFFFFFFE``, and ``0xFFFFFFFF``; fixed denominators
  include halves 2, quarters 4, eighths 8, sixteenths 16, tenths 10, and
  hundredths 100.
* ``PaddingType``: no padding 0, zeros 1, spaces 2.
* ``CellInteractionType``: value editing 0, formula editing 1, stock 2,
  category summary 3, stepper 4, slider 5, rating 6, popup 7, toggle 8.
* ``NumberFormatConditionType``: none -1, equal 0, less-than 1,
  less-than-or-equal 2, greater-than 3, greater-than-or-equal 4.

The separate ``constants.CellValueType`` is an internal value-kind enum
(``NIL_TYPE=1``, boolean 2, date 3, number 4, string 5), used when constructing
control values such as popup-menu items. It must not be confused with the
proto ``TST.CellValueType`` enum. The date/time formatting token table in
:src_pkg:`constants.py` maps Numbers tokens for years, months, days, weekday names,
week numbers, hours, minutes, seconds, fractional seconds, AM/PM, era, and
quarter to the display implementation; the supported spellings are listed in
``DATETIME_FIELD_MAP``.

Merges and stable coordinates
=============================

Merged cells are not represented by putting a value in every covered cell.
The table has merge ranges and/or merge-owner/formula metadata; the parser
combines these into a merge anchor and references for the cells covered by
the merge. ``MergeRegionMapArchive`` stores ranges, while
``MergeOperationArchive`` and merge-owner/formula structures are also
defined in :src_proto:`TSTArchives.proto`. :src_pkg:`model.py` recognizes the representations
it encounters, and :src_pkg:`cell.py` exposes a merged anchor and covered-cell
references to the higher-level API.

Numbers has stored merges in three different places, and ``_NumbersModel.merge_cells``
tries them in this order, stopping at the first one that yields any ranges:

1. **Merge-owner formula store** (current Numbers). ``TableModelArchive.merge_owner``
   holds a ``TST.FormulaStoreArchive``. Each merge is a formula whose first AST node
   is a ``COLON_TRACT_NODE``; the rows and columns are ranges in the node's
   ``AST_colon_tract``.
2. **Merge-owner dependency archives**. A ``TSCE.FormulaOwnerDependenciesArchive``
   with ``owner_kind`` of ``MERGE_OWNER`` records each merge as a back-dependency
   whose ``internal_range_reference`` names the owning table and the range.
3. **Merge region map** (legacy fallback). ``DataStore.merge_region_map`` points to a
   ``MergeRegionMapArchive`` of packed cell ranges.

The merge-owner formula store is defined in :src_proto:`TSTArchives.proto`:

.. code-block:: protobuf

   message TableModelArchive {
     optional .TST.MergeOwnerArchive merge_owner = 47;
     // ...
   }

   message MergeOwnerArchive {
     required .TSP.CFUUIDArchive owner_id = 1;
     optional .TST.FormulaStoreArchive formula_store = 2;
   }

   message FormulaStoreArchive {
     message FormulaStorePair {
       required uint32 formula_index = 1;
       required .TSCE.FormulaArchive formula = 2;
     }
     required uint32 next_formula_index = 2;
     repeated .TST.FormulaStoreArchive.FormulaStorePair formulas = 3;
   }

The dependency-archive form is defined in :src_proto:`TSCEArchives.proto`:

.. code-block:: protobuf

   message RangeBackDependencyArchive {
     required uint32 cell_coord_row = 1;
     required uint32 cell_coord_column = 2;
     optional .TSCE.RangeReferenceArchive range_reference = 3;
     optional .TSCE.InternalRangeReferenceArchive internal_range_reference = 4;
   }

   message RangeDependenciesArchive {
     repeated .TSCE.RangeBackDependencyArchive back_dependency = 2;
   }

The legacy region map, which only matters for older documents, is:

.. code-block:: protobuf

   message DataStore {
     optional .TSP.Reference merge_region_map = 13;
     // ...
   }

   message MergeRegionMapArchive {
     repeated .TST.CellRange cell_range = 1;
   }

   message CellRange {
     required .TST.CellID origin = 1;
     required .TST.TableSize size = 2;
   }

``calculate_merges_using_region_map`` splits ``packedData`` into column
(``>> 16``) and row (``& 0xFFFF``) and computes inclusive end coordinates. All three
paths call ``add_merge_range``, which builds the map of anchor and covered cells;
``Table.merge_ranges`` turns anchors into A1 ranges, and :src_pkg:`cell.py` returns
``MergedCell`` objects for covered positions.

UUIDs, owners, and table relationships
======================================

These protobuf structures extend the formula, merge, and stable-coordinate
descriptions above. Some behaviors described here are observations from sample
documents, not requirements for every Numbers version.
The relevant definitions are in :src_proto:`TSCEArchives.proto`, :src_proto:`TSTArchives.proto`, and
:src_proto:`TSPMessages.proto`.

Archive object identifiers, internal owner ids, and UUIDs are distinct
identifiers. A ``TSP.Reference`` points to an archive object by its numeric
``identifier``; an owner id map relates an internal integer to a UUID. Do not
use one as a substitute for another.

The schemas use two representations for 128-bit UUID values
(:src_proto:`TSPMessages.proto`). ``UUID`` stores upper and lower 64-bit words;
``CFUUIDArchive`` stores four 32-bit words (or optional raw bytes):

.. code-block:: protobuf

   message UUID {
     required uint64 lower = 1;
     required uint64 upper = 2;
   }

   message CFUUIDArchive {
     optional bytes uuid_bytes = 1;
     optional uint32 uuid_w0 = 2;
     optional uint32 uuid_w1 = 3;
     optional uint32 uuid_w2 = 4;
     optional uint32 uuid_w3 = 5;
   }

``NumbersUUID`` converts ``UUID`` values from their upper/lower words and
``CFUUIDArchive`` values from their four 32-bit word fields to one 128-bit
integer; it emits the corresponding word representation when writing.
Although the schema also permits ``CFUUIDArchive.uuid_bytes``, this conversion
uses the word fields. ``uuid_to_hex`` normalizes these values for comparisons
and maps, which is important because owner-map and formula-owner messages use
different protobuf UUID types.

Formula owners and table identities
-----------------------------------

The document's ``TSCE.CalculationEngineArchive`` contains a
``DependencyTrackerArchive``. Its ``formula_owner_dependencies`` references
``FormulaOwnerDependenciesArchive`` records, while ``formula_owner_info``
contains cell and range dependency information. The tracker's
``owner_id_map`` maps each ``internal_owner_id`` to a UUID. This map is used to
resolve dependency records that refer to an owner by integer. The
``owner_kind`` field is a numeric value; the parser names the observed table
model (1), merge owner (5), and haunted owner (35) kinds in
:src_pkg:`constants.py` (``OwnerKind``).
These values are not declared as an enum alongside the protobuf field.

The owner map relates 32-bit internal ids to UUIDs, while the owner
dependency records carry UUIDs and dependency metadata
(:src_proto:`TSCEArchives.proto`):

.. code-block:: protobuf

   message OwnerIDMapArchive {
     message OwnerIDMapArchiveEntry {
       required uint32 internal_owner_id = 1;
       required .TSP.CFUUIDArchive owner_id = 2;
     }
     repeated .TSCE.OwnerIDMapArchive.OwnerIDMapArchiveEntry map_entry = 1;
   }

   message FormulaOwnerDependenciesArchive {
     required .TSP.UUID formula_owner_uid = 1;
     required uint32 internal_formula_owner_id = 2;
     optional uint32 owner_kind = 3 [default = 0];
     optional .TSCE.RangeDependenciesArchive range_dependencies = 5;
     optional .TSP.Reference formula_owner = 11;
     optional .TSP.UUID base_owner_uid = 12;
     // ...
   }

   message DependencyTrackerArchive {
     optional .TSCE.OwnerIDMapArchive owner_id_map = 3;
     repeated .TSP.Reference formula_owner_dependencies = 6;
   }

The ``owner_id_map`` accessor reads those map entries into a dictionary from
internal owner id to normalized UUID hex. The table UUID mapping is separate:
``calculate_table_uuid_map`` finds haunted-owner dependency archives, maps each
``formula_owner_uid`` to its ``base_owner_uid``, and matches the table model's
``haunted_owner.owner_uid`` against that formula-owner UUID. The resulting
base-owner UUID is the stable table identity used to match cross-table formula
references. When no haunted-owner records exist (as in some older documents),
the model leaves this table mapping empty.

The table model's schema supplies the haunted-owner UUID that participates in
this match (:src_proto:`TSTArchives.proto` and :src_proto:`TSCEArchives.proto`):

.. code-block:: protobuf

   message TableModelArchive {
     optional .TSCE.HauntedOwnerArchive haunted_owner = 84;
     // ...
   }

   message HauntedOwnerArchive {
     required .TSP.UUID owner_uid = 1;
   }

In observed files, some formula-owner UUIDs share their upper 112 bits while
their lower 16 bits vary with the formula id. This pattern is an observation,
not a UUID-generation rule.

For a table, the ``TableModelArchive.haunted_owner.owner_uid`` has been
observed to match the ``formula_owner_uid`` of a dependency archive whose
``owner_kind`` is 35. That archive's ``base_owner_uid`` supplies the stable
UUID used to associate dependencies with the table. In the implementation,
``calculate_table_uuid_map()`` in :src_pkg:`model.py`
builds this mapping; documents without these dependency archives can lack it.
A separate ``owner_kind=1`` dependency archive represents the table model and
its ``formula_owner`` reference can point to the table's
``TableInfoArchive``. The table's optional
``conditional_style_formula_owner_id`` is another UUID field and should not
be confused with either owner mapping.

Formula-cell locations can also be recovered from each
``FormulaOwnerInfoArchive.cell_dependencies.cell_record``: records include
coordinates and a ``contains_a_formula`` flag. These records identify
formula-bearing cells; the cell buffer itself holds each formula's cached
result, while the formula list holds the expression.

The example documents in ``Numbers.md`` (on the ``feat/format-docs`` branch) show table-model
dependency archives with spanning ranges for both the whole table and its body. They also show a
variation in ``tiled_cell_dependencies``: the first example had no tile
reference, while later examples referred to ``CellRecordTileArchive`` records.
The referenced tiles carry an ``internal_owner_id`` and tile row/column
origins. Treat this as observed variation, not a rule that all later tables
must have tiles or that all first tables omit them.

Merge and formula ranges
------------------------

One range representation is ``RangePrecedentsTileArchive``. Its
``from_to_range`` entries pair a starting coordinate with a rectangle, and
``to_owner_id`` identifies the destination owner. Resolve that integer using
the calculation engine's ``owner_id_map`` before relating the rectangle to a
table UUID. Separately, merge-owner ``FormulaOwnerDependenciesArchive``
records carry ``range_dependencies``; the parser resolves their internal owner
ids and keeps ranges whose base-owner UUID matches the table.

The schemas describe both the formula-owner dependency form and the table's
merge-owner formula store (:src_proto:`TSCEArchives.proto` and
:src_proto:`TSTArchives.proto`):

.. code-block:: protobuf

   message RangePrecedentsTileArchive {
     message FromToRangeArchive {
       required .TSCE.CellCoordinateArchive from_coord = 1;
       required .TSCE.CellRectArchive refers_to_rect = 2;
     }
     required uint32 to_owner_id = 1;
     repeated .TSCE.RangePrecedentsTileArchive.FromToRangeArchive from_to_range = 2;
   }

   message RangeDependenciesArchive {
     repeated .TSCE.RangeBackDependencyArchive back_dependency = 2;
   }

   message RangeBackDependencyArchive {
     required uint32 cell_coord_row = 1;
     required uint32 cell_coord_column = 2;
     optional .TSCE.RangeReferenceArchive range_reference = 3;
     optional .TSCE.InternalRangeReferenceArchive internal_range_reference = 4;
   }

   message InternalRangeReferenceArchive {
     required uint32 owner_id = 1;
     required .TSCE.RangeCoordinateArchive range = 2;
   }

   message MergeOwnerArchive {
     required .TSP.CFUUIDArchive owner_id = 1;
     optional .TST.FormulaStoreArchive formula_store = 2;
   }

   message FormulaStoreArchive {
     message FormulaStorePair {
       required uint32 formula_index = 1;
       required .TSCE.FormulaArchive formula = 2;
     }
     required uint32 next_formula_index = 2;
     repeated .TST.FormulaStoreArchive.FormulaStorePair formulas = 3;
   }

The earlier ``Merges and stable coordinates`` section shows the
``DataStore.merge_region_map`` reference and its ``CellRange`` list. The table
model's merge-owner reference and the packed coordinate fields are defined in
:src_proto:`TSTArchives.proto`:

.. code-block:: protobuf

   message TableModelArchive {
     optional .TST.MergeOwnerArchive merge_owner = 47;
     // ...
   }

``CellRange.origin`` and ``CellRange.size`` use ``CellID.packedData`` and
``TableSize.packedData``. ``model.py`` unpacks their high and low 16-bit halves
to obtain column/row starts and counts, then computes inclusive ends.

.. code-block:: protobuf

   message CellID {
     required fixed32 packedData = 1;
     // ...
   }

   message TableSize {
     required fixed32 packedData = 1;
     // ...
   }

In :src_pkg:`model.py`, merge extraction tries the merge owner's formula store
first, then formula-owner dependency archives, then ``merge_region_map``.
``add_merge_range`` records each anchor and covered cell; the public
``Table.merge_ranges`` property in :src_pkg:`document.py` converts anchors to
A1-style ranges. The packed map representation is also used when writing
updated merges.

Header names and row/column UUIDs
---------------------------------

The calculation engine can reference a ``TST.HeaderNameMgrArchive``. Its
``per_tables`` entries associate a table UUID and precedent coordinate with
header row and column UUIDs. A ``HeaderNameMgrTileArchive`` stores name
fragments, each with a precedent cell coordinate and optional UUID-based
references to cells using that fragment. The row and column UUIDs can be
correlated with ``ColumnRowUIDMapArchive.sorted_row_uids`` and
``sorted_column_uids``. Research has also observed ``nrm_owner_uid`` matching
a formula-owner UUID, but its role in that mapping remains unresolved.

The header-name manager associates per-table header identities with coordinate
context, and stores name-fragment tiles by reference. Each fragment includes
its text and precedent coordinate; its optional UUID reference set identifies
cells using that fragment. The table's UID map contains sorted UUID arrays and
index translation arrays (:src_proto:`TSTArchives.proto`):

.. code-block:: protobuf

   message HeaderNameMgrArchive {
     message PerTableArchive {
       required .TSP.UUID table_uid = 1;
       required .TSCE.CellCoordinateArchive per_table_precedent = 2;
       repeated .TSP.UUID header_row_uids = 5;
       repeated .TSP.UUID header_column_uids = 6;
       // ...
     }
     required .TSP.UUID owner_uid = 1;
     optional .TSP.UUID nrm_owner_uid = 2;
     repeated .TST.HeaderNameMgrArchive.PerTableArchive per_tables = 3;
     repeated .TSP.Reference name_frag_tiles = 4;
     // ...
   }

   message HeaderNameMgrTileArchive {
     message NameFragmentArchive {
       required string name_fragment = 1;
       required .TSCE.CellCoordinateArchive name_precedent = 2;
       optional .TSCE.UidCellRefSetArchive uses_of_name_fragment = 3;
     }
     required string first_fragment = 1;
     required string last_fragment = 2;
     repeated .TST.HeaderNameMgrTileArchive.NameFragmentArchive name_frag_entries = 3;
   }

   message ColumnRowUIDMapArchive {
     repeated .TSP.UUID sorted_column_uids = 1;
     repeated uint32 column_index_for_uid = 2;
     repeated uint32 column_uid_for_index = 3;
     repeated .TSP.UUID sorted_row_uids = 4;
     repeated uint32 row_index_for_uid = 5;
     repeated uint32 row_uid_for_index = 6;
   }

The ``*_uid_for_index`` arrays translate table positions to entries in the
sorted UUID arrays; the ``*_index_for_uid`` arrays provide the reverse lookup.
The model uses ``ColumnRowUIDMapArchive.row_uid_for_index`` while reconstructing
category row relationships. Header names and row/column UUIDs therefore form
stable identities alongside coordinate-based references, rather than replacing
the archive identifiers used by ``TSP.Reference``.

Captions and text storage
-------------------------

Caption text is stored through a nested shape/storage structure rather than
directly in a caption record. ``TSA.CaptionInfoArchive`` contains a
``TSWP.ShapeInfoArchive``; its ``owned_storage`` reference points to a
``TSWP.StorageArchive`` whose ``text`` field contains the text. The shape
messages also carry drawable and placement information. The
:src_proto:`TSAArchives.proto`, :src_proto:`TSWPArchives.proto`, and
:src_proto:`TSDArchives.proto` definitions show the required ``super`` chain
and storage fields. The caption's ``owned_storage`` reference is what connects
the caption metadata to its text content.
In the sample files examined for this guide, the storage archive was reached
through this ``owned_storage`` reference. ``Metadata.json`` exposed the
object-UUID-map listing, but not a direct caption-storage reference.

The following abbreviated protobuf snippets (``// ...`` marks omitted fields)
show the relevant messages:

.. code-block:: protobuf

   message CaptionInfoArchive {
     required .TSWP.ShapeInfoArchive super = 1;
     optional .TSP.Reference placement = 2;
     optional .TSD.CaptionOrTitleKind childInfoKind = 3;
     // ...
   }

   message ShapeInfoArchive {
     required .TSD.ShapeArchive super = 1;
     optional .TSP.Reference owned_storage = 4;
     optional bool is_text_box = 6;
     // ...
   }

   message StorageArchive {
     repeated string text = 3;
     // ...
   }

In :src_pkg:`model.py`, ``caption_text`` starts from the table's
``TableInfoArchive`` and resolves its caption through ``self.objects``. It
then reads or replaces the first string in the owned ``StorageArchive.text``.
If the document has a stand-in caption and a caption is assigned,
``create_caption_archive`` creates the placement, caption, and storage
archives and connects their references. The public ``Table.caption`` and
``Table.caption_enabled`` properties in :src_pkg:`document.py` delegate to
these model operations; visibility is represented by the drawable's
``caption_hidden`` flag.

Remaining concepts to document
==============================

The following lookups are made through ``ObjectStore`` (``self.objects`` in
:src_pkg:`model.py`) but are not yet described above. Each entry names the
protobuf fields followed and the code that follows them.

Object store queries
--------------------

* ``ObjectStore.find_refs`` (:src_pkg:`containers.py`) returns the identifiers of
  every archive whose Python class name matches a string. It is used through
  ``_NumbersModel.find_refs`` to locate ``TableInfoArchive``,
  ``StylesheetArchive``, ``ParagraphStyleArchive``,
  ``FormulaOwnerDependenciesArchive``, ``CalculationEngineArchive`` and
  ``GroupNodeArchive`` objects. The results are not cached because tables and
  sheets can be added at run time. The relationship between a class name,
  its ``.proto`` message and its ``ArchiveInfo`` type id is not described.
* ``ObjectStore.remove_unreferenced_objects`` (:src_pkg:`containers.py`) walks
  every ``TSP.Reference`` in every archive to find objects that can be dropped
  on save. The rules for which fields count as references are not documented.
* ``ObjectStore.new_message_id`` and ``create_object_from_dict``
  (:src_pkg:`containers.py`) allocate identifiers and register new archives in
  a component. How a new object is assigned to an ``.iwa`` file and recorded
  in the package component list (``PACKAGE_ID`` in ``components``) is only
  partly described.
* ``find_extension`` (:src_pkg:`iwafile.py`) reads protobuf extension fields
  such as ``paragraph_style_presets`` from ``TSS.ThemeArchive.super``
  (:src_proto:`TSSArchives.proto`). Protobuf extensions are not covered.

Stylesheet and theme
--------------------

* ``TN.DocumentArchive.stylesheet`` and ``theme`` references resolve to
  ``TSS.StylesheetArchive`` and ``TSS.ThemeArchive``
  (:src_proto:`TSSArchives.proto`). ``StylesheetArchive.styles`` and
  ``identifier_to_style_map`` are appended to and searched by name when
  paragraph and cell styles are created (``add_paragraph_style``,
  ``add_cell_style`` and ``find_style_id`` and ``custom_style_name`` in :src_pkg:`model.py`).
* ``ParagraphStyleArchive`` (:src_proto:`TSWPArchives.proto`) and
  ``CellStyleArchive`` (:src_proto:`TSTArchives.proto`) are linked to their
  parents through ``super.parent``. The inheritance chain followed when
  resolving a style property is not described.
* ``TableModelArchive.body_text_style``, ``header_row_text_style``,
  ``header_column_text_style`` and ``footer_row_text_style``
  (:src_proto:`TSTArchives.proto`) select the default text style for a cell
  based on its position.

Custom formats
--------------

* ``TSK.DocumentArchive.custom_format_list`` resolves to a
  ``TSK.CustomFormatListArchive`` (:src_proto:`TSKArchives.proto`) whose
  ``custom_formats`` and parallel ``uuids`` lists are read and extended
  by the format lookups in :src_pkg:`model.py`. The pairing between a table's
  format entries and this list is only summarized above.

Rich text and bullets
---------------------

* ``TableDataList`` rich text entries follow ``rich_text_table`` to
  ``RichTextPayloadArchive`` (:src_proto:`TSTArchives.proto`), then
  ``payload.storage`` to ``TSWP.StorageArchive``
  (:src_proto:`TSWPArchives.proto`), then the storage's attribute tables to
  ``ListStyleArchive`` objects for bullets and numbering. Hyperlinks are
  mentioned above but the list-style traversal is not.

Categories and grouping
-----------------------

* ``TableModelArchive.category_owner`` resolves to
  ``TST.CategoryOwnerArchive``, whose ``group_by`` reference leads to a
  ``GroupByArchive``. ``TableInfoArchive.category_order`` resolves to a
  ``CategoryOrderArchive`` whose ``uid_map`` maps row UUIDs to display
  order. ``GroupNodeArchive`` objects are found with ``find_refs`` and give
  group UUIDs and cell values (all in :src_proto:`TSTArchives.proto`). The
  decoding is in ``calculate_table_categories`` and ``group_uuid_values``
  in :src_pkg:`model.py` and is not documented.

Strokes
-------

* The stroke sidecar traversal is described, but not the ordering of
  ``StrokeLayerArchive.row_column_index`` lookups used when a layer for a given
  row or column is found or created in :src_pkg:`model.py`.

Logical cell messages
---------------------

* ``TST.Cell`` (:src_proto:`TSTArchives.proto`) has fields such as
  ``valueType``, ``numberValue``, ``stringValue``, ``richText``,
  ``formulaError``, styles, formats, comments and decimal high/low words. It
  is not the encoding used in the primary tile buffers, which use the v5 byte
  layout described above. It occurs in command, change, pasteboard and
  concurrent-cell archives, which are not yet covered
  (:src_proto:`TNCommandArchives.proto` and other command-archive schemas).

Style property archiving
------------------------

* :src_proto:`TSTStylePropertyArchiving.proto` holds the cell-specific style
  properties read by ``cell_property`` and related accessors in
  :src_pkg:`model.py`. :src_proto:`TSKArchives.proto` shared formatting
  structures and colors are only touched on above.

Package-level data
------------------

* ``PACKAGE_ID`` ``datas`` entries (``TSP.PackageMetadata`` in
  :src_proto:`TSPArchiveMessages.proto`) map image data identifiers to stored
  file names and are searched and extended when images are added in
  :src_pkg:`model.py`. ``Cell`` image lookup in :src_pkg:`cell.py` reads
  ``ObjectStore.file_store`` directly. The naming rules for stored data files
  are not documented.
