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
   and maintained by AI.

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

Protobuf wire data is not self-describing. ``MessageInfo.type`` is resolved
through the Numbers/common registry extracted from the iWork applications;
the resulting maps are checked into :src_root:`src/numbers_parser/generated/mapping.py`.
The same numeric id can mean a different class in another iWork application.
The schema field types and message definitions live in the ``.proto`` files,
not in the bytes on disk. Likewise, a ``TSP.Reference`` identifies an object
but does not say what kind it points to.

``ArchiveInfo.identifier`` is the object's numeric identity within the
document. ``MessageInfo.object_references`` and ``data_references`` record
cross-object/resource references; the semantic links are also represented as
typed ``TSP.Reference`` and ``TSP.DataReference`` fields in payload
messages. The schema also permits ``should_merge`` and patch metadata
(``base_message_index``, ``diff_field_path`` and related fields); this is
protobuf-level object patching, not a separate cell encoding. The current
``ProtobufPatch`` support in ``iwafile.py`` is intentionally limited.

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

For a table, the central chain is::

    TN.DocumentArchive
      └── sheets[] ──> TN.SheetArchive
                       └── drawable_infos[] ──> TST.TableInfoArchive
                                                └── tableModel ──> TST.TableModelArchive
                                                                   └── base_data_store ──> TST.DataStore
                                                                                           └── tiles ──> TST.Tile
                                                                                                         └── rowInfos[]

``TST.TableInfoArchive`` wraps the drawable and refers to its
``TableModelArchive``. The model stores the persistent table UUID, dimensions,
name, default row/column sizes, header/footer counts and row/column freezing,
plus references to default body/header/footer styles. It also carries
information for sorting, hidden and filtered rows, merges, styles, categories,
pivot tables, and other newer table capabilities. Do not mistake a
``TableInfoArchive`` (view/drawable metadata) for the ``TableModelArchive``
(table's dimensions, data and model properties).

The reader treats object identifiers 1 and 2 as the document and package
roots (``DOCUMENT_ID`` and ``PACKAGE_ID`` in ``constants.py``). The package
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

``TST.TileRowInfo`` contains a ``cell_count``, row index, cell storage bytes,
cell offsets, a storage version, and an optional ``has_wide_offsets`` flag.
The offsets map column positions to the corresponding cell record in the
row's storage byte buffer. A negative offset marks a column with no cell
record. A wide-offset row stores offsets in units of four bytes; the parser
scales those offsets before slicing the byte buffer. Empty columns between
stored cells therefore do not consume a cell record. ``get_storage_buffers_for_row``
uses the next nonnegative offset (or the end of the buffer) to delimit each
cell.

The shared values are represented by ``TST.TableDataList`` entries. Each
entry has a numeric ``key`` and ``refcount``; its payload depends on the list
type. The proto enumerates string, format, formula, style, formula-error,
custom-format, choice-list-format, rich-text, conditional-style, comment,
import-warning, and control-cell-spec lists. A string entry contains a
string; other entries may hold a protobuf value or a reference to another
archive. The cell's stored key indexes its table's list, not a document-wide
string table. ``DataLists`` in ``model.py`` caches lookups, reuses existing
values, and allocates subsequent keys when writing new entries.

The plain string table and rich-text table are distinct. A plain text cell
looks up its string id in ``stringTable``. Rich text uses a rich-text payload
reference and ``TSWP.StorageArchive`` text runs, with character-indexed
attribute tables for paragraph and character styling. Hyperlinks are attached
to rich-text fragments (for example a ``TSWP.HyperlinkFieldArchive``), not
stored as an independent cell-level URL. This differs from formats that have
a hyperlink property on each cell.

Cell storage: v5 binary records
===============================

The cell protobuf ``TST.Cell`` describes the logical value and formatting,
but table tiles also store a compact binary record per cell. The current
``numbers-parser`` cell reader and writer support the v5 layout (called BNC
or post-BNC in the research sources). ``model.py`` rejects a tile that is
not marked as saved in BNC form, and ``cell.py`` rejects a cell record whose
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

The field mask has these meanings in the v5 layout (``src/protos/TSTArchives.proto``
and the SheetsJS research describe the storage map; ``cell.py`` implements
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
``_unpack_decimal128`` in ``cell.py`` extracts a signed integer mantissa and
biased base-10 exponent using ``DECIMAL128_BIAS`` from ``constants.py``;
the API currently exposes the resulting number as a Python numeric value.
When writing, ``_pack_decimal128`` builds the 16-byte representation and
numbers are rounded to the library's configured significant-digit limit
before serialization. The separate double bit is used for boolean values
and durations.

Numbers dates are represented as seconds relative to 2001-01-01, defined as
``EPOCH`` in ``constants.py``. Date records carry an eight-byte double
seconds value. Durations also carry seconds as a double, but are elapsed
intervals, not dates relative to the epoch. The format id controls how these
values are displayed. Number, currency, date, duration, text, and boolean
format ids are separate optional fields in v5 and resolve through the
table's format list.

The ``CellType`` enum exported by ``constants.py`` is an API-level
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
entry contains a :src_proto:`TSCE.FormulaArchive`` (see ``TSCEArchives.proto`). A
formula archive contains an abstract syntax tree, host row/column and table
identity, translation flags, and related formula metadata. The AST is made
of typed nodes for literals, operators, functions, and cell/range
references. References can be local or cross-table and may use row/column
UUIDs; these stable identities matter when rows or columns move. The parser
uses current cell values rather than recalculating formula expressions.
``cell.py`` resolves a formula id and its coordinates, while ``formula.py``
handles formula rendering and references.

Formatting is also indirect. The v5 cell record can carry style ids,
conditional-style data, format ids and a control-spec id. Table list entries
map keys to the actual style, format, formula, or control data. A cell style
can reference ``TST.CellStyleArchive`` and its cell properties; text
properties use the TSWP text/style messages. The table model supplies default
body, header-row, header-column and footer styles, with per-cell entries
providing overrides. The document-level stylesheet and theme provide shared
style definitions and theme context.

``constants.py`` names the API and format enumerations for standard formats
(base, currency, date/time, fraction, number, percentage, scientific, text,
checkbox, rating, duration and custom formats), interactive controls (popup,
rating, slider, stepper and tickbox), negative-number style, fraction
accuracy, and duration units. These enum values describe properties and
formats, not the type byte or field mask in the cell buffer.

Controls have their own ``TST.CellSpecArchive``. Its interaction kind and
optional range limits, increment, or popup-model reference provide the
behavior behind the cell's control-spec id. ``FormattingType`` and related
maps in ``constants.py`` connect API formatting choices to protobuf format
archives. Custom formats also have a document-level custom-format list and
UUIDs, while table entries refer to the corresponding custom format.

The numeric ``FormatType`` codes in ``constants.py`` are: boolean 1, decimal
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
``constants.py`` maps Numbers tokens for years, months, days, weekday names,
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
defined in :src_proto:`TSTArchives.proto`. ``model.py`` recognizes the representations
it encounters, and ``cell.py`` exposes a merged anchor and covered-cell
references to the higher-level API.

Rows and columns can have numeric indexes and stable UUID identities. The
schema includes row/column UID maps and UUID-based ranges as well as older
coordinate-based ranges. Formula references, cell selections, merge maps and
table category/pivot features may use these stable identifiers. Table UUID
mapping and formula-owner dependencies are handled in ``model.py``; they are
not equivalent to the ``ArchiveInfo.identifier`` used to locate a protobuf
object.

Other archive content
=====================

Numbers archives contain much more than cell values. The following schema
families explain other content that can be encountered while traversing a
document:

* :src_proto:`TSAArchives.proto` and :src_proto:`TSDArchives.proto` describe common document
  and drawable objects, geometry, fills, strokes, images and media placement.
  A sheet's drawings and tables are both drawable objects.
* :src_proto:`TSSArchives.proto` describes styles, style properties, themes and
  style networks. :src_proto:`TSTStylePropertyArchiving.proto` includes cell-specific
  style properties.
* :src_proto:`TSWPArchives.proto` describes text storage, character/paragraph
  attributes, attachments and hyperlinks used by rich text and other
  document text.
* :src_proto:`TSKArchives.proto` contains formatting structures, custom formats,
  colors and shared application-level properties.
* :src_proto:`TSCEArchives.proto` describes calculation-engine formulas, references,
  cell values, dependencies, spill data and related calculation metadata.
* :src_proto:`TSCHArchives.proto` and :src_proto:`TSCH3DArchives.proto` describe charts and
  their data/format state. :src_proto:`TNArchives.proto` adds Numbers-specific chart
  mediation and sheet/document details.
* :src_proto:`TNCommandArchives.proto` and the other ``*CommandArchives.proto`` files
  describe editing commands/history. Command records are not the canonical
  table-value model, but are useful when investigating archives containing
  edit or patch state.
* :src_proto:`TSPArchiveMessages.proto` contains package/component/data metadata,
  object UUID maps and serialization metadata. :src_proto:`TSPMessages.proto` defines
  generic references, UUIDs, geometry primitives and shared value types.

The protobuf definition named ``TST.Cell`` is a logical/message
representation with fields such as ``valueType``, ``numberValue``,
``stringValue``, ``richText``, ``formulaError``, styles, formats, comments,
and decimal high/low words. It is important not to assume that this message
is the literal encoding of cells in the table's primary tile buffers. The
v5 tile cell buffers use the compact byte layout described above; related
``TST.Cell`` messages also occur in command, change, pasteboard, and
concurrent-cell archives.

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

On write, updated cell values are encoded as v5 cell buffers; strings,
formats, styles and formulas are inserted into their relevant table lists.
The model rebuilds tile-row buffers and offsets, updates protobuf lengths
and object references, serializes segments, recompresses IWA chunks, and
writes the ZIP file or package. Non-IWA blobs (for example, image data) are
kept in the file store. ``src/numbers_parser/model.py`` and
``src/numbers_parser/iwafile.py`` are the best references for following this
process end to end.

Appendix: UUIDs, owners, and table relationships
================================================

The following details extend the formula, merge, and stable-coordinate
descriptions above with observations from the project's
`Numbers.md research <https://github.com/masaccio/numbers-parser/blob/main/docs/Numbers.md>`_.
They describe observed structures, not requirements for every Numbers version.
The relevant definitions are in
`TSCEArchives.proto <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSCEArchives.proto>`_,
`TSTArchives.proto <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSTArchives.proto>`_,
and `TSPMessages.proto <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSPMessages.proto>`_.

Archive object identifiers, internal owner ids, and UUIDs are distinct
identifiers. A ``TSP.Reference`` points to an archive object by its numeric
``identifier``; an owner id map relates an internal integer to a UUID. Do not
use one as a substitute for another.

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
`constants.OwnerKind <https://github.com/masaccio/numbers-parser/blob/main/src/numbers_parser/constants.py#L256-L259>`_.
These values are not declared as an enum alongside the protobuf field.

In observed files, some formula-owner UUIDs share their upper 112 bits while
their lower 16 bits vary with the formula id. This pattern is an observation,
not a UUID-generation rule.

For a table, the ``TableModelArchive.haunted_owner.owner_uid`` has been
observed to match the ``formula_owner_uid`` of a dependency archive whose
``owner_kind`` is 35. That archive's ``base_owner_uid`` supplies the stable
UUID used to associate dependencies with the table. In the implementation,
`calculate_table_uuid_map() in model.py
<https://github.com/masaccio/numbers-parser/blob/main/src/numbers_parser/model.py#L763-L801>`_
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

The example documents in ``Numbers.md`` show table-model dependency archives
with spanning ranges for both the whole table and its body. They also show a
variation in ``tiled_cell_dependencies``: the first example had no tile
reference, while later examples referred to ``CellRecordTileArchive`` records.
The referenced tiles carry an ``internal_owner_id`` and tile row/column
origins. Treat this as observed variation, not a rule that all later tables
must have tiles or that all first tables omit them.

Merge and formula ranges
------------------------

One observed merge representation is ``RangePrecedentsTileArchive``. Its
``from_to_range`` entries pair a starting coordinate with a rectangle, and
``to_owner_id`` identifies the destination owner. Resolve that integer using
the calculation engine's ``owner_id_map`` before relating the rectangle to a
table UUID. The current parser also handles merge ranges in
``FormulaOwnerDependenciesArchive`` records of merge-owner kind: it reads
their range dependencies, resolves their owner ids, and keeps ranges whose
base-owner UUID matches the table. These are distinct archive structures
representing related range information.

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

Captions and text storage
-------------------------

Caption text is stored through a nested shape/storage structure rather than
directly in a caption record. ``TSA.CaptionInfoArchive`` contains a
``TSWP.ShapeInfoArchive``; its ``owned_storage`` reference points to a
``TSWP.StorageArchive`` whose ``text`` field contains the text. The shape
messages also carry drawable and placement information. The
`TSAArchives.proto <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSAArchives.proto>`_,
`TSWPArchives.proto <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSWPArchives.proto>`_,
and `TSDArchives.proto <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSDArchives.proto>`_
definitions show the required ``super`` chain and storage fields. In the
observed documents, the storage archive was located through the caption's
``owned_storage`` reference; the research notes report no direct
``Metadata.json`` reference except for its ``object_uuid_map_entries`` listing.

Appendix: decoding references
=============================

* Original IWA description: ``docs/thirdparty/obriensp_docs.md`` on the
  ``feat/format-docs`` branch.
* SheetsJS IWA and v5 cell-format research:
  ``docs/thirdparty/SheetsJS.html`` on the ``feat/format-docs`` branch. The
  format discussion is based on that document's v5 section only.
* Stingray's relevant historical material:
  ``docs/thirdparty/stingray/html/protobuf.html``,
  ``docs/thirdparty/stingray/html/snappy.html``, and
  ``docs/thirdparty/stingray/html/workbook/numbers_13.html``. Its
  protobuf-format account and Numbers 13 examples supplement the primary
  research; unrelated XLS, COBOL and other format material is outside this
  chapter.
* Schemas: the ``.proto`` files under ``src/protos``. Start with
  `TSPArchiveMessages.proto <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSPArchiveMessages.proto>`_
  for archive framing metadata, `TNArchives.proto
  <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TNArchives.proto>`_
  for Numbers document and sheet structure, `TSTArchives.proto
  <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSTArchives.proto>`_
  for tables/cells/data lists, and `TSCEArchives.proto
  <https://github.com/masaccio/numbers-parser/blob/main/src/protos/TSCEArchives.proto>`_
  for formulas and calculation references.
* Implementation: ``src/numbers_parser/iwork.py``,
  ``src/numbers_parser/iwafile.py``, ``src/numbers_parser/containers.py``,
  ``src/numbers_parser/model.py``, ``src/numbers_parser/cell.py``, and
  ``src/numbers_parser/constants.py``.
