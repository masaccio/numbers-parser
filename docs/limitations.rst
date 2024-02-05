Limitations
-----------

Current known limitations of ``numbers-parser`` are:

-  Formulas cannot be written to a document
-  Table styles that allow new tables to adopt a style across the whole
   table are not planned.
-  Creating cells of type ``BulletedTextCell`` is not supported
-  New tables are inserted with a fixed offset below the last table in a
   worksheet which does not take into account title or caption size
-  New sheets insert tables with formats copied from the first table in
   the previous sheet rather than default table formats
-  Creating custom cell formats and cell data formats is experimental
   and not all formats are supported. See
   `Table.set_cell_formatting <https://masaccio.github.io/numbers-parser/#numbers_parser.Table.set_cell_formatting>`__
   for more details.
-  Due to a limitation in Python's
   `ZipFile <https://docs.python.org/3/library/zipfile.html>`__, Python
   versions older than 3.11 do not support image filenames with UTF-8
   characters (see `issue
   69 <https://github.com/masaccio/numbers-parser/issues/69>`__).
   `Cell.style.bg_image <https://masaccio.github.io/numbers-parser/#numbers_parser.Style>`__
   returns ``None`` for such files and issues a ``RuntimeWarning``.
- Pivot tables are unsupported, but re-saving a document is believed to work. Saving a document with a pivot table issues a `UnsupportedWarning`.
