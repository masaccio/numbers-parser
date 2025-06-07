Limitations
-----------

Current known limitations of ``numbers-parser`` which may be implemented in the future are:

- Table styles that allow new tables to adopt a style across the whole
  table are not suppported
- Creating cells of type ``BulletedTextCell`` is not supported
- New tables are inserted with a fixed offset below the last table in a
  worksheet which does not take into account title or caption size
- Captions can be created and edited as of `numbers-parser` version 4.12, but cannot
  be styled. New captions adopt the first caption style available in the current
  document
- Formulas cannot be written to a document
- Pivot tables are unsupported and saving a document with a pivot table issues
  a `UnsupportedWarning` (see :github:`issue 73 <issues/73>` for details).  

The following limitations are expected to always remain:

- New sheets insert tables with formats copied from the first table in
  the previous sheet rather than default table formats
- Due to a limitation in Python's
  `ZipFile <https://docs.python.org/3/library/zipfile.html>`__, Python
  versions older than 3.11 do not support image filenames with UTF-8 characters
  :pages:`Cell.add_style.bg_image() <api/sheet.html#numbers_parser.Style>` returns
  ``None`` for such files and issues a ``RuntimeWarning``
  (see :github:`issue 69 <issues/69>` for details).  
- Password-encrypted documents cannot be opened. You must first re-save without
  a password to read (see :github:`issue 88 <issues/88>` for details).
  A `UnsupportedError` exception is raised when such documents are opened.
- Due to changes in the format of Numbers documents, decoding of category groups
  (introduced in ``numbers-parser`` version 4.16) is supported only for documents
  created by Numbers 12.0 and later. No warnings are issued for earlier
  Numbers documents.