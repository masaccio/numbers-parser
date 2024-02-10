:hidetoc: 1

Sheet Class
###########

.. currentmodule:: numbers_parser

Each `Sheet` object represents a single sheet in a Numbers document. Sheets are returned by
the the :py:class:`~numbers_parser.Table` property :py:attr:`~numbers_parser.Document.sheets`
and can be indexed using either list style or dict style indexes:

.. code:: python

   >>> from numbers_parser import Document
   >>> doc = Document("mydoc.numbers")
   >>> doc.sheets[0]
   <numbers_parser.document.Sheet object at 0x105f12390>
   >>> doc.sheets[0].name
   'Sheet 1'
   >>> doc.sheets["Sheet 1"].name
   'Sheet 1'

.. NOTE::
   Do not instantiate directly. Sheets are created by :py:class:`~numbers_parser.Document`.

.. autoclass:: Sheet()
   :members:
