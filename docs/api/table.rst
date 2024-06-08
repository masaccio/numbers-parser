:hidetoc: 1

Table Class
###########

.. currentmodule:: numbers_parser

Each `Table` object represents a single table in a Numbers sheet. Tables are returned by
the the :py:class:`~numbers_parser.Sheet` property :py:attr:`~numbers_parser.Sheet.tables`
and can be indexed using either list style or dict style indexes:

.. code:: python

   >>> from numbers_parser import Document
   >>> doc = Document("mydoc.numbers")
   >>> sheet = doc.sheets["Sheet 1"]
   >>> sheet.tables[0]
   <numbers_parser.document.Table object at 0x1063e5d10>
   >>> sheet.tables[0].name
   'Table 1'
   >>> sheet.tables["Table 1"].name
   'Table 1'

.. NOTE::

   Do not instantiate directly. Tables are created by :py:class:`~numbers_parser.Document`.

.. autoclass:: Table()
   :members:
