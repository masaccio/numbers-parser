:hidetoc: 1

Style Class
###########

.. currentmodule:: numbers_parser

``numbers_parser`` currently only supports paragraph styles and cell
styles. The following paragraph styles are supported:

-  font attributes: bold, italic, underline, strikethrough
-  font selection and size
-  text foreground color
-  horizontal and vertical alignment
-  cell background color
-  cell indents (first line, left, right, and text inset)

Numbers conflates style attributes that can be stored in paragraph
styles (the style menu in the text panel) with the settings that are
available on the Style tab of the Text panel. Some attributes in Numbers
are not applied to new cells when a style is applied. To keep the API
simple, the :py:class:`~numbers_parser.Style` class holds both paragraph
styles and cell styles. When a document is saved, the attributes not stored in a
paragraph style are applied to each cell that includes it.

Since :py:class:`~numbers_parser.Style` objects are shared, changing the
:py:attr:`~numbers_parser.Cell.style` property of a :py:class:`~numbers_parser.Cell`
object will change the style for that cell and in turn affect the styles
of all cells using that style.

.. autoclass:: Style()
   :members:
