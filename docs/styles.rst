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
are not applied to new cells when a style is applied.


.. only:: MarkdownDocs

    To keep the API simple, ``numbers-parser`` packs all styling into a single
    :pages:`Style <api/style.html>` object. When a document is saved, the attributes 
    not stored in a paragraph style are applied to each cell that includes it.

    Styles are read from cells using the 
    :pages:`Cell.style <api/cells.html#numbers_parser.Cell.style>` property and you can
    add new styles with 
    :pages:`Document.add_style <api/document.html#numbers_parser.Document.add_style>`.

.. only:: HtmlDocs

    To keep the API simple, ``numbers-parser`` packs all styling into a single 
    :class:`~numbers_parser.Style` object. When a document is saved, the attributes not 
    stored in a paragraph style are applied to each cell that includes it.

    Styles are read from cells using the :py:class:`~numbers_parser.Cell` property
    :py:attr:`~numbers_parser.Cell.style` and you can add new styles to a
    :py:class:`~numbers_parser.Document` using :py:meth:`~numbers_parser.Document.add_style`.

.. code:: python

   red_text = doc.add_style(
       name="Red Text",
       font_name="Lucida Grande",
       font_color=RGB(230, 25, 25),
       font_size=14.0,
       bold=True,
       italic=True,
       alignment=Alignment("right", "top"),
   )
   table.write("B2", "Red", style=red_text)
   table.set_cell_style("C2", red_text)