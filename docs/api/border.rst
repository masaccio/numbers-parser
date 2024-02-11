:hidetoc: 1

Border Class
############

.. currentmodule:: numbers_parser

``numbers-parser`` supports reading and writing cell borders, though the
interface for each differs. Individual cells can have each of their four
borders tested, but when drawing new borders, these are set for the
table to allow for drawing borders across multiple cells. Setting the
border of merged cells is not possible unless the edge of the cells is
at the end of the merged region.

Borders are represented using the
`Border <https://masaccio.github.io/numbers-parser/border.html>`__
class that can be initialized with line width, color and line style. The
current state of a cell border is read using the
`Cell.border <https://masaccio.github.io/numbers-parser/api/cells.html#numbers_parser.Cell.border>`__

property. The
`Table.set_cell_border <https://masaccio.github.io/numbers-parser/api/table.html#numbers_parser.Table.set_cell_border>`__
sets the border for a cell edge or a range of cells.

.. autoclass:: Border
   :members:
