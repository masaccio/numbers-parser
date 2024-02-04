:hidetoc: 1

Cell Classes
############

.. currentmodule:: numbers_parser

.. autoclass:: Cell()
   :no-undoc-members:
   :members:

.. _table_cell_merged_cells:

.. autoclass:: MergedCell()
   :no-undoc-members:
   :members:

``Cell.is_merged`` returns ``True`` for any cell that is the result of
merging rows and/or columns. Cells eliminated from the table by the
merge can still be indexed using ``Table.cell()`` and are of type
``MergedCell``.

.. raw:: html

   <table border="1">
         <tr>
            <td style="padding:10px">A1</td>
            <td style="padding:10px" rowspan=2>B1</td>
         </tr>
         <tr>
            <td style="padding:10px">A2</td>
         </tr>
   </table>
   <br>

The properties of merges are tested using the following properties:

+------+------------+-----------+---------------+----------+--------------+-----------------+
| Cell | Type       | ``value`` | ``is_merged`` | ``size`` | ``rect``     | ``merge_range`` |
+======+============+===========+===============+==========+==============+=================+
| A1   | TextCell   | ``A1``    | ``False``     | (1, 1)   | ``None``     | ``None``        |
+------+------------+-----------+---------------+----------+--------------+-----------------+
| A2   | TextCell   | ``A2``    | ``False``     | (1, 1)   | ``None``     | ``None``        |
+------+------------+-----------+---------------+----------+--------------+-----------------+
| B1   | TextCell   | ``B1``    | ``True``      | (2, 1)   | ``None``     | ``None``        |
+------+------------+-----------+---------------+----------+--------------+-----------------+
| B2   | MergedCell | ``None``  | ``False``     | ``None`` | (1, 0, 2, 0) | ``"B1:B2"``     |
+------+------------+-----------+---------------+----------+--------------+-----------------+

The tuple values of the ``rect`` property of a ``MergedCell`` are also
available using the properties ``row_start``, ``col_start``,
``row_end``, and ``col_end``.

.. autoclass:: RichTextCell()
   :show-inheritance:
   :no-undoc-members:
   :members:
