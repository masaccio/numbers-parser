.. raw:: html

   <style>
   .red {
      color:red;
   }
   </style>

.. role:: red


numbers-parser API
==================

Changes in version 4.0
~~~~~~~~~~~~~~~~~~~~~~

To better partition cell styles, background image data which was supported in earlier versions
through the methods ``image_data`` and ``image_filename`` is now part of the new
``cell_style`` property. Using the deprecated methods ``image_data`` and ``image_filename`` 
will issue a ``DeprecationWarning`` if used.The legacy methods will be removed in a
future version of numbers-parser.

:class:`NumberCell` cell values limited to 15 significant figures to match the implementation
of floating point numbers in Apple Numbers. For example, the value ``1234567890123456``
is rounded to ``1234567890123460`` in the same way as in Numbers. Previously, using
native ``float`` with no checking resulted in rounding errors in unpacking internal numbers.
Attempting to write a number with too many significant digits results in a ``RuntimeWarning``.

The previously deprecated methods ``Document.sheets()`` and ```Sheet.tables()`` are now only
available using the properties of the same name (see examples in this README).

API
~~~

.. currentmodule:: numbers_parser

.. autoclass:: Document
   :members:

.. autoclass:: Sheet()
   :members:

.. autoclass:: Table()
   :members:

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


.. autoclass:: Style
   :members:

.. _negative_formats:

.. autoenum:: NegativeNumberStyle
    :members:
    

.. _datetime_formats:

Date/time Formatting
~~~~~~~~~~~~~~~~~~~~

Date/time formats use Numbers notation rather than POSIX ``strftime`` as there are a number
of extensions. Date components are specified using directives which must be separated by
whitespace. Supported directives are:

+-----------+---------------------------+------------------------+
| Directive | Meaning                   | Example                |
+===========+===========================+========================+
| ``a``     | Locale's AM or PM         | am, pm                 |
+-----------+---------------------------+------------------------+
| ``EEEE``  | Full weekday name         | Monday, Tuesday, ...   |
+-----------+---------------------------+------------------------+
| ``EEE``   | Abbreviated weekday name  | Mon, Tue, ...          |
+-----------+---------------------------+------------------------+
| ``yyyy``  | Year with century as a    | 1999, 2023, etc.       |
|           | decimal number            |                        |
+-----------+---------------------------+------------------------+
| ``yy``    | Year without century as a | 00, 01, ... 99         |
|           | zero-padded decimal       |                        |
|           | number                    |                        |
+-----------+---------------------------+------------------------+
| ``y``     | Year without century as a | 0, 1, ... 99           |
|           | decimal number            |                        |
+-----------+---------------------------+------------------------+
| ``MMMM``  | Full month name           | January, February, ... |
+-----------+---------------------------+------------------------+
| ``MMM``   | Abbreviated month name    | Jan, Feb, ...          |
+-----------+---------------------------+------------------------+
| ``MM``    | Month as a zero-padded    | 01, 02, ... 12         |
|           | decimal number            |                        |
+-----------+---------------------------+------------------------+
| ``M``     | Month as a decimal number | 1, 2, ... 12           |
+-----------+---------------------------+------------------------+
| ``d``     | Day as a decimal number   | 1, 2, ... 31           |
+-----------+---------------------------+------------------------+
| ``dd``    | Day as a zero-padded      | 01, 02, ... 31         |
|           | decimal number            |                        |
+-----------+---------------------------+------------------------+
| ``DDD``   | Day of the year as a      | 001 - 366              |
|           | zero-padded 3-digit       |                        |
|           | number                    |                        |
+-----------+---------------------------+------------------------+
| ``DD``    | Day of the year as a      | 01 - 366               |
|           | minimum zero-padded       |                        |
|           | 2-digit number            |                        |
+-----------+---------------------------+------------------------+
| ``D``     | Day of the year           | 1 - 366                |
+-----------+---------------------------+------------------------+
| ``HH``    | Hour (24-hour clock) as a | 00, 01, ... 23         |
|           | zero-padded decimal       |                        |
|           | number                    |                        |
+-----------+---------------------------+------------------------+
| ``H``     | Hour (24-hour clock) as a | 0, 1, ... 23           |
|           | decimal number            |                        |
+-----------+---------------------------+------------------------+
| ``hh``    | Hour (12-hour clock) as a | 01, 02, ... 12         |
|           | zero-padded decimal       |                        |
|           | number                    |                        |
+-----------+---------------------------+------------------------+
| ``h``     | Hour (12-hour clock) as a | 1, 2, ... 12           |
|           | decimal number            |                        |
+-----------+---------------------------+------------------------+
| ``k``     | Hour (24-hour clock) as a | 1, 2, ... 24           |
|           | decimal number to 24      |                        |
+-----------+---------------------------+------------------------+
| ``kk``    | Hour (24-hour clock) as a | 01, 02, ... 24         |
|           | zero-padded decimal       |                        |
|           | number to 24              |                        |
+-----------+---------------------------+------------------------+
| ``K``     | Hour (12-hour clock) as a | 0, 1, ... 11           |
|           | decimal number from 0     |                        |
+-----------+---------------------------+------------------------+
| ``KK``    | Hour (12-hour clock) as a | 00, 01, ... 11         |
|           | zero-padded decimal       |                        |
|           | number from 0             |                        |
+-----------+---------------------------+------------------------+
| ``mm``    | Minutes as a zero-padded  | 00, 01, ... 59         |
|           | number                    |                        |
+-----------+---------------------------+------------------------+
| ``m``     | Minutes as a number       | 0, 1, ... 59           |
+-----------+---------------------------+------------------------+
| ``ss``    | Seconds as a zero-padded  | 00, 01, ... 59         |
|           | number                    |                        |
+-----------+---------------------------+------------------------+
| ``s``     | Seconds as a number       | 0, 1, ... 59           |
+-----------+---------------------------+------------------------+
| ``W``     | Week number in the month  | 0, 1, ... 5            |
|           | (first week is zero)      |                        |
+-----------+---------------------------+------------------------+
| ``ww``    | Week number of the year   | 0, 1, ... 53           |
|           | (Monday as the first day  |                        |
|           | of the week)              |                        |
+-----------+---------------------------+------------------------+
| ``G``     | AD or BC (only AD is      | AD                     |
|           | supported)                |                        |
+-----------+---------------------------+------------------------+
| ``F``     | How many times the day of | 1, 2, ... 5            |
|           | falls in the month        |                        |
+-----------+---------------------------+------------------------+
| ``S``     | Seconds to one decimal    | 0 - 9                  |
|           | place                     |                        |
+-----------+---------------------------+------------------------+
| ``SS``    | Seconds to two decimal    | 00 - 99                |
|           | places                    |                        |
+-----------+---------------------------+------------------------+
| ``SSS``   | Seconds to three decimal  | 000 - 999              |
|           | places                    |                        |
+-----------+---------------------------+------------------------+
| ``SSSS``  | Seconds to four decimal   | 0000 - 9999            |
|           | places                    |                        |
+-----------+---------------------------+------------------------+
| ``SSSSS`` | Seconds to five decimal   | 00000 - 9999           |
|           | places                    |                        |
+-----------+---------------------------+------------------------+
