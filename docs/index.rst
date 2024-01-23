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

To better partition cell styles, background image data which was
supported in earlier versions through the methods ``image_data`` and
``image_filename`` is now part of the new ``cell_style`` property. Using
the deprecated methods ``image_data`` and ``image_filename`` will issue
a ``DeprecationWarning`` if used.The legacy methods will be removed in a
future version of numbers-parser.

``NumberCell`` cell values are now limited to 15 significant figures to
match the implementation of floating point numbers in Apple Numbers. For
example, the value ``1234567890123456`` is rounded to
``1234567890123460`` in the same way as in Numbers. Previously, using
native ``float`` with no checking resulted in rounding errors in
unpacking internal numbers. Attempting to write a number with too many
significant digits results in a ``RuntimeWarning``.

The previously deprecated methods ``Document.sheets()`` and
``Sheet.tables()`` are now only available using the properties of the
same name (see examples in this README).

API
~~~

.. currentmodule:: numbers_parser

.. autoclass:: Document
   :members:

.. autoclass:: Sheet()
   :members:

.. autoclass:: Table()
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
| ``a``     | Locale’s AM or PM         | am, pm                 |
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


.. _negative_formats:

Negative Number Formats
~~~~~~~~~~~~~~~~~~~~~~~


.. currentmodule:: numbers_parser

.. autoenum:: NegativeNumberStyle
    :members:

**Example**

======================= =================
Value                   Examples
======================= =================
``MINUS``               -1234.560
``RED``                 :red:`1234.560`
``PARENTHESES``         (1234.560)
``RED_AND_PARENTHESES`` :red:`(1234.560)`
======================= =================
