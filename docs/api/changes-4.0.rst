:hidetoc: 1

Changes in version 4.0
======================

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

From version 4.15, a number of dependencies have been removed to simplify installation and
to remove some large dependencies such as Rust for Pendulum.