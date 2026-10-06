##########################
The ``stingray`` Package
##########################

The :py:mod:`stingray` package implements a Schema-based File Reader.
This allows us to use flat-file and workbook data from Python without having
to clone an application for each physical file format or logical layout.

It also allows us to use semi-structured data like JSON, YAML or an outline.
This can be handled consistently with structured data.

This package includes

-   A definition for "workbook", "sheet" and "cell".  This can subsume :py:mod:`csv`,
    :py:mod:`xlrd` as well as XML parsers for XLSX and ODS files.  This makes the
    physical format transparent to an application.

-   A definition for a schema.  Not a complex XSD, but the limited,
    flat schema appropriate for rows in sheets of a workbook in approximately First Normal Form.
    This is extended to handle the simple hierarcical COBOL features.

-   Classes to load a schema that's embedded either in a sheet or in a separate file.

-   A COBOL schema loader as an extension to the default loaders.

This depends on Python 3.3.

The structure of the ``stingray`` package is as follows.

-   ``__init__.py``.  Some essential overhead.  See :ref:`stingray_init`.

-   ``cell.py``.  :ref:`cells` defines the :py:class:`Cell`
    class hierarchy.  Imported as :py:mod:`stingray.cell`.

-   ``sheet.py``.  :ref:`sheets` defines the :py:class:`Sheet` class hierarchy
    that supports sheets with embedded as well as external schema.
    Imported as :py:mod:`stingray.sheet`.

-   ``workbook.py``.  The :ref:`workbook` families.
    Imported as :py:mod:`stingray.workbook`.

-   ``schema``.  This package defines a schema and schema loaders.

    -   ``__init__.py``.  :ref:`schema`.  Imported as :py:mod:`stingray.schema`.
        This is the generic, flat schema and superclasses for the
        more complex COBOL schema.

    -   ``loader.py``.   :ref:`schema_loader`.
        A loader for the generic, flat schema.  Imported as :py:mod:`stingray.schema.loader`.
        Applications will often extend schema loaders to handle peculiar formats
        or multi-line headings or other workbook formatting.

-   ``cobol``.  This package extends a schema and schema loaders to handle COBOL files.

    -   ``__init__.py``.  :ref:`cobol`.
        Imported as :py:mod:`stingray.cobol`.
        These are extensions to :py:mod:`stingray.cell`, :py:mod:`stingray.sheet` and :py:mod:`stingray.workbook`.

    -   ``loader.py`` :ref:`cobol_loader`.
        A loader for COBOL-syntax schema.
        Imported as :py:mod:`stingray.cobol.loader`.
        
    -   ``defs.py``. :ref:`cobol_defs`.
        Base definitions used by both ``__init__.py`` and ``loader.py``.

-   ``snappy``. This module is a minimal implementation of a reader for files written
    with Snappy compression.

-   ``protobuf``. This module is a minimal implementation of a reader for objects represented
    using protobuf.

..  toctree::
    :hidden:

    stingray_init
