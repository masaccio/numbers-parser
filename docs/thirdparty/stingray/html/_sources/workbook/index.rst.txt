.. _`workbook`:

###############################################################
Workbook Package -- Uniform Wrappers for Workbooks
###############################################################


..  py:module:: workbook

A *Workbook* is a collection of *Sheets*.  It's also a set of decoding
rules required to translate bytes (or XML text) into meaningful *Cell* instances.

Access to cells of a Workbook requires two levels of schema:

-   *Physical Format*.  The format required to locate cells.
    CSV, XLS, XLSX, ODS, are all well-known physical formats and the physical
    schema is implied by the file type.
    Fixed format and COBOL format, are not well-known, and a physical
    schema is required.

-   *Logical Layout*. The columns or data elements present in the file.
    This may depend on an embedded schema in the first rows of a Sheet.
    Or it may depend on an external schema defined in another Workbook.

This package addresses the physical format issues. It provides a common
abstraction over a number of forms of workbook data.  It makes the physical
format largely transparent to an application.

It's difficult to make the logical layout transparent.
See :ref:`developer` for guidelines on developing applications that
are flexible with respect to logical layout.

In a way, a Workbook is a factory for :py:class:`sheet.Sheet` and
:py:class:`sheet.Row` objects.

More interestingly, a Workbook is a factory for :py:class:`cell.Cell` instances.
This is because the decoding of bytes to create a cell is entirely a feature
of the Workbook.

Use Case
==============

See :ref:`intro` for our physical-format independence use case.
A :py:func:`workbook.open_workbook` function allows a program to be
independent of physical format.

..  parsed-literal::

    def process_workbook_file( input ):
        with workbook.open_workbook( input ) as source:
            process_workbook( source );

    if __name__ == "__main__":
        *application startup*
        for input in args.file:
            process_workbook_file( input )


This does not address logical layout issues, however, which are handled by a
:py:class:`schema.Schema`.  We might load an embedded schema or an external schema.

..  parsed-literal::


    def process_workbook( source ):
        for name in source.sheets():
            sheet= source.sheet( name,
                sheet.EmbeddedSchemaSheet,
                loader_class=schema.loader.HeadingRowSchemaLoader )
            counts= process_sheet( sheet )
            pprint.pprint( counts )

    def process_sheet( sheet ):
        """Separated to facilitate unit testing"""
        counts= defaultdict( int )
        for rows in sheet.rows():
            *process cells of this row*
        return counts


Physical Formats
=======================================

Much data is transferred via formats
tied to desktop spreadsheet software or
informed by legacy mainframe design patterns.
Data that comes from spreadsheet applications
will have all the rich variety of desktop tools.

-   :ref:`workbook_csv`. This includes the "quote-comma" dialects as used by spreadsheets
    as well as "tab" or "pipe" dialects favored by Linux applications.

-   :ref:`workbook_ods`. This is a zipped archive of XML documents from which data can be extracted.
    This is an ECMA standard.  This is the Open Office Spreadsheet structure.
    Most of the relevant data is in a content.xml member.

-   :ref:`workbook_xlsx`. 
    This is a zipped archive of XML documents from which data can be extracted.
    This is an ECMA standard.

-   :ref:`workbook_xls`. 
    This is the proprietary "Horrible Spreadsheet Format" (HSSF) as used by
    Microsoft products.  
    We require `xlrd <http://www.lexicon.net/sjmachin/xlrd.htm>`_ 
    to extract data from these files.
    
    If we can't import the :mod:`xlrd` module, an error will be raised only when trying
    to open one of these files.

-   :ref:`workbook_number09`.
    The iWorks '09 physical format is a simple ZipFile with a big XML document. 
    In many respects it's similar to XLSX format.

-   :ref:`workbook_number13`.
    iWorks '13 physical format is the "bundle" or "package" format; the document
    is a directory, which contains a zip archive of .IWA files. These use snappy
    compression and protobuf object representation. The :ref:`other_modules`
    are separate from this workbook module.

-   :ref:`workbook_fixed`. 
    Yes, these files still exist.  For
    these files, schema information is *required* to determine where
    the fields are, since there's no puctuation. We can convert EBCDIC bytes or work
    in Unicode-compatible text. ASCII encoding is usually handled trivially by
    Python's ``io`` module.

-   JSON, YAML, XML. For example, an Omni Outliner outlines with a normalized format.
    This is a possible future direction.


We'll call ``CSV``, ``XLS``, ``XLSX`` / ``XLSM`` and ``ODS``
the "well-known physical formats."
They don't require physical schema information in order
to identify the data items.

The Fixed and COBOL format files, on the other hand, require physical schema information.  
We'll look at COBOL in depth, in :ref:`cobol`.

Model
======

..  code-block:: none

    http://yuml.me/diagram/scruffy;/class/
    #workbook,
    [Workbook]^[CSV_Workbook],
    [Workbook]^[XLS_Workbook],
    [Workbook]^[XLSX_Workbook],
    [Workbook]^[Fixed_Workbook],
    [Workbook]^[ODS_Workbook],
    [Workbook]<>-[Sheet],
    [Sheet]<>-[Row],
    [Workbook]->[Schema].


..  image:: workbook.png
    :width: 6in

Workbook Implementation
========================

These modules implement the various kinds of workbooks that Stingray
can process.

..  toctree::
    :maxdepth: 1

    init
    base
    csv
    xls
    xlsx
    ods
    numbers_09
    numbers_13
    fixed
