
..  _`history`:

##############
History
##############

Latest release is 4.5

Version 4
==========

Version 4 dates from March, 2014. It switches to Python3.

Change Details:

-   Added support for parent-level USAGE clauses.

-   Added a "replacing" option to the COBOL Schema Loader.

    ..  parsed-literal::

        with open("xyzzy.cob", "r") as cobol:
            schema= stingray.cobol.loader.COBOLSchemaLoader( cobol, replacing=("'WORD'", "BAR") )

    ..

    The idea is to permit a copybook that expects "REPLACING" to be parsed.

-   Handle multiple 01-level declarations.
    
    When we look at the context for :py:meth:`cobol.loader.RecordFactory.makeRecord`, this 
    becomes an iterator which yields each ``01`` level DDE. 
    
    When the stack is popped
    down to empty, yield the structure and then start parsing again.
    
    The :py:meth:`cobol.loader.COBOLSchemaLoader.load` method builds the
    schema, iterates through all the ``01`` objects returned by :py:meth:`cobol.loader.RecordFactory.makeRecord`,
    and returns ``v.schema`` as the resulting schema.
    
    Any unit test that uses :py:meth:`cobol.loader.RecordFactory.makeRecord` must change to reflect
    it's revised interface as an iterator.

-   Add some COBOL demos.

-   Handle more complex VALUES clauses for more complex 88-level items.

-   Restructure cobol, cobol.loader to add a cobol.defs module.

-   Handle Occurs Depending On. Parse the syntax for ODO. Update LazyRow to 
    tweak size and offset information for each row fetched.

-   Add Z/OS RECFM handling in the :py:class:`cobol.EBCDIC_File` class.
    This will allow processing "Raw" EBCDIC files with RECFM of V and
    RECFM of VB -- these files include BDW and RDW headers on blocks 
    and records.

-   Fix the :py:func:`cobol.dump` to properly iterate through all fields.

-   Fix the :py:class:`cobol.defs.Usage` hierarchy to properly handle
    data which can't be converted. An ErrorCell is created in the (all too common)
    case where the COBOL data is invalid.

-   Support iWork '09 and iWork '13 Numbers Workbook files.
    This lead to a profound refactoring of the :py:mod:`stingray.workbook` module
    into a package.

-   Remove ``%`` string formatting and ``from __future__``. Correct class
    references. Replace ``u''`` Unicode string literals with simple string literals.
    https://sourceforge.net/p/stingrayreader/tickets/6/
    
-   Updated documentation. 
    https://sourceforge.net/p/stingrayreader/tickets/7/

-   Handled precision of comp3 correctly.
    https://sourceforge.net/p/stingrayreader/tickets/9/

-   Added :py:class:`cobol.loader.Lexer_Long_Lines` to parse copybooks with
    junk in positions 72:80 of each line. 
    https://sourceforge.net/p/stingrayreader/tickets/11/

-   Added RECFM=N to handle variable-length files with NO BDW/RDW words.
    This is the default. 
    https://sourceforge.net/p/stingrayreader/tickets/12/
    
-   Fixed Occurs Depending On Calculation initialization error.
    https://sourceforge.net/p/stingrayreader/tickets/15/
    
-   Tweaked performance slightly based on profile results.

-   Make embedded schema loader tolerate blank sheets by producing 
    a warning and returning ``None`` instead of raising an ``StopIteration`` exception.
    Tweak the Data validation demo to handle the None-instead-of-schema feature.
    
-   Changed :py:meth:`cobol.COBOL_file.row_get` to leave trailing spaces
    intact. This may disrupt applications that expected stripping of usage DISPLAY
    fields.
    
    This created a problem of trashing COMP items that had values
    of 0x40 exactly -- an EBCDIC space.
    
-   Handle Compiler Control Statements ``EJECT``, ``SKIP1``, ``SKIP2``, and ``SKIP3`` 
    by silently dropping them in the lexical scanner.
    
-   Changed ``RENAMES`` handling to be a warning instead of an exception.
    This allows compiling -- but not fully processing -- DDE's with 
    RENAMES clauses.

-   Handle "subrecord" processing. See :py:class:`stingray.test.cobol_2.Test_Copybook_13( DDE_Test )`.
    The idea is that we can do 
    ``subrow= data.subrow( self.segment_abc, row.cell(schema_header_dict['GENERIC-FIELD'])  )``
    to map a field, ``GENERIC-FIELD``, to an 01-level schema, ``self.segment_abc``.
    We can then pick fields out of ``subrow`` using fields defined in ``self.segment_abc``.
    
-   Add :py:func:`cobol.loader.COBOL_schema()` and :py:func:`cobol.loader.COBOL_schemata()`
    functions to provide a higher-level API for building a record schema or 
    a (less common) multiple schemata.
    
-   Fix a bug in cobol.RECFM_VB.bdw_iter() function.

-   Fix a bug in handling signed usage display EBCDIC numbers.

-   Fix a bug in handling complex picture clauses with ``9(x)v9(y)`` syntax.
    
-   Added some unit tests to confirm some previous fixes. Cleanup testing
    and build to make it easier to test a single class as part of debugging.
    
Version 3
==============

Version 3 dates from August of 2011.  It unifies COBOL DDE 
processing with Workbook processing.  They're both essentially the 
same thing.

The idea is to provide a schema that structures access to a file.

And release a much better version of the data profiling for COBOL files.

Version 2
============

An almost -- but not quite -- unrelated development was a library to unify
various kinds of workbooks.

This was started in '06 or so.  The context was econometric data analysis.
The sources were rarely formatted consistently.  While spreadsheets were
common, fixed-format files (clearly produced by COBOL) had to be handled 
gracefully.

The misdirection of following the :py:mod:`csv` design patterns for eager
loading of rows created small complications that were worked around badly
because lazy row loading was missing from the design.

The idea of the separation of physical format from logical layout was
the obvious consequence of the endless format and layout variations 
for a single conceptual schema.

Also, this reinforced the uselessness of trying to create a data-mapping
DSL when Python expresses the data mapping perfectly.

The data mapping triple is

..  parsed-literal:: 

    target = source.conversion()
    
Since this is -- effectively -- Python code, a DSL to do this is a waste of time.

Version 1
=============

Version 1 started in '02 or so.  Again, the context is data warehouse processing.

COBOL-based files were being examined as part of a data profiling exercise.

The data profiling use case  was very simple.  In effect, it was something 
like the following.

..  parsed-literal::

    summary = defaultdict( lambda: defaultdict(int) )
    def process_sheet( sheet ):
        for row in schema.rows_as_dict_iter(sheet.rows()):
            for k in row:
                summary[k][row[k]] += 1
        for attribute in summary:
            print( attribute )
            for k in summary[attribute]:
                print( k, summary[attribute][k] )

This version was a rewrite of the original C into Python.   

It was posted into SourceForge as https://sourceforge.net/projects/cobol-dde/.  

Version 0
================

Version 0 started in the late 90's.  In the context of data warehouse processing,
COBOL-based files were being moved from mainframe to a VAX (later a Dec Alpha).

The original processing included the following.

1.  Convert the EBCDIC files from mixed display and COMP-3 to all display.

2.  Copy the files from Z/OS to the VAX (or Alpha) via a magnetic tape transfer.
    This handled EBCDIC to ASCII conversion. (It was the 90's.) 

3.  Convert the resulting ASCII files from all display back to the original
    mixture of display and COMP-3 to resurrect the original data.
    
4.  Process the files for warehouse loading.

The first version of this schema-based file reader did away with the painful,
not-system-utility copying steps.  It reduced the processing to this.

1.  Copy the files from Z/OS to the VAX (or Alpha) via a magnetic tape transfer.
    Do no conversion.  The resulting file was mixed display and COMP-3 in EBCDIC 
    encoding.
    
2.  Use the COBOL source DDE to determine field encoding rules.  Copy the
    source file from mixed display and COMP-3 in EBCDIC 
    encoding to mixed display and COMP-3 in ASCII 
    encoding.
    
3.  Process the files for warehouse loading.

This was written in C.  

