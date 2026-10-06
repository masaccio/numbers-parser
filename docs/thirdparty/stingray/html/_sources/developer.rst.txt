
.. _`developer`:

######################################
The **Stingray** Developer's Guide
######################################

We use **Stingray** to work with data files where the schema is 
either external or complex (or both). We can tackle this question:

    **How do we process a file simply and consistently?**
    
Or, more concretely, 

    **How do we make a program independent of Physical Format 
    and Logical Layout?**
    
We can also use **Stingray** to answer questions about files, the schema those
files purport to use, and the data on those files.
Specifically, we can tackle this question:

    **How do we assure that a file and an application use the same schema?**
        
There are two sides to schema use: 

-   **Application Use.** Given a data file, we need to bind schema information to that file.
    For a workbook, the schema may be in the column headings of the sheet.
    Or it may be in a separate sheet, or a separate workbook.
    For a COBOL file, the schema is always in a separate COBOL-language
    data definition.

-   **Schema Conformance.** Given a data file, how do we confirm that some schema matches the file? 
    This is the data quality assurance question.

We do need to note the following.

    **If it was simple, we wouldn't need this package, would we?**
    
Concepts
========

As noted in :ref:`intro`, there are three levels of schema that need to be bound to a file.

-   **Physical Format**.  We can make this transparent to our applications.
    See :ref:`workbook` and :ref:`cobol` for details. Everything is a workbook
    with a fairly limited API.
    
-   **Logical Layout**.  This is how an application program will make use
    of the data found in a file. 
    Sometimes the Logical layout information can be embedded in a file: 
    it might be the top row of a sheet in a workbook, for example.
    Sometimes the logical layout can be separate: a COBOL data definition, or perhaps
    a metadata sheet in a workbook.
    
    We can't make the logical layout transparent.
    Our applications and files both need to agree on a logical layout. 
    The column names in the metadata, for example, must be agreed to.
    The order or position of the columns, however, need not be fixed.
    
-   **Conceptual Content**.  
    A single conceptual schema may be implemented by a number of physical
    formats and logical layouts.  
    An application should be able to tolerate variability in the logical
    layout as long is it matches the expected **conceptual content**.
    
    Since we're working in Python, the conceptual schema is often a class
    definition. It might be a namedtuple or a SimpleNamespace, also. 
    
We'll tackle the schema binding in several pieces.

-   **File Schema**.  `Binding a Schema to a File`_ describes some preliminary
    operational steps that make Stingray work more simply and reliably.

-   **Processing**. These are the basic concepts. `Data Attribute Mapping -- Using a Schema`_, 
    and `Data Transformation`_.

-   **Application Design Patterns**. These are more complex issues.
    We can then dig into **Stingray** application programming in `stingray Application Design`_,
    `Variant Records and COBOL REDEFINES`_, and `Big Data Performance`_.

-   **Data Management**. `File Naming and External Schema`_,  
    `Binding a Schema to an Application`_, and `Schema Version Numbering`_.

-   **FAQ**. Some other design questions. `Frequently Asked Questions`_.
        
We'll look at some demonstration programs in :ref:`demo`.

Binding a Schema to a File 
=================================

We're only going to bind two levels of schema to a file.  The conceptual schema
would require some kind of formal ontology, something that's rarely available.

**Logical Layout**.  We rely on a schema, :py:class:`schema.Schema` to manage
the logical layout of a file.

A workbook has two ways to bind logical layout to data values. 
Our :py:class:`sheet.Sheet` subclass hierarchy requires a schema object.
We'll load the schema using a :py:mod:`schema.loader` 
components.

-   **Embedded**.  This is commonly seen as column titles within the sheet.  Or
    any variation on that theme. The common case of column titles is handled
    by a built-in schema loader, :py:class:`schema.loader.HeadingRowSchemaLoader`.
    Other variations are managed by building different schema loaders.
    
-   **External**.  This is a separate sheet or separate file. In this case, we
    can start with :py:class:`schema.loader.BareExternalSchemaLoader` to read
    an external schema. In the case of COBOL files, there's a separate 
    :py:class:`cobol.loader.COBOLSchemaLoader` that parses COBOL source to create
    a usable schema.
    
**Physical Format**.  Generally, a file name provides a hint as to the physical file format.
:file:`.csv`, :file:`.xls`, :file:`.xlsx`, :file:`.xlsm`, :file:`.ods`,
:file:`.numbers` describe the physical format.   

Our :py:class:`cell.Cell`, :py:class:`sheet.Sheet`, and
:py:class:`workbook.base.Workbook` handles many physical format details nicely.

Data Attribute Mapping -- Using a Schema
==========================================

Using a schema is the heart of the semantic problem.

We want to have just one application that is adaptable to a number
of closely-related data file schemata.  Ideally, there's one, 
but as a practical matter, there
are often several similar schema.

We need to provide three pieces of information, minimally.

-   Target attribute within our application.

-   Target data type conversion.

-   Source attribute based on attribute name or position.

We could use a number of different encodings for this relationship.
We could write it as properties, or XML, or some other notation.

However, that triple is essentially a Python assignment statement
with *target*, *to_type* and *source*. The simplest description
is the following:

..  parsed-literal::

    *target* = row.cell( schema['*source*'] ).\ *to_type*\ ()

There is a tiny bit of boilerplate in this assignment statement. 
When using repeating groups, items with duplicated column names, or REDEFINES clauses,
the "boilerplate" expands to some additional code required to locate the source
data.

For multiple attributes, this is our use case: a **Builder Function**:

..  parsed-literal::

    def build_record( aRow ):
        return dict(
            name = row.cell( schema['some column'] ).to_str(),
            address = row.cell( schema['another column'] ).to_str(),
            zip = row.cell( schema['zip'] ).to_digit_str(5),
            phone = row.cell( schema['phone'] ).to_digit_str(),
        )
        
The idea is to build a single function that defines the
application-specific mapping from a row
in a file, given the logical layout information buried in the schema
definition.

Of course, the schema can lie, and the application can misuse the data.
Those are inevitable (and therefore insoluble) problems.  This is why
we must write customized software to handle these data sources.

In the case of variant schemata, we would like something like this.

..  parsed-literal::

    def build_record_1( aRow ):
        return dict(
            name = row.cell( schema['some column'] ).to_str(),
            address = row.cell( schema['another column'] ).to_str(),
            zip = row.cell( schema['zip'] ).to_digit_str(5),
            phone = row.cell( schema['phone'] ).to_digit_str(),
        )

    def build_record_2( aRow ):
        return dict(
            name = row.cell( schema['variant column'] ).to_str(),
            address = row.cell( schema['something different'] ).to_str(),
            zip = row.cell( schema['zip'] ).to_digit_str(5),
            phone = row.cell( schema['phone'] ).to_digit_str(),
        )

We can then define a handy factory which picks a builder based on the schema 
version.

..  py:function:: make_builder(args)

    Create a builder object from the args.

    :param args: schema version
    :returns: appropriate builder function for the schema
        
..  parsed-literal::

    def make_builder( args ):
        return eval( 'build_record_{0}'.format(args.layout) )

The :py:func:`make_builder` function selects one of the available
builders based on a run-time option.

Adding Fluency
---------------

In the case where there are no repeating groups, nor REDEFINES clauses, we can make our API
slightly more fluent by building a row dictionary from row and schema.  This kind of
eager cell processing is risky for COBOL files. It often works, however, for
well-known spreadsheet files.

..  parsed-literal::

    row_dict = dict( 
        (a.name, row.cell(a)) for a in schema )

This allows the following  *target*, *to_type* and *source* triple.

..  parsed-literal::

    *target* = row['*source*'].\ *to_type*\ ()

This parallels the :py:mod:`csv` module's ``DictReader`` class.


Data Transformation
=================================

In the :ref:`cells` chapter, we noted that there are two 
parts to data handling: **Capture** and **Conversion**.  Capture is part
of parsing the physical format.  Conversion is part of the final 
application, and has nothing to do with the schema that describes
the data source.

A target data type transformation (or conversion) could be considerably more complex
than the trivial case of decoding a floating-point number to a digit 
string.  It could involve any combination of filtering, cleansing, 
conforming to an existing database, or rewriting.

Here's a much more complex Builder Function.

..  parsed-literal::

    def build_record_3( aRow ):
        if aRow['flag'].is_empty():
            return None
        zip_str = aRow['zip'].to_str()
        if '-' not in zip:
            if len(zip) <= 5:
                zip= aRow['zip'].to_digit_str(5)
            else:
                zip= aRow['zip'].to_digit_str(9)
        else:
            zip= zip_str.replace('-'.'')
        return dict(
            name = aRow['variant column'].to_str(),
            address = arow['different column'].to_str(),
            zip = zip,
            phone = aRow['phone'].to_digit_str(),
        )
        
This shows filtering and cleasing operations.  Yes, it's complex.
Indeed, it's complex enough that alternative domain-specific languages (i.e., properties,
XLST, etc.) are much less clear than the Python.

**Stingray** Application Design
=================================

Generally, there are two kinds of testing.  Conventional unit testing
applies to our application to be sure the schemata
are used properly. Beyond that, we have data quality testing.

For application unit testing, our programs should be decomposed into three tiers of
processing.

-   Row-Level.  Inidividual Python objects built from one row of the input.
    This involves our builder functions.

-   Sheet-Level.  Collections of Python objects built from all rows of a sheet.
    This involves sheet processing functions.

-   Workbook-Level.  Collections of sheets.

Each of these tiers should be tested independently.

In :ref:`demo_sqa`, we'll look at how we 
validate that the the input files have the expected schema. This is a kind
of **Data Quality** testing. It can use the unit testing framework, but it applies
to data, not applications.

Row-Level Processing
----------------------

Row-level processing is centered on a suite of builder functions.
These handle the detailed mapping 
from variant logical layouts to a single, standardized conceptual schema.

A builder function should create a simple dictionary or :py:class:`types.SimpleNamespace`.
Each dictionary key is the standardized
attribute names used by internal Python class definitions.

**Q**.  Why not build the final Python objects from the source row?

**A**.  The source row needs to be validated to see if a valid object can be built.  
It seems simpler to map the logical layout in the source document to a 
standardized structure that matches the conceptual schema.  This standardized 
structure can be validated. Then the Python object built from that structure.

This follows the design patterns from the Django project where a ``ModelForm`` 
is used to validate data before attempting to build a ``Model`` instance.

Here's the three-part operation: **Build, Validate and Construct**.

..  parsed-literal::

    def builder_1( row ):
        return dict(
            *key* = row.cell( row.sheet.schema['field'] ).to_type(),
        )
        
    def is_valid( row_dict ):
        *All present or accounted for?*
        return *state*

    def construct_object( row_dict ):
        return App_Object( \*\*row_dict )

The validation rules rarely change. The object construction doesn't really
need to be a separate function, it can often be a simple class name. 

Our sheet processing can include a function like this. We'll look at this
below.

..  parsed-literal::

    builder= make_builder( args )
    for row in sheet:
        intermediate= builder( row )
        if is_valid(intermediate):
            yield construct_object(intermediate)
        else:
            log.error( row )

The builder, however, varies with the file's actual schema.
We need to pick the builder based on a "logical layout" command-line
option.  Something like the following is used to make an application 
flexible with respect to layout.

..  parsed-literal::

    def make_builder( args ):
        if args.layout in ("1", "D", "d"):
            return builder_1
        elif args.layout == "2":
            return builder_2
        else 
            raise Exception( "Unknown layout value: {0}".format(args.layout) )

The builders are tested individually.  They are subject to considerable change.
New builders are created frequently.

The validation should be common to all logical layouts.  
It's not subject to much variation.  
The validation and object construction doesn't have the change velocity that builders have.

Configuration Options
---------------------

We might want to package all builders in a separate module.
This provides a focused location for change that leaves the application
untouched when handling Yet Another Logical Layout.

..  parsed-literal::

    def make_builder( args ):
        builder_name = 'builder_{0}'.format( args.layout )
        globals = {}
        execfile( 'builders.py', globals )
        return globals[builder_name]
    
Or

..  parsed-literal::

    def make_builder( args ):
        import builders
        return eval('builders.builder_{0}'.format( args.layout ))

This allows a single application to be separated into invariant portions
and the builders which need to be tweaked when the source file layouts
change.

Sheet-Level Processing
------------------------

The next higher layer is sheet-level processing.  For the most part, 
sheets are generally rows of a single logcal type.  In exceptional cases,
a sheet may have multiple types coincedentally bound into a single sheet.
We'll return to the multiple-types-per-sheet issue below.

For the single-type-per-sheet, we have a processing function like
the following.

..  py:function:: process_sheet( sheet, builder )

    Process the given sheet using the given builder.

..  parsed-literal::
        
    def process_sheet( sheet, builder=builder_1 ):
        counts= defaultdict( int )
        object_iter = ( 
            builder(row))
            for row in sheet.schema.rows_as_dict_iter(sheet) )
        for source in object_iter:
            counts['read'] += 1
            if is_valid(source):
                counts['processed'] += 1
                # *The real processing*
                obj= make_app_object( source )
                obj.save()
            else:
                counts['rejected'] += 1
        return counts

This kind of sheet is tested two ways.  First, with a test fixture that provides
specific rows based on requirements, edge-cases and other "white-box" considerations.

It is also tested with "live-file".  The counts can be checked.  This actually
tests the file as much as it tests the sheet processing function.

Workbook Processing
---------------------

The overall processing of a given workbook input looks like this.

..  py:function:: process_workbook( source, builder )

    Process all sheets of the workbook using the given builder.

..  parsed-literal::

    def process_workbook( source, builder ):
        for name in source.sheets():
            # *Sheet filter?  Or multi-way elif switch?*
            sheet= source.sheet( name, 
                sheet.EmbeddedSchemaSheet,
                loader_class=schema.loader.HeadingRowSchemaLoader )
            counts= process_sheet( sheet, builder )
            pprint.pprint( counts )

This makes two claims about the workbook.

-   All sheets in the workbook have the same schema rules.
    In this example, it's an embedded schema in each sheet and the schema is the heading row.
    We could easily use an :class:`sheet.ExternalSchemaSheet` and an external schema.
    
-   A single :py:func:`process_sheet` function is appropriate for all sheets.

If a workbook doesn't meet these criteria, then a (potentially) more complex
workbook processing function is needed.  A sheet filter is usually necessary.

Sheet name filtering is also subject to the kind of change that
builders are subject to.  Each variant logical layout may also have
a variation in sheet names.  It helps to separate the sheet filter functions
in the same way builders are separated.   New functions are added with 
remarkable regularity

..  parsed-literal::
    
    def sheet_filter_1( name ):
        return re.match( r'*pattern*', name ) 

Or, perhaps something like this that uses a shell file-name pattern instead of a
more sophisticated regular expression. 

..  parsed-literal::
    
    def sheet_filter_2( name ):
        return fnmatch.fnmatch( name, '*pattern*' ) 

Command-Line Interface
----------------------

We have an optional argument for verbosity and a positional argument that
provides all the files to profile.

::

    def parse_args():
        parser= argparse.ArgumentParser()
        parser.add_argument( 'file', nargs='+' )
        parser.add_argument( '-l', '--layout' )
        parser.add_argument( '-v', '--verbose', dest='verbosity',
            default=logging.INFO, action='store_const', const=logging.DEBUG )
        return parser.parse_args()

The overall main program looks something like this.

::

    if __name__ == "__main__":
        logging.basicConfig( stream=sys.stderr )
        args= parse_args()
        logging.getLogger().setLevel( args.verbosity )
        builder= make_builder( args )
        try:
            for file in args:
                with workbook.open_workbook( input ) as source:
                    process_workbook( source, builder )
            status= 0
        except Exception as e:
            logging.exception( e )
            status= 3 
        logging.shutdown()
        sys.exit( status )
        
This main program switch allows us to test the various functions (:func:`process_workbook`, :func:`process_sheet`, the builders, etc.) in isolation.

It also allows us to reuse these functions to build larger (and more complete) 
applications from smaller components.

In :ref:`demo` we'll look at two demonstration applications, as well as a unit
test.


Variant Records and COBOL REDEFINES
====================================

Ideally, a data source is in "first normal form": all the rows are a single type
of data. We can apply a **Build, Validate, Construct** sequence simply.

In too many cases, a data source has multiple types of data. In COBOL files, it's common
to have header records or trailer records which are summaries of the details
sandwiched in the middle.

Similarly, a spreadsheet may be populated with summary rows that must be discarded or
handled separately. We might, for example, write the summary to a different destination 
and use it to confirm that all rows were properly processed.

Because of the COBOL REDEFINES clause, there may be invalid data. We can't assume that
all attributes have valid data. This makes our processing slightly different because
we can't necessarily do eager evaluation of each row of data.

We'll look at a number of techniques for handling variant records.

Trivial Filtering
------------------

When using an Embedded Schema Loader based on :py:class:`schema.loader.HeadingRowSchemaLoader`
we can extend this loader to reject rows. 

The :py:meth:`schema.loader.HeadingRowSchemaLoader.rows` method can do simple filtering.
This is most appropriate for excluding blank rows or summary rows from a spreadsheet.


Multiple Passes and Filters
----------------------------

When we have multiple data types within a single sheet, we can process this data
using Multiple Passes and Filters. Each pass uses different filters to separate the 
various types of data.

This relies on an eager production of an intermediate object from the raw data.
This may not work well for COBOL files.

The multiple-pass option looks like this.  Each pass applies a filter and 
then does the appropriate processing.

..  parsed-literal::
        
    def process_sheet_filter_1( sheet ):
        counts= defaultdict( int )
        object_iter = ( 
            builder(row))
            for row in sheet.schema.rows_as_dict_iter(sheet) )
        for source in object_iter:
            counts['read'] += 1
            if *filter_1(source)*\ :
                counts['filter_1'] += 1
                *processing_1(source)*
            else:
                counts['rejected'] += 1                
        return counts

Each filter is a simple boolean function like this.

..  parsed-literal::

    def filter_1( source ):
        return *some condition*
        
The conditions may be trivial: ``source['column'] == value``. The conditions
may be more complex. It's good to encapsulate them as distinct, named functions.
        
We could make the filter function and processing function an argument 
to a generic ``process_sheet()`` function. 
Sometimes this is a good simplification, sometimes it leads to extra 
complexity of little value.

One Pass and Switch
--------------------

When we have multiple data types within a single sheet,
We can make  single pass over the data, using an if-elif "switch"
to distinguish the different types of rows. Each type of row is
handled separately.

This relies on an eager production of an intermediate object from the raw data.
This may not work well for COBOL files.

The one-pass option looks like this.  A "switch" function is used to 
discriminate each kind of row that is found in the sheet.

..  parsed-literal::
        
    def process_sheet_switch( sheet ):
        counts= defaultdict( int )
        object_iter = ( 
            builder(row))
            for row in sheet.schema.rows_as_dict_iter(sheet) )
        for source in object_iter:
            counts['read'] += 1
            if *switch_1(source)*\ :
                *processing_1(source)*
                counts['switch_1'] += 1
            elif *switch_2(source)*\ :
                *processing_2(source)*
                counts['switch_2'] += 1
            *elif etc.*
            else:
                counts['rejected'] += 1                
        return counts

Each switch function is a simple boolean function like this.

..  parsed-literal::

    def switch_1( source ):
        return *some condition*
        
The conditions may be trivial: ``source['column'] == value``. The conditions
may be more complex. It's good to encapsulate them as distinct, named functions.

We may be able to build a simple mapping from switch function to process function.

..  parsed-literal::

    switch_process = ( 
        (*switch_1*, *processing_1*),
        (*switch_2*, *processing_2*),
    )
    
This allows us to write a generic sheet processing function.

..  parsed-literal::
        
    def process_sheet_switch( sheet, mapping ):
        counts= defaultdict( int )
        object_iter = ( 
            builder(row))
            for row in sheet.schema.rows_as_dict_iter(sheet) )
        for source in object_iter:
            counts['read'] += 1
            processed= None
            for switch, process in mapping:
                if switch(source):
                    processed= switch.__name__
                    process( source )
                    counts[processed] += 1
            if not processed:
                counts['rejected'] += 1                
        return counts

This can more easily be extended by adding to the ``switch_process`` mapping.

Lazy Switch Processing
-----------------------

The above two examples rely on building an iterator over intermediate objects.
The ``object_iter`` builds objects eagerly.
This may not always work for COBOL files. Here's a variation that might be helpful.

We'll decompose the builders so that the switch is applied first. Then the 
builder and processing can depend on the switch.

..  parsed-literal::

    switch_build_process = ( 
        (*switch_1*, *builder_1*, *processing_1*),
        (*switch_2*, *builder_2*, *processing_2*),
    )

This structure can be used with the following generic sheet processing.

..  parsed-literal::
        
    def process_sheet_switch( sheet, mapping ):
        counts= defaultdict( int )
        for row in sheet.schema.rows(sheet):
            counts['read'] += 1
            processed= None
            for switch, builder, process in mapping:
                if switch(row):
                    processed= switch.__name__
                    source= builder( row )
                    process( source )
                    counts[processed] += 1
            if not processed:
                counts['rejected'] += 1                
        return counts

This is slightly more complex. It the advantage of not attempting to process
a row unless some initial sanity check has been done. Once the switch function
determines the type of the row, then an appropriate builder can be invoked
and the row processed.

In many cases, the processing starts with more complex data quality validation.
If so, that can be added to the mapping. It would become a
switch-builder-validator-process mapping that decomposes each step of the 
processing pipeline.

Big Data Performance
=====================

We've broken our processing down into separate steps which
work with generic Python data structures. The idea is that we can use
multiprocessing to spread the pipeline into separate processors or cores.

Each stage of the **Build, Validate, Construct** sequence can be decomposed.
We can read raw data from the source file, apply a switch and put the 
raw "Row" objects into a processing queue.

A builder process can dequeue row objects from the processing queue, 
apply a builder, and put objects into a validation queue.

A validator process can dequeue built objects (dictionaries, for example) and
validate them. Invalid objects can be written to a queue for logging.
Valid objects can be written to another queue for processing.

The final processing queue will get a sequence of valid objects all of a single type.
The final processing might involve (slow) database transactions, and there
may need to be multiple worker processes fetching from this queue.

File Naming and External Schema
===============================

Some data management discipline is also essential be sure that the schema and file match
up properly.  Naming conventions and standardized directory structures are
*essential* for working with external schema. 

Well Known Formats
--------------------

For well-known physical formats (:file:`.csv`, :file:`.xls`, :file:`.xlsx`, :file:`.xlsm`, :file:`.ods`,
:file:`.numbers`) the filename extension describes the physical format. Additional
information is required to determine the Logical Layout.

The schema may be loaded from column headers, in which case the binding is handled 
via an embedded schema loader. If the  :py:class:`schema.loader.HeadingRowSchemaLoader`
is used, no more information is required. If a customized schema loader is used
(because the headings are not trivially the first row of a sheet), then we must
-- somehow -- bind application to external schema. The filename extension doesn't 
really help with this. The best choice is to use a configuration file of some kind.

If the schema is external, and we're working with a well-known physical format, then
we have an even more complex binding issue. The schema will often require a customized
schema loader. Each file must be associated with a schema file and a schema loader
class name. File naming conventions won't help. This, too, should rely on a configuration
file.

Fixed Formats and COBOL
------------------------

For fixed-format files,
the filename extension does not describe the physical layout.
Further, a fixed format schema must combine logical layout and physical format into
a single description. 

For fixed format files, the following conventions help
bind a file to its schema.

-   The data file extension is the base name of a schema file. 
    :file:`mydata.someschema`. Do not use ``.dat`` or something else uninformative.

-   Schema files must be be either a DDE file or a 
    workbook in a well-known format.
    :samp:`someschema.cob` or :file:`someschema.xlsx`.
    
**Examples**.  We might see the following file names.

.. parsed-literal::

    september_2001.exchange_1
    november_2011.some_dde_name
    october_2011.some_dde_name
    exchange_1.xls
    some_dde_name.cob
    
The ``september_2001.exchange_1`` file is a fixed format file 
which requires the ``exchange_1.xls`` metadata workbook.

The ``november_2011.some_dde_name`` and ``october_2011.some_dde_name`` files
are fixed format files which require the ``some_dde_name.cob`` metadata.

External Schema Workbooks
-------------------------

A workbook with an external schema sheet must adhere to a few conventions to be usable.
These rules form the basis for the :py:class:`schema.loader.BareExternalSchemaLoader`
class. To change the rules, extend that class.

-   The standard sheet name ``"Schema"``  defines the appropriate sheet.

-   There must be an internal meta-schema on line one of the sheet that provides the expected column names.

-   The column names "name", "offset", "size", "type" are used.

-   Only "name" is required  in general.

-   For fixed format files, "offset", "size" and "type" will also be required.

-   Additional columns are allowed, but will be ignored.

-   Type definitions are "text", "number", "date" and "boolean".  

Binding a Schema to an Application
====================================

We would like to be sure that our application's expectations for a
schema are aligned with the schema actually present.  
An application has several ways to bind its schema information.

-   **Implicitly**.  The code simply mentions specific columns
    (either by name or position). 
    
-   **Explicitly**.   The code has a formal "requires" check to be sure
    that the schema provided by the input file actually matches the 
    schema required by the application.

Explicit schema binding parallels the configuration management issue of module
dependency. A file can be said to *provide* a given schema and an
application *requires* a given schema.

Sadly, we don't always have a pithy summary of a schema.  We can't trivially examine
a file to be sure it conforms to a schema. In the case of well-known file formats with an
embedded schema, we can do a test like this to determine if the schema is what we expect.

..  parsed-literal::

    all( req in schema for req in ('some', 'list', 'of', 'columns') )
    
This covers 80% of the use cases. For all other cases, we don't have a reliable way
to confirm the file's schema.

If we don't implement this, 
we're left with implicit schema binding in our applications.  In short, we have
to include data validity checks, a debugging log, and some kind of warning technique.

Schema Version Numbering
=================================

XSD's can have version numbers.  This is a very cool.

See http://www.xfront.com/Versioning.pdf for detailed discussion of how
to represent schema version information.

Databases, however, lack version numbering in the schema.  This leads to potential
compatibilty issues between application programs that expect version 3 of the
schema and an older database that implements version 2 of the schema.

Our file schema, similarly, don't have a tidy, unambiguous numbering.

For external schema, we can embed the version in the file names.
We might want to use something like this ``econometrics_vendor_1.2``.
This identifies the 
generic type of data, the source for that file, and the schema version
number. 

    Within a SQL database, we can easily use the schema name to carry
    version information.  We could have a :samp:`name_{version}` kind of
    convention for all schema, allowing an application to confirm schema
    compatibility with a trivial SQL query.

For embedded schema, however, we have no *easy* to handle schema identification
and version numbering.  We're forced to 
build an algorithm to examine the actual names in the embedded schema to deduce
the version.  

This problem with embedded schema leads to using data profiling to reason out what the file is.  
This may devolve to a manual examination
of the data profiling results to allow a human to determine the schema.
Then, once the schema has been identified, command-line options
can be used to bind the schema to file for correct processing.

Frequently Asked Questions
==========================

Junk Data
----------

For inexplicable reasons, we can wind up with files that are damaged in some way.

    "there is a 65-byte "header" at the start of the file, what would be the best 
    (least hacky) way to skip over the first 65 bytes?"
    
This is one of the reasons why use both a file name and an open file object as
arguments for opening a workbook.

..  parsed-literal::

    with open("file_with_junk.some_schema","rb") as cobol:
        cobol.seek( 66 )
        wb = stingray.cobol.EBCDIC_File( cobol.filename, file_object=cobol )
        
This should permit skipping past the junk.
