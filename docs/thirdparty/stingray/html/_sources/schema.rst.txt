..  #!/usr/bin/env python3

..  _`schema`:

########################################################
Schema Package -- Schema and Attribute Definitions
########################################################

..  py:module:: schema

A *Schema* contains data descriptions.  It's a collection of *Attribute* specifications.

For our purposes, we're trying to cover the following bases:

-   Workbooks in a variety of forms: CSV, Tab, XLSX, ODS, etc.
    The common feature of these schema is that they're flat.  A single
    index (or name) identifies each column. 
   
    The schema can be in a variety of places.

    -   The schema may be a header row within a sheet.

    -   The schema may be elsewhere in the sheet.  It may be a header
        that's not the first row.

    -   The schema may be in a separate document.

-   COBOL Files. The schema is highly structured: a path must be 
    used to unambiguously identify each field. Some COBOL schema are
    designed so that all leaf names are unique.

    The schema is a separate document, encoded in
    COBOL source code.
   
We'll also happen to cover relational database table definitions.  However,
this isn't our focus.  This is simply coincidence.

Note that we are not trying to cover XML schemas or a **complete** relational
database schema. There are a lot of possible features not present here.

Load a Schema Use Case
===============================

The objective of this module is to provide a handy base class family
that can be used to load schema information from any of a variety
of sources.

-   Embedded schema in a sheet of a workbook.  Either a header row
    or some more complex parsing.  This implies some kind of schema
    parser that reads a workbook sheet and gets little more than
    column names and ordinal positions.

-   External schema in another workbook.  This implies a
    parser that gathers some information (name, type, position, offset,
    or size) from a sheet using a fixed schema.

-   External schema in COBOL.  This implies a parser for COBOL source.

Embedded schema loading would look like this.

..  parsed-literal::

    with *open* as wb:
        counts= defauldict( int )
        sheet = EmbeddedSchemaSheet( workbook, 'Sheet1', HeadingRowSchemaLoader )
        counts= process_sheet( sheet )
        pprint.pprint( counts )

The schema object isn't made explicit.  It's available in ``sheet.schema``.

External schema loading would look like this.

..  parsed-literal::

    with *open schema* as wb_s:
        esl = ExternalSchemaLoader( wb_s, 'Schema' )
        schema = esl.load()
    with *open data* as wb:
        sheet = ExternalSchemaSheet( wb, 'Sheet1', schema )
        counts= process_sheet( sheet )
        pprint.pprint( counts )

Use a Schema Use Case
==============================

We use a schema to access fields by name. Here's the use case from :ref:`intro`.

..  parsed-literal::

    foo= row['foo'].to_str()

    bar= row['bar'].to_float()

This, of course, only works in the cases where the field name
is unique and the row's values can be built eagerly.  The good news
is that 80% of the time this is true.  The other 20% of the time, we
need something more complex.

For the 80% case, we can do this.

..  parsed-literal::

    for row_seq in sheet.rows():
        row= dict(
            (a.name, row_seq.cell(a))
            for a in schema
        )
        foo= row['foo'].to_str()
        bar= row['bar'].to_float()

Or this.

..  parsed-literal::

    for row in schema.rows_as_dict_iter( sheet.rows() ):
        foo= row['foo'].to_str()
        bar= row['bar'].to_float()

In the 20% case, we can't build a row eagerly.  In this case, we have to do
the following to fetch cell values using a properly lazy row.

..  parsed-literal::

    schema_dict= dict( (a.name, a) for a in schema )
    for row_seq in sheet.rows():
        foo= row.cell( schema_dict['foo'] ).to_str()
        bar= row.cell( schema_dict['bar'] ).to_float()

This involves a three-step dance because a row (and a schema) have a number of ambiguities.
In particular, names may be duplicated, forcing us to use position or more  complex naming conventions.  Also, the attribute may have repeating groups,
requiring some indexing, as well as naming.

1.  Build a schema mapping by some name.  We can use the attribute name (if it's unique)
    or (for COBOL schema) a unique path name.  Or, we may have some other, more
    complex naming for attributes.  Really.

2.  Find the attribute in our mapping.

3.  Find the cell value based on the attribute.

Application Object Builder
==========================

The whole point of this is to facilitate writing applications that work
with objects encoded as rows in a spreadsheet. We're not merely picking
a value from cells of a row: we're creating Python objects from the row
as a whole. These application objects are also based on a common Conceptual Schema;
they are yet another implementation of the concept.

Experience indicates that the conceptual schema may have several variant
logical layouts.  Because of the logical layout variability,
a two-step dance is
required between source rows and final Python objects.

1.  Transform input row to a standardized dictionary.  The mapping
    from input layout to dictionary changes.  Frequently.  The dictionary
    matches the conceptual schema.  The input is one of the variant
    logical layouts.

2.  Create the application object from the standardized dictionary.

..  parsed-literal::

    def build_dict_1( aRow ):
        return dict(
            attribute= aRow['column'].to_str(),
            another= aRow['data'].to_float(),
        )

    def make_app_object( aDict ):
        return Object( \*\*aDict )

    def process_sheet( sheet, builder=build_dict_1 ):
        counts= defaultdict( int )
        object_iter = (
            make_app_object(builder(row))
            for row in sheet.schema.rows_as_dict_iter(sheet.rows()) )
        for obj in object_iter:
            *process object*
        return counts

This allows us to configure an appropriate builder function depending
on which variation of the logical layout the file actually has.

The :py:mod:`csv` module offers trivial support for an eager "row-as-dict"
processing. This can actually introduce problems.  Specifically, COBOL
record layouts may have a number of fields named ``Filler``.  Really.
Also, it's common to get data where column names are duplicated in the
embedded schema.  Finally COBOL redefines means that lazy construction of
Cell instances is more appropriate.

Data conversions can become a :math:`n \times m` issue.  Each of
:math:`n` input types can be mapped to each of :math:`m` output types.  In this instance, we try to keep it to :math:`n + m`.  We do this by
acquiring a Python-specific type from the source.

-   For fixed format files, this may involve decoding characters (not bytes).

-   For COBOL format files, this may involve decoding bytes if EBCDIC
    and COMP-3 data are involved.  It may simply involve decoding characters
    if the file happens to be encoded into proper characters.

-   For all other physical formats (CSV, XLS, XSLX, etc.) there is no separate
    decoding.   There are about five canonical cell types (mostly floats) with
    decodings defined by the format.

Byte to Character conversion is not part of the schema problem.  That's part
of the physical format.  For the most part, the physical format defines the
encoding of the bytes.  COBOL files in Unicode (or ASCII) strings,
require standard, default decoding.  COBOL files in EBCDIC, however, may require
highly customized decoding.  This will be delegated to that module.

Output conversions are not part of the schema problem. They're part of
the application.  All this module does is get a workbook :py:class:`cell.Cell` instance.

It leads us builder functions that might look like this.

..  parsed-literal::

    def build_dict_2( aRow ):
        if aRow['flag'].to_str() == "C":
            value= math.pi*aRow['r'].to_float()**2
        else:
            value= aRow['l'].to_float()*aRow['w'].to_float()
        return dict(
            attribute= aRow['column'].to_str(),
            another= aRow['data'].to_float(),
            value= value,
        )

We've applied a number of input transformations to create Python objects. These
objects are designed to fit as closely as possible to the conceptual schema.
We can then output these in any required format.

Model
======

..  code-block:: none

    http://yuml.me/diagram/scruffy;/class/
    #schema,
    [Schema]<>-[Attribute],
    [Attribute]->[TextCell],
    [Workbook]->[Attribute].

..  image:: schema.png


Overheads
===============

A schema depends only on the definitions in :py:mod:`cell`.

::

  """stingray.schema -- Defines an overall Schema of Attributes which can be
  embedded as a row of a sheet or encoded in a separate sheet.  A
  schema defines the logical layout of a workbook or flat file.
  """

  import datetime
  import time
  import stingray.cell

Schema Class
=============

..  py:class:: Schema

    The core Schema definition is an extension to ``list``. In addition to 
    a sequence of attributes, it also has an "info" object that's a dictionary
    of additional keywords.

    The :py:meth:`schema.Schema.rows_as_dict_iter` method uses the sheet's
    :py:meth:`sheet.Sheet.rows` iterator to create simple row-as-list values.
    These are transformed into the row-as-dict values.  If the attribute names
    involve duplicates, then one of the duplicated values will be chosen; the
    choice is arbitrary.

    ..  py:attribute:: info
   
        Dict of additional information about this schema. Meta-metadata.
        For COBOL schema, this includes the source DDE.
   
    ..  py:attribute:: names
   
        Attribute names for rows_as_dict_iter()

::

  class Schema( list ):
      """A Mutable Sequence of attributes.  Order matters.
      """
      def __init__( self, *attr, **kw ):
          """Build a schema from collection of attributes."""
          super().__init__( attr )
          for p, a in enumerate( self ):
              a.position= p
          self.info= kw
      def __repr__( self ):
          attr_list= map( repr, self )
          return "Schema( {0} )".format( ", ".join(attr_list) )
      def rows_as_dict_iter( self, sheet ):
          self.names= tuple(a.name for a in self)
          for r in sheet.rows():
              yield dict(
                  (a.name, r.cell(a)) for a in sheet.schema )
      def append( self, child ):
          child.position= len(self)
          super().append( child )

Possibly helpful method to expand a row based on the schema information.

::

      def expand( self, aRow ):
          """Expand each attribute to create a dictionary of cells."""
          return dict( (attr.name, aRow.cell(attr)) for attr in self )
        
For parsing COBOL data, we often need to know the total length of the defined schema.
This only works for records without an Occurs Depending On.

::

      def lrecl( self ):
          return max( a.offset + a.size for a in self )

A Schema needs to handle two common use cases.

-   Cell values are defined by the physical format. Data can be fetched positionally.
    Names map to positions.

-   Fixed and COBOL.  The cell values are not defined by the physical format, but by an external
    schema associated with the :py:class:`sheet.Sheet`. Names map to offsets and sizes;
    the bytes found there must be converted to Python objects. In the case of 
    Occurs Depending On (ODO), the offsets depend on both schema and data.
   
    COBOL data may have elements which are invalid, but unused due to application
    logic in selecting a proper REDEFINES alias.

The simple positional schema isn't really appropriate for all purposes.
For COBOL and fixed format files with external schema, we often
must process things lazily by field name.

This is unlike spreadsheets where we can process all fields eagerly and in order.
   
..  todo:: Index by name and path, also.

    This will eliminate some complexity in COBOL schema handling where
    we create the a "schema dictionary" using simple names and path names.

Attribute Class
=================

..  py:class:: Attribute

    An Attribute definition has a required value of a name and a class that will
    be created to hold the data.

    Here are the essential attributes of an Attribute.

    ..  py:attribute:: name
   
        The attribute name. Typically always available for most kinds of schema.

    ..  py:attribute:: create
   
        Cell class to create.  If omitted, the class-level
        :py:data:`Attribute.default_cell` will be used.
        By default, this refers to :py:class:`stingray.cell.TextCell`.

    The additional
    values commonly provided by simple fixed format file schemata.

    ..  py:attribute:: offset
   
        Optional offset into a buffer. For simple fixed-layout files,
        this is a constant. For COBOL files with Occurs Depending On,
        however, this must be a function based on the actual record
        being processed.

    ..  py:attribute:: size
   
        Optional size within the buffer.

    ..  py:attribute:: position
   
        Optional sequential position.

    A subclass might introduce yet more attributes.

::

  class Attribute:
      """Essential definition of a single source data element."""
      default_cell= stingray.cell.TextCell
      def __init__(self, name, offset=None, size=None, create=None, position=None, **kw):
          """Build an Attribute.
          :param name: The attribute name.
          :param offset: Optional offset into a buffer.
          :param size: Optional size within the buffer.
          :param create: Cell class to create.  If omitted, the class-level
              :py:data:`Attribute.default_cell` will be used.
          :param position: Optional sequential position.
          """
          self.name, self.offset, self.size, self.create, self.position =\
          name, offset, size, create, position
          if not self.create:
              self.create= self.default_cell
          self.__dict__.update( kw )
      def __repr__( self ):
          return "Attribute( name={0.name!r}, position={0.position}, offset={0.offset}, size={0.size} )".format( self )

An :py:class:`schema.Attribute` is used by a :py:class:`workbook.base.Workbook` to
extract cell data from a row.

The use case looks like this for a Fixed format workbook.  For other
workbooks, other kinds of conversion functions might be used.

..  parsed-literal::

    def cell( sheet, attribute, data ):
        a= attribute
        return a.create( data[a.offset:a.offset+a.size], sheet.workbook )

The attribute might be declared as follows.

..  parsed-literal::

    Attribute( name= "mm-dd-yy", size= *n*, offset= *m*,
        create=SomeCellSubclass )
