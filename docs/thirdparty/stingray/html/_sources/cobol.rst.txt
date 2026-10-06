.. _`cobol`:

######################################
The COBOL Package
######################################

If you're interested in the implementation, you can cut to the chase.
See :ref:`cobol_impl`.

-   :ref:`cobol_init`.

-   :ref:`cobol_loader`.

-   :ref:`cobol_defs`.

This section answers the following question:

    "How do I access COBOL-defined data from Python"?

We have three problems to solve to get COBOL data into Python applications.

-   How do we represent the data?

    Most of the :py:class:`cell.Cell` subclasses aren't really appropriate for    
    data coming from COBOL applications.   
    Indeed, only :py:class:`cell.TextCell` is really appropriate.
    
    Given that all raw data is bytes -- often in EBCDIC -- how do we understand what those 
    bytes represent? For this, we'll need a more complex schema.
    
-   How do we represent this more complex schema?

    The schema more complex than the 
    flat-file schema expected by :py:mod:`schema`.
    
    -   The structure is hierarchical.

    -   Each attribute has a large number of properties required for proper 
        conversion to Python types.

    -   There are repeating groups (based on the "OCCURS" clauses).

    -   There are attributes with locations based on other attributes.
        These are defined by Occurs Depending On clauses.
        
    -   There are fields that may have invalid data due to a "REDEFINES" clause.
        The USAGE information indicates the encoding of the data which is 
        expected to be in the field, the actual data may not match the definition.
        
    We'll need to load this more complex schema.
    
-   How do we build the schema?

    The schema is encoded in COBOL.  That is the subject of :ref:`cobol_loader`.
    
Once we can load a schema, we can then use the schema to determine what a sequence
of bytes represent. Once we know what the bytes represents, we can work with COBOL
data.

Requirements
==============

The essential COBOL use case is to create usable Python objects from the source file data.

For each source file row, there's a two-step operation.

1.  Access elements of each row using the COBOL DDE structure.
    
2.  Build Python objects from the Cells found in the row.  
    Building Python objects is best done with a "builder" function, 
    as shown above in :ref:`schema`, :ref:`developer`, and :ref:`demo`.
    
This isn't all. In order to locate the rows, we may have to understand overall
file organization. This includes the following feature:
    
-   Split the file into sections for parallel processing.
    The GNU/Linux ``split`` command won't work with EBCDIC files, so
    we have to use the low-level RECFM definitions to parse and split
    a file.

High-Level Processing
----------------------

A :py:class:`sheet.Row` appears to be a sequential collection of 
:py:class:`cell.Cell` instances.  The schema is 
used for **lazy**  creation of :py:class:`cell.Cell` instances.

A candidate cells may be based on bytes that aren't valid for the
cell's detailed description. This is common when there is a `REDEFINES`. 
This means that a row cannot be built eagerly.

Here's our example code. This is generally how we want to proceed:

..  parsed-literal::

    with open("sample/zipcty.cob", "r") as cobol:
        schema= stingray.cobol.loader.COBOLSchemaLoader( cobol ).load()
        #pprint.pprint( schema )
    for filename in 'sample/zipcty1', 'sample/zipcty2':
        with stingray.cobol.Character_File( filename, schema=schema ) as wb:
            sheet= wb.sheet( filename )
            counts= process_sheet( sheet )
            pprint.pprint( counts )
            
We want to load a schema. Then we want to open and process a file based on this
schema. This parallels an :py:class:`sheet.ExternalSchemaSheet`, with an 
:py:class:`schema.loader.ExternalSchemaLoader`.

The :py:class:`cobol.Character_File` class is for files in which all fields
are  all character encoded in ASCII. 
No packed decimal. No EBCDIC.
These kinds of files are expected to have proper ``'\n'`` characters at the end of each record.

For an EBCDIC file, use the :py:class:`cobol.EBCDIC_File` class. These files 
lack a specific record delimiter character. If the "RECFM" (Record Format) is
V or VB, there's a header word with length instead of a delimiter.
        
Processing a COBOL "sheet" is a slight extension to processing workbook sheets.
We can turned the simple sequential COBOL-based schema into a dictionary. 
The keys are the simple field names as well as the full field paths.


..  parsed-literal::

    def process_sheet( sheet ):
        # Simple field names
        schema_dict= dict( (a.name, a) for a in sheet.schema )
        # Dot-puncutated field paths
        schema_dict.update( dict( (a.path, a) for a in sheet.schema ) )
        
        counts= { 'read': 0 }
        row_iter= sheet.rows()
        row= next(row_iter)
        header= header_builder( row, schema_dict )
        print( header )
        
        for row in row_iter:
            data= row_builder( row, schema_dict )
            print( data )
            counts['read'] += 1
        return counts
        
This relies on two functions to access elements: ``header_builder()`` and ``row_builder()``.
Each of these will use the dictionary schema to access individual elements.

Access Elements Via Schema
------------------------------

A function like the following builds an object from a row using the schema.

..  parsed-literal::

    def header_builder(row, schema):
        return types.SimpleNamespace( 
            file_version_year= row.cell(schema['FILE-VERSION-YEAR']).to_str(),
            file_version_month= row.cell(schema['FILE-VERSION-MONTH']).to_str(),
            copyright_symbol= row.cell(schema['COPYRIGHT-SYMBOL']).to_str(),
            tape_sequence_no= row.cell(schema['TAPE-SEQUENCE-NO']).to_str(),
        )

We use ``schema['FILE-VERSION-YEAR']`` to track down the attribute details.
We use this information to build a ``Cell`` from the current row.

We also need to handle repeating groups. 
In the case of repeating elements,  the repeating collection of elements
creates a tuple-of-tuples structure which we can index.

This gives us two possible ways to provide an OCCURS index values.  The first 
possibility is to associate the index with the :py:class:`schema.Attribute`.

..  parsed-literal::

    def process_sheet( sheet ):
        for row in sheet.rows():
            foo= row.cell(schema[*n*]).to_str()
            bar= row.cell(schema[*m*].index(*i*, *j*)).to_str()

While the above is weird, it parallels COBOL usage of indexing and element naming.  
The idea is to provide the lowest level of name and all the indexes
above it in one construct.

This is appealing because we can use the existing :py:meth:`sheet.Row.cell` method without
any modification. The downside of this is that we're creating a tweaked attribute
definition, which is a terrible idea. 
On balance, the use of :py:meth:`schema.Attribute.index` has to be rejected.

The second way to provide an index is provide a subset of the indexes and get a kind of slice.
This is like the positional version, shown above.

..  parsed-literal::

    def process_sheet( sheet ):
        for row in sheet.rows():
            foo= row.cell(schema[*n*]).to_str()
            bar= row.cell(schema[*m*])[*i*][*j*].to_str()
            
Handling 01-Level Subrecords
---------------------------------------

In some COBOL files, there can be 01-level "subrecords" buried within an 01-level record.

A programming technique seen in the wild is this.

..  parsed-literal::

    01 GENERIC-RECORD.
       05 HEADER PIC X(3).
       05 GENERIC-FIELD PIC X(17).
       
    01 ABC-SPECIFIC-RECORD.
       05 ITEM-1 PIC X(10).
       05 ITEM-2 PIC X(7).
       
    01 DEF-ANOTHER-RECORD.
       05 ITEM-3 PIC X(7).
       05 ITEM-4 PIC X(10).
       
The original COBOL program likely did something like this to make the two
other 01-level items work as sub-records to a parent 01-level.

..  parsed-literal::

     IF HEADER EQUALS "ABC"
     THEN
          MOVE GENERIC-FIELD TO ABC-SPECIFIC-RECORD
          *process this record*
     ELSE IF HEADER EQUALS "DEF"
          MOVE GENERIC-FIELD TO DEF-ANOTHER-RECORD
          *process this record*
          
How do we work with multiple :py:class:`sheet.Row` objects here?

-   We have the ``GENERIC-RECORD`` created by a COBOL workbook bound to 
    a schema for this record.
    
    For this, ``row.cell( schema_dict['HEADER'] )`` fetches the value.
    
-   We have the ``ABC-SPECIFIC-RECORD`` and the ``DEF-ANOTHER-RECORD`` 
    bound to a subset of bytes within the source record.
    
    This case requires Python code which vaguely parallels the COBOL 
    code. 
    In effect, we're creating a new :py:class:`cobol.ODO_LazyRow` object from
    ``buffer= row.cell( schema_dict['GENERIC-FIELD'] ).raw`` and the
    proper schema for that variant.

    ..  parsed-literal::

        segment = ODO_LazyRow(
            ExternalSchemaSheet( workbook, "ABC-SPECIFIC-RECORD", subschema ),
            data= row.cell( schema_dict['GENERIC-FIELD'] ).raw,
        )
    
This is implemented as the :py:meth:`cobol.COBOL_File.subrow` method.
It allows us to work with a single field (``schema_dict['GENERIC-FIELD']``) 
as if it was a row with a given schema (``subschema``).
       
Possible Added Fluency
----------------------

There are several places where extra fluency might help.

-   Complex nested COBOL layouts with duplicated bottom-level names.

-   Complex indexing. Do we handle this in COBOL or Python syntax?

In the case of duplicated bottom-level names, we are forced to use 
the COBOL schema path names. These names are awkward to use since they must include every level
from the ``01`` to the relevant item.

A subclass of :py:class:`schema.Schema` *could* introduce
a method like ``get_name()`` to make a schema into a slightly more fluent
mapping.

..  parsed-literal::

            foo= row.cell(schema_dict.get_name('bar'))
            baz_i_j= row.cell(schema_dict.get_name('baz').index(*i*, *j*))

While handy for intermediate names, it still wouldn't work for non-unique bottom-level names. 

The COBOL ``OF`` syntax *could* be modeled using a fluent ``of()`` method.

..  parsed-literal::

            foo= row.cell(schema_dict.get('foo').of('bar'))
            baz_i_j= row.cell(schema_dict.get('baz').index(*i*, *j*))
            baz_i_quux_j = row.cell(schema_dict.get('baz').index(*i*).of('quux').index(*j*))

This is potentially useful for hyper-complex COBOL record layouts.

Complex indices can be handled in COBOL syntax creating just a single Cell value.
Or they can be handled by Python syntax by creating a Python structure of multiple Cell
values.  This leads us to use partial and incremental index calculation.

Incremental index calculation involves creating an interim, stateful Attribute definition.
A partial index efficiently implemented as stateful intermediate objects. We can accumulate
index information incrementally.

..  parsed-literal::

            foo= row.cell(schema_dict.get('foo').of('bar'))
            baz_i_j= row.cell(schema_dict.get('baz'), *i*)
            baz_i_quux_j = row.cell(schema_dict.get('baz').of('quux'), (*i*, *j*))
            
The first example provides no index, a multi-dimensional sequence of `Cell` objects is returned.

The second example provides too few indices, a sequence of `Cell` objects is returned.

The third example provides all indices, an individual `Cell` is returned.

Generic Processing
---------------------

The :py:func:`cobol.dump` function can dump a record showing
raw bytes. 

This relies on the :py:func:`cobol.dump_iter` function. This function
iterates through all DDE elements (in order) providing a five-tuple of the 
original DDE, the derived attribute, a tuple of specific index values used, the
raw bytes found and any ``Cell`` object if the data was actually valid.

The indices tuple may be an empty tuple for non-OCCURS data. For OCCURS data, 
all combinations of index values are used, so large, nested structures may produce
long lists of values.

The raw bytes and the ``Cell`` object are (technically) redundant, since all
subclasses of ``Cell`` used by the ``cobol`` package have a ``raw`` attribute with
the raw bytes. However, it's sometimes simpler to have this expanded in the tuple.

Foundational Definitions
---------------------------------

There are two kinds of DDE's: group and elementary. A group DDE is a collection of DDE's. 
An elementary does not contain anything; it is a leaf. This makes a DDE a proper tree.

A DDE, :math:`d`, has two interesting properties: the offset, :math:`\operatorname{o}(d)`, and the
size, :math:`\operatorname{s}(d)`. We're interested in the bytes associated with a particular DDE 
element, :math:`B[d]`. These bytes are fetched from a larger buffer, :math:`B = \{ B_0, B_1, B_2, ..., B_n \}`.

This allows us to fetch the bytes from a buffer by getting
bytes :math:`B[d] = \{B_x | \operatorname{o}(d) \leq x < \operatorname{o}(d)+\operatorname{s}(d)\}`.

The size of an elementary DDE, :math:`\operatorname{s}(d)`, is fixed by the picture clause.

The size of a group DDE is the sum of the children. :math:`\operatorname{s}(d) = \sum_{c \in d}\operatorname{s}(c)`.
While this seems clear, it doesn't account for the ``OCCURS`` clause issues: we'll return to
those below.

The DDE's in a tree can have three optional species of relationships.

-   **Predecessor/Successor**. The predecessor of :math:`d` is :math:`P(d)`.
    The offset of an item is :math:`\operatorname{o}(d) = \operatorname{o}(P(d)) + \operatorname{s}(P(d))`. 
    This applies recursively
    to the first item in the DDE collection.
    For the very first item in a DDE, :math:`d_0`, :math:`\operatorname{o}(d_0)=0`.
    For the first item in a group, however, the rule for the offset is based on the parent.
    
    Also, for the very first item, and the first item in a group, there is no 
    predecessor. We can say :math:`P(d_0)=\bot` to indicate it has a null value.
    Or we can say :math:`\nexists x \vert x = P(d_0)`, to indicate that there's 
    no such value, it's an exception to ask for it.

-   **Group/Elementary** or Parent/Child. The parent is a group of items; the parent of :math:`d` is :math:`G(d)`,
    we can say :math:`d \in G(d)`. For top-level items in the DDE, there is  no 
    parent group.  We can say  :math:`G(d_0)=\bot` for null or 
    :math:`\nexists x \vert x = G(d_0)` to show that it's an exception to ask.
    
    The offsets inside this group depend on the 
    base offset of the group :math:`\operatorname{o}(G(d)) \Rightarrow \operatorname{o}(d)`.
    This is a vague "influences" relationship. We can be more formal by noting
    that :math:`G(d) = \{ d_0, d_1, d_2, ..., d, ..., d_n \}`. This means that
    :math:`\operatorname{o}(d_0) = \operatorname{o}(G(d))`, and we can use the :math:`P(d)`
    predecessor to locate the cumulative offset between each :math:`d_i` and :math:`d_0`,
    including the cumulative offset between :math:`d` and :math:`d_0`.

-   **Redefines**. This DDE's offset (and size) of :math:`d` is defined by another item, :math:`R(d)`.
    This, too, is optional. We can say  :math:`R(d_0)=\bot` for null or 
    :math:`\nexists x \vert x = R(d_0)` to show that it's an exception to ask.
    
    The offset, :math:`\operatorname{o}(d) = \operatorname{o}(R(d))`. This often applies
    at a group level, which means that the :math:`\operatorname{o}(G(d)) = \operatorname{o}(R(G(d)))`.
    A group redefines influences the offset of all items in the group
    :math:`\operatorname{o}(R(G(d))) \Rightarrow \{\operatorname{o}(x) \in G(d)\}`. 

The predecessor/successor relationship is specified by the order of the DDE's as they're 
compiled. Every item after the first has a predecessor. 

The child relationship is specified via the level numbers. A larger level number
implies a child of a lower level number. If they have the same level number, they're successors.

The redefines relationship is specified via the ``REDEFINES`` clause. 

Occurs Depending On
---------------------------------

The previous definitions aret complicated by the ``OCCURS`` clause. 
There are two versions in the COBOL language
manual, to which we'll add a third.

-   "Format 1" has a fixed number of occurrences: :math:`O(d) = n`, comes from ``OCCURS n TIMES``.

-   "Format 2" means the number of occurrences depends on a piece of data, 
    in another DDE, the depends-n basis, :math:`D(d)`. This means that
    :math:`O(d) = B[ D(d) ]`. This comes from ``OCCURS n TO m TIMES DEPENDING ON b``.
    The lower and upper bounds, *n* and *m*, are irrelevant. Only the dependency matters.

-   To this, we can add "Format 0" which has a single occurrence, :math:`O(d)=1`.
    This is the default if no occurrence information is provided.
    
This adds another interesting attribute, the total size of an elementary item, 
:math:`T(d) = \operatorname{s}(d) \times O(d)`. 
The total size is the elementary size times the number of occurrences. 

This expands our definition of size of a group item to the sum of the total sizes,
not the sum of the simple size. 

..  math::

    T(g) = \operatorname{s}(g) = \sum_{d \in g}{T(d)} \times O(g)
    
This shows how the occurs information for a group applies to all the elements inside a group.

We can think of :math:`O(d) = O(G(d))` as the rule for applying the occurs information 
to all members of a group. If we do this, then there must be a root group, :math:`G_0`
for which :math:`O(G_0) = 1`. This allows us to walk "up" the tree locating all of the
parental occurrence information for a given elementary item.

The occurs clause also changes how we fetch the bytes, :math:`B[d]`, because we need index information
for each ``OCCURS`` clause in the parents of *d*. In COBOL we might be getting ``D( I )``. 

..    math::

    B[d:i] = \{ B_x | \operatorname{o}(d) + i \operatorname{s}(d) \leq x < \operatorname{o}(d) + (i+1)\operatorname{s}(d) \}

The indices may be much more complex, however. A common situation is a group-level item
with an occurrence that nests an elementary item with an occurrence. This is a two-dimensional
structure. In COBOL it might be ``D( J ) OF G( I )``. More commonly, it's written ``D( I, J )`` where
the indices are applied from top-most group DDE down the tree to the lowest-level items.

..    math::

    B[d:i,j] = \{ B_x | \operatorname{o}(G(d:i)) + j \operatorname{s}(d) \leq x < \operatorname{o}(G(d:i))+(j+1)\operatorname{s}(d) \}

The offset, :math:`\operatorname{o}(d)` is computed recursively using
a combination of predecessors, :math:`P(d)`, groups, :math:`G(d)`, and redefinitions,
:math:`R(d)`. There are several cases.

-   :math:`\operatorname{o}(d) = \operatorname{o}(R(d))` if this item is defined by a REDEFINES clause.

-   :math:`\operatorname{o}(d) = \operatorname{o}(P(d)) + \operatorname{s}(P(d)) \times O(P(d))`
    if this item has a predecessor.
    
    The predecessor could involve an occurs clause that depends on actual data:
    :math:`O(P(d)) = B[ D(P(d)) ]`. 

-   :math:`\operatorname{o}(d) = \operatorname{o}(G(d)) + \operatorname{s}(G(d)) \times O(G(d))`
    if this item is the first member of a containing group, and has no predecessor.

    This is more complex than it appears because a group could contain an occurs clause,
    which influences all items within the group.
    :math:`\operatorname{o}(G(d:i)) = \operatorname{o}(G(d)) + i \operatorname{s}(G(d))`.
    
    This, too, can depend on actual data: :math:`O(G(d)) = B[ D(G(d)) ]`.

-   :math:`\operatorname{o}(d) = 0`, if there's neither predecessor nor group.

We've called out the two cases which involve the occurs information
that may include "format 2" occurs and depend on data within an actual record.
These mean that the calculation cannot be done eagerly in this situation.

Variable-Length COBOL Records
------------------------------

See 

https://publib.boulder.ibm.com/infocenter/zos/basics/index.jsp?topic=/com.ibm.zos.zconcepts/zconcepts_159.htm

V (Variable)

    This format has one logical record as one physical block. A
    variable-length logical record consists of a record descriptor word
    (RDW) followed by the data. The record descriptor word is a 4-byte field
    describing the record. The first 2 bytes contain the length of the
    logical record (including the 4-byte RDW). The length can be from 4 to
    32,760 bytes. All bits of the third and fourth bytes must be 0, because
    other values are used for spanned records. This format is seldom used.

VB (Variable Blocked)

    This format places several variable-length logical records (each with an
    RDW) in one physical block. The software must place an additional Block
    Descriptor Word (BDW) at the beginning of the block, containing the
    total length of the block.
    
Here's a handy diagram showing five kinds of RECFM.

..  image:: zOSB037.png

See http://en.wikipedia.org/wiki/Data_set_(IBM_mainframe)

See http://www.simotime.com/vrecex01.htm

This leads us to three relevant encoding possibilities. These are the COBOL "RECFM" 
options in the JCL for the file. Each record is preceded by a 
Record Descriptor Word (RDW).

``V``   Variable. 

    The four bytes preceding the logical record is the Record Descriptor Word. The content is as follows.

    =====   ============
    Bytes   Description
    =====   ============
    1-2     This is the length of the logical record plus the length of the four-byte Descriptor Word. 
    3-4     Usually low values
    =====   ============

``VB``  Variable Blocked.

    The four bytes preceding the logical record (Descriptor Word) are as follows.

    =====   ============
    Bytes   Description
    =====   ============
    1-2     This is the length of the logical record plus the length of the four-byte Descriptor Word. 
    3-4     The length of the block including four-byte Descriptor Word.
    =====   ============
    
    A block can have multiple records in it. The block length must be >= record length.

``VBS``     Variable Blocked Spanned.

    The four bytes preceding the logical record is the Segment Descriptor Word. The content is as follows.

    =====   ============
    Bytes   Description
    =====   ============
    1-2     This is the length of the logical record plus the length of the four-byte Descriptor Word.
    3       Segment Control Codes: see below.
    4       Low value, reserved for future use
    =====   ============

    Segment Control Code
    
    =====   ================================================
    Value   Relative position of segment
    =====   ================================================
    00      A complete logical record 
    01      The first segment of a multiple segments record
    02      The last segment of a multiple segments record
    03      A middle segment of a multiple segments record
    =====   ================================================

This RECFM detail must be provided as part of opening the workbook/sheet so that rows can be 
properly located within the content.

We've added a series of RECFM classes as a **Strategy** to 
read files with variable length records.

Low-Level Split Processing
---------------------------------

We may have a need to split an EBCDIC file, similar to the Posix ``split`` command.
This is done using :py:class:`cobol.RECFM` parsers to read records and write to 
new file(s).

A splitter looks like this:

..  parsed-literal::

    import itertools
    import stringray.cobol
    import collections
    import pprint
    
    batch_size= 1000
    counts= collections.defaultdict(int)
    with open( "some_file.schema", "rb" ) as source:
        reader= stringray.cobol.RECFM_VB( source ).bdw_iter()
        batches= itertools.groupby( enumerate(reader), lambda x: x[0]//batch_size ):
        for group, group_iter in batches:
            with open( "some_file_{0}.schema".format(group), "wb" ) as target:
            for id, row in group_iter:
                target.write( row )
                counts['rows'] += 1
                counts[str(group)] += 1
    pprint.pprint( dict(counts) )
                
There are several possible variations on the construction of the ``reader`` object.

-   cobol.RECFM_F( source ).record_iter() -- result is RECFM_F.

-   cobol.RECFM_F( source ).rdw_iter() -- result is RECFM_V; RDW's have been added.

-   cobol.RECFM_V( source ).rdw_iter() -- result is RECFM_V; RDW's have been preserved.

-   cobol.RECFM_VB( source ).rdw_iter() -- result is RECFM_V; RDW's have been preserved;
    BDW's have been discarded.

-   cobol.RECFM_VB( source ).bdw_iter() -- result is RECFM_VB; BDW's and RDW's have been 
    preserved. The batch size is the number of blocks, not the number of records.

    
The Bad Data Problem
----------------------

Even with a :py:class:`sheet.LazyRow`, we have to be tolerant of COBOL data which doesn't match
the schema. The :py:meth:`cobol.defs.Usage.create_func` function may encounter an exception.
If so, then a :py:class:`cobol.defs.ErrorCell` is created instead of the class
defined by the Usage and Picture classes.

The :py:class:`cobol.defs.ErrorCell` includes the raw bytes, but a value of ``None``.

..  _`cobol_impl`: 

COBOL Implementation
======================

These are the modules that extend the core functionality of Stingray
to handle COBOL files and EBCDIC data.

..  toctree::
    :maxdepth: 1

    cobol_init
    cobol_loader
    cobol_defs
