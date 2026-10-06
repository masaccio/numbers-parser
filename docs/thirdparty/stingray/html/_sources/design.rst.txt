Design Considerations
=====================

There are several important data representation issues. We'll look at 
`Workbook Schema Issues`_ and `Some COBOL File Issues`_ that
will inform the rest of the design.

Workbook Schema Issues
-------------------------------------

There are several sensible ways to represent workbook schema information. 
Two of these are seen in the wild, and must be handled gracefully. Others
could be seen, and should be part of the use cases for this package.

1.  One (or more) rows in each workbook sheet that provides the attribute
    name.  The type is often left implicit.  This is often bound to the
    data, making it quite useful in practice.
        
#.  A distinct workbook sheet that lists name (and optionally type) 
    for each attribute.  Offset and size can
    also be provided, making this useful for fixed-format COBOL files.
    This may be bound to data in a separate sheet of the same workbook,
    making it more likely to be useful. If it describes a COBOL file,
    however, the binding is more tenuous.

#.  A Python module that's built from source information.  This
    allows us to trivially ``import schema.foo`` and have lots of cool
    classes and functions in the ``schema.foo`` module. Ideally the 
    module is properly bound to the file, but there's little guarantee 
    of this.

#.  Some standardized metadata format.  XSD (or even XMI) pops into mind.
    These are detailed and potentially useful.
    This is vanishingly rare in practice. The external XSD may not 
    be bound in any way to the data.

Options 1 and 2 (workbook-based schema) cover over 99%
of the cases in practice.  While the data is casual and error-prone, it's
often readily available.

Option 3 (a Python module) -- while cool -- breaks the DRY
principle.  Refreshing a Python schema when a source :file:`.XLS` document
changes means that we have a multi-step binding: document to schema followed
by schema to application.

Option 4 is rarely used in practice.  In the rare cases when an organization
will consent to providing XSD files, they're often prepared separately
from the data and do not actually reflect the application software
that is used to build the data file.

How do we extend this to handle COBOL DDE's? 

How do we extend this to handle JSON, YAML, or XML?

Clearly, we need to load schema information from a source every time
it's needed.  There are two paths:

-   Either we'll parse the embedded schema buried in each
    sheet of a workbook,

-   or we'll load an external schema definition from
    a file.
    
We'll need to design schema loaders as well as data extractors which
depend on the schema loaders.

Some COBOL File Issues
-------------------------

When dealing with "Flat Files" from legacy COBOL problems, there are several
additional problems that need to be solved.

1.  The files have a fixed field layout, without delimiters.
    This means that the offset of each field must be used to
    decompose the record into its individual elements.

#.  Numeric fields can have an implied decimal point, making
    it difficult to determine the value of a string of digits.
    The COBOL DDE is essential for parsing the file contents.

#.  COBOL can make use of numeric data represented in a variety
    of "Computational" forms.  The "Computational-3" ("COMP-3")
    form is very popular and rather complex because decimal digits are
    packed two per byte and the final half-byte encodes
    sign information.

#.  The string data may be encoded in EBCDIC bytes.

#.  COBOL encourages the use of data aliases (or "unions") via the ``REDEFINES`` clause.
    Without the entire unviverse of COBOL programs that work with a given file,
    the general handling of ``REDEFINES``
    data elements can become an insoluable problem.
    Only lazy field access can work;
    eager creation of individual cell values is doomed because a
    ``REDEFINES`` alternative may be defined over invalid data.
    
#.  COBOL has an ``OCCURS DEPENDING ON`` (ODO) feature where one attribute
    determines the size of another attribute. This means the data 
    of every attribute after the ODO attribute has a location which varies.
    The positions within the flat file cannot be computed statically.

Generally, COBOL files are defined by a "Data Definition Entry" (DDE)
that provides the record layout.
It's essential to
parse this source DDE, which has the original COBOL
definition for the file.  A schema can be built from the parsed DDE.
There's no good reason to rely on any intermediate description separate from
the DDE's themselves.
