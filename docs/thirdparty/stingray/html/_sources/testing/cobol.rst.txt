#######################
COBOL Schema Testing
#######################

Processing COBOL (and possibly EBCDIC) files requires a more sophisticated
schema.  And a more sophisticated schema loader.

See :ref:`cobol`.

Overheads
=================

::

  """stingray.cobol Unit Tests."""
  import unittest
  import decimal
  import weakref
  import logging, sys
  import io
  import types

  import stingray.cobol
  from stingray.cobol.loader import Picture


Handy Mocks
==============

::

  class MockDefaultCell:
      def __init__( self, value, workbook, attr ):
          self.value= value
          self.workbook= workbook
          sekf.attr= attr
      def __eq__( self, other ):
          return all( (
              self.__class__ == other.__class__,
              self.value == other.value,
              self.attr == other.attr
              ) )
      def __ne__( self, other ):
          return not self.__eq__( other )
                                
  class MockDDE:
      def __init__( self, **kw ):
          self.totalSize= 0 # if otherwise unspecified
          self.kw= kw
          self.__dict__.update( kw )
          self.children=[]
      def __repr__( self ):
          nv= ", ".join( "{0}={1!r}".format( k, self.kw[k] ) for k in self.kw )
          return "MockDDE( {1} )".format( self.children, nv )
      def addChild( self, child ):
          self.children.append( child )
          self.totalSize += child.totalSize

  class MockNonRedefine:
      def __init__( self, **kw ):
          self.__dict__.update( kw )
        
  class MockOccurs:
      def __init__( self, number ):
          self._number= number
      def number( self, aRow ):
          return self._number
      def __repr__( self ):
          return "{_class_name}({_number})".format(_class_name=self.__class__.__name__, **self.__dict__)
        
  class MockSchema( list ):
      def __init__( self, *args, **kw ):
          super().__init__( *args )
          self.info= kw
      def lrecl( self ):
          return max( a.offset+a.size for a in self )
                
Repeating Attribute
=====================

We subclass :py:class:`schema.Attribute` to make it depend on  :py:class:`MockDefaultCell` 
instead of some working class in :py:mod:`cell`.

::
        
  class AttributeFixture( stingray.cobol.RepeatingAttribute ):
      default_cell= MockDefaultCell

Now we can test simple DDE's without an occurs clause.

::

  class TestNonRepeatingAttribute( unittest.TestCase ):
      def setUp( self ):
          self.simple_dde= MockDDE( level="05", name="SOME-COLUMN", occurs=MockOccurs(1),
              size=5, totalSize=5, offset=0 )
          self.simple_dde.dimensionality= ()
          self.simple= AttributeFixture( 
              name="SOME-COLUMN", dde=weakref.ref(self.simple_dde),
              size=5, create=MockDefaultCell,  )
      def test_should_have_names( self ):
          self.assertEqual( "SOME-COLUMN", self.simple.name )
      def test_should_not_index_simple( self ):
          try:
              simple= self.simple.index(2)
              self.fail()
          except IndexError:
              pass

And we can test simple DDE's with an occurs clause.

:: 

  class TestRepeatingAttribute( unittest.TestCase ):
      def setUp( self ):
          self.dde= MockDDE( level="05", name="REPEAT", occurs=MockOccurs(4), size=3, totalSize=12, offset=5)
          self.dde.dimensionality=(self.dde,) 
          self.complex= AttributeFixture( 
              name="REPEAT", dde=weakref.ref(self.dde),
              create=MockDefaultCell, occurs=4, size=3, totalSize=12, )
      def test_should_have_names( self ):
          self.assertEqual( "REPEAT", self.complex.name )
      def test_should_index_complex( self ):
          """1-based indexing in COBOL is ignored here."""
          complex= self.complex.index(0)
          self.assertEqual( 5, complex.offset )
          self.assertEqual( 3, complex.size )
          complex= self.complex.index(1)
          self.assertEqual( 8, complex.offset )
          self.assertEqual( 3, complex.size )
          complex= self.complex.index(2)
          self.assertEqual( 11, complex.offset )
          self.assertEqual( 3, complex.size )
          complex= self.complex.index(3) # out of range!
          self.assertEqual( 14, complex.offset )
          self.assertEqual( 3, complex.size )

And we can test nested DDE's with  multiple occurs clauses.

:: 

  class TestNestedRepeatingAttribute( unittest.TestCase ):
      def setUp( self ):
          self.child= MockDDE( level="10", name="ITEM", 
              occurs=MockOccurs(4), size=2, totalSize=8, offset=5,
              )
          self.parent= MockDDE( level="05", name="REPEAT", 
              occurs=MockOccurs(5), size=8, totalSize=30,  offset=5, children=[self.child],
              )
          self.parent.dimensionality=(self.parent,) 
          self.child.dimensionality=(self.parent,self.child)
          self.complex_05= AttributeFixture( 
              name="REPEAT", dde=weakref.ref(self.parent), 
              create=MockDefaultCell, occurs=5, size=8, totalSize=40, )
          self.complex_10= AttributeFixture( 
              name="ITEM", dde=weakref.ref(self.child), 
              create=MockDefaultCell, occurs=4, size=2, totalSize=8,  )
      def test_should_have_names( self ):
          self.assertEqual( "REPEAT", self.complex_05.name )
          self.assertEqual( "ITEM", self.complex_10.name )
      def test_should_index_complex_item( self ):
          """Partial index works from top to bottom.
          COBOL is 1-based, but that's ignored here."""
          complex= self.complex_10.index(0)
          self.assertEqual( 5, complex.offset )
          self.assertEqual( 2, complex.size )
          complex= self.complex_10.index(1)
          self.assertEqual( 13, complex.offset )
          self.assertEqual( 2, complex.size )
          complex= self.complex_10.index(2)
          self.assertEqual( 21, complex.offset )
          self.assertEqual( 2, complex.size )
          complex= self.complex_10.index(3) 
          self.assertEqual( 29, complex.offset )
          self.assertEqual( 2, complex.size )
      def test_should_index_complex_group( self ):
          complex= self.complex_10.index(1,0)
          self.assertEqual( 13, complex.offset )
          self.assertEqual( 2, complex.size )
          complex= self.complex_10.index(1,1)
          self.assertEqual( 15, complex.offset )
          self.assertEqual( 2, complex.size )
          complex= self.complex_10.index(2,0)
          self.assertEqual( 21, complex.offset )
          self.assertEqual( 2, complex.size )
          complex= self.complex_10.index(2,1)
          self.assertEqual( 23, complex.offset )
          self.assertEqual( 2, complex.size )
        
DDE Data Access
=====================

The data access mediated by a DDE-as-schema is a bit more complex than the
trivial (flat) data access mediated by the  schema representation
from the :py:mod:`schema` package.

Note that we don't test :py:class:`sheet.LazyRow` on :py:class:`workbook.Fixed_Workbook`.  It doesn't work properly because
the :py:class:`workbook.Fixed_Workbook` doesn't know about the new 
:py:class:`cobol.RepeatingAttribute` that models repeating groups.

The :py:class:`cobol.Character_File`, however, can depend on :py:class:`cobol.RepeatingAttribute`.

::

  class Test_Dimensional_Access( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE( level='01', name='SURVEY-RESPONSES', 
              allocation=MockNonRedefine(), picture=None, offset=0, size=60, occurs=MockOccurs(1),
              variably_located= False,
              )
          self.top.top= self.top
          self.group_05 = MockDDE( level='05', name='QUESTION-NUMBER', 
              allocation=MockNonRedefine(), picture=None, occurs=MockOccurs(10), offset=0, totalSize=60, size=6 )
          self.top.addChild( self.group_05 )
          self.group_10 = MockDDE( level='10', name='RESPONSE-CATEGORY', 
              allocation=MockNonRedefine(), picture=None, occurs=MockOccurs(3), offset=0, totalSize=6, size=2 )
          self.group_05.addChild( self.group_10 )
          self.group_15 = MockDDE( level='15', name='ANSWER', 
              allocation=MockNonRedefine(), picture="99", offset=0, totalSize=2, size=2 )
          self.group_15.dimensionality= (self.group_05, self.group_10,)
          self.group_10.addChild( self.group_15 )

          self.schema= MockSchema(
              (stingray.cobol.RepeatingAttribute( name="ANSWER", dde=weakref.ref(self.group_15),
                  size=2, totalSize=60, ),
              ),
              dde= [self.top]
          )
        
          self.data = stingray.cobol.Character_File( name="", 
              file_object= ["111213212223313233414243515253616263717273818283919293010203",],
              schema=self.schema )
            
      def test_should_get_cell( self ):
          """Get individual cell values."""
          # 1-based indexing in COBOL, Python, however, is zero-based.
          row= next( self.data.sheet( "" ).rows() )
          self.assertEqual( "11", row.cell( self.schema[0].index(0,0) ).to_str() )
          self.assertEqual( "21", row.cell( self.schema[0].index(1,0) ).to_str() )
          self.assertEqual( "22", row.cell( self.schema[0].index(1,1) ).to_str() )
          self.assertEqual( "23", row.cell( self.schema[0].index(1,2) ).to_str() )
          self.assertEqual( "93", row.cell( self.schema[0].index(8,2) ).to_str() )
      def test_should_get_array( self ):
          """Get tuple of values from incomplete index specification."""
          # 1-based indexing in COBOL, Python, however, is zero-based.
          row= next( self.data.sheet( "" ).rows() )
          #print( "Original and Tweaked =", self.schema[0], self.schema[0].index(1) )
          slice= row.cell( self.schema[0].index(1) )
          #print( "Offset, Slice =", self.schema[0].index(1).offset, slice )
          self.assertEqual( ('21', '22', '23'), tuple( c.to_str() for c in slice ) )
        

Workbook File Access    
========================

Character File and EBCDIC File access covers most of the bases.

Here are mocks to define a schema and the resulting Cell instances.

::

  class MockAttribute:
      def __init__( self, **kw ):
          self.__dict__.update( kw )
      def __repr__( self ):
          return repr( self.__dict__ )
        
  class MockTextCell:
      def __init__( self, buffer, workbook, attr ):
          self.raw, self.workbook, self.attr= buffer, workbook, attr
          self.value= workbook.text( self.raw, self.attr )
        
  class MockDisplayCell( MockTextCell ):
      def __init__( self, buffer, workbook, attr ):
          self.raw, self.workbook, self.attr= buffer, workbook, attr
          self.value= workbook.number_display( self.raw, self.attr )

  class MockComp3Cell( MockTextCell ):
      def __init__( self, buffer, workbook, attr ):
          self.raw, self.workbook, self.attr= buffer, workbook, attr
          self.value= workbook.number_comp3( self.raw, self.attr )

  class MockCompCell( MockTextCell ):
      def __init__( self, buffer, workbook, attr ):
          self.raw, self.workbook, self.attr= buffer, workbook, attr
          self.value= workbook.number_comp( self.raw, self.attr )

In principle, we need to mock the :py:class:`sheet.ExternalSchemaSheet` that
gets built.

Given a schema and some data, pick apart the row of a Character File. Generally,
we'll deal with ordinary text using no particular encoding. Python can generally
handle ASCII bytes or Unicode pretty simply via the "encoding" argument to
the :py:func:`open`.

::

  class TestCharacterFile_Text( unittest.TestCase ):
      def setUp( self ):
          self.parent= MockDDE( level="01", name="PARENT", variably_located= False, )
          self.attr_word= MockAttribute( 
              name="WORD", offset=0, size=3, create=MockTextCell, dimensionality=None,
              size_scale_precision=Picture("XXX",True,3,0,0,False,None))
          self.attr_num1= MockAttribute( 
              name="NUMBER-1", offset=3, size=6, create=MockDisplayCell, picture="999.99", dimensionality=None,
              size_scale_precision=Picture("999.99",False,6,0,2,True,"."))
          self.attr_num2= MockAttribute( 
              name="NUMBER-2", offset=9, size=5, create=MockDisplayCell, picture="S999V99", dimensionality=None,
              size_scale_precision=Picture("999V99",False,6,0,2,True,"V"))
          self.schema= MockSchema( 
              ( self.attr_word, self.attr_num1, self.attr_num2
              ),
              dde= [ self.parent ],
          )
          self.data= ["ASC123.4567890XYZ",]
          self.wb= stingray.cobol.Character_File( "name", self.data, self.schema )
      def test_should_get_cells( self ):
          row= next( self.wb.sheet( "IGNORED" ).rows() )
          self.assertEqual( "ASC", row.cell( self.attr_word ).value )
          self.assertEqual( decimal.Decimal('123.45'), row.cell( self.attr_num1 ).value )
          self.assertEqual( decimal.Decimal('678.90'), row.cell( self.attr_num2 ).value )

Do we need ``TestCharacterFile_Bytes`` as a separate test case? Probably not. The
:py:class:`cobol.Character_File` is a :py:class:`workbook.Fixed_Workbook` which is
text, not bytes.

Given a schema and some data, pick apart the row of a fake EBCDIC File.
This file has the default RECFM of F -- no blocking and no header words.

::

  class TestEBCDICFile_Fixed( unittest.TestCase ):
      def setUp( self ):
          self.parent= MockDDE( level="01", name="PARENT", variably_located= False, )
          self.attr_word= MockAttribute( 
              name="WORD", offset=0, size=3, create=MockTextCell, dimensionality=None, 
              size_scale_precision=Picture("XXX",True,3,0,0,False,None), )
          self.attr_num1= MockAttribute( 
              name="NUMBER-1", offset=3, size=6, create=MockDisplayCell, picture="999.99", dimensionality=None, 
              size_scale_precision=Picture("999.99",False,6,0,2,True,"."), )
          self.attr_num2= MockAttribute( 
              name="NUMBER-2", offset=9, size=5, create=MockDisplayCell, picture="S999V99", dimensionality=None, 
              size_scale_precision=Picture("999V99",False,5,0,2,True,"V"), )
          self.attr_comp= MockAttribute( 
              name="NUMBER-3", offset=14, size=4, create=MockCompCell, picture="S99999999", dimensionality=None, 
              size_scale_precision=Picture("999V99",False,5,0,2,True,"V"), )
          self.attr_comp3= MockAttribute( 
              name="NUMBER-4", offset=18, size=3, create=MockComp3Cell, picture="S999V99", dimensionality=None, 
              size_scale_precision=Picture("999V99",False,3,0,2,True,"V"), )
          self.schema= MockSchema( 
              ( self.attr_word, self.attr_num1, self.attr_num2, self.attr_comp, self.attr_comp3
              ),
              dde= [ self.parent ],
          )
      def test_should_get_cells( self ):
          self.data= io.BytesIO( 
              b"\xe9\xd6\xe2" # WORD="ZOS"
              b"\xf1\xf2\xf3K\xf4\xf5" # NUMBER-1="123.45"
              b"\xf6\xf7\xf8\xf9\xf0" # NUMBER-2="678.90"
              b"\x00\x00\x12\x34" # NUMBER-3=4660
              b"\x98\x76\x5d" # NUMBER-4=-987.65
          )
          self.wb= stingray.cobol.EBCDIC_File( "name", self.data, self.schema )
          row= next( self.wb.sheet( "IGNORED" ).rows() )
          self.assertEqual( 21, self.schema.lrecl() ) # last field has offset 18, size 3
          self.assertEqual( "ZOS", row.cell( self.attr_word ).value )
          self.assertEqual( decimal.Decimal('123.45'), row.cell( self.attr_num1 ).value )
          self.assertEqual( decimal.Decimal('678.90'), row.cell( self.attr_num2 ).value )
          self.assertEqual( decimal.Decimal('4660'), row.cell( self.attr_comp ).value )
          self.assertEqual( decimal.Decimal('-987.65'), row.cell( self.attr_comp3 ).value )
  
Given a schema and some data, pick apart the row of a fake EBCDIC File.
This file has a RECFM of V, it has Record Descriptor Word (RDW)

::

  class TestEBCDICFile_Variable( TestEBCDICFile_Fixed ):
      def test_should_get_cells( self ):
          """Data has 4 byte RDW in front of the row."""
          self.data= io.BytesIO( 
              b"\x00\x19\x00\x00" # RDW
              b"\xe9\xd6\xe2" # WORD="ZOS"
              b"\xf1\xf2\xf3K\xf4\xf5" # NUMBER-1="123.45"
              b"\xf6\xf7\xf8\xf9\xf0" # NUMBER-2="678.90"
              b"\x00\x00\x12\x34" # NUMBER-3=4660
              b"\x98\x76\x5d" # NUMBER-4=-987.65                
          )
          self.wb= stingray.cobol.EBCDIC_File( "name", self.data, self.schema, RECFM="V" )
          row= next( self.wb.sheet( "IGNORED" ).rows() )
          self.assertEqual( 21, self.schema.lrecl() ) # last field has offset 18, size 3
          self.assertEqual( "ZOS", row.cell( self.attr_word ).value )
          self.assertEqual( decimal.Decimal('123.45'), row.cell( self.attr_num1 ).value )
          self.assertEqual( decimal.Decimal('678.90'), row.cell( self.attr_num2 ).value )
          self.assertEqual( decimal.Decimal('4660'), row.cell( self.attr_comp ).value )
          self.assertEqual( decimal.Decimal('-987.65'), row.cell( self.attr_comp3 ).value )

EBCDIC File with VB format. It has Block Descriptor Word (BDW) and 
Record Descriptor Word (RDW).

::

  class TestEBCDICFile_VariableBlocked( TestEBCDICFile_Fixed ):
      def test_should_get_cells( self ):
          """Data has 4 byte BDW and 4 byte RDW in front of the row."""
          # Build 2 blocks each with two records.
          self.data= io.BytesIO( 
          b"\x00\x36\x00\x00" #BDW: 54= 4+sum of RDW's
              b"\x00\x19\x00\x00" # RDW: 25=4+len(data)
                  b"\xe9\xd6\xe2" # WORD="ZOS"
                  b"\xf1\xf2\xf3K\xf4\xf5" # NUMBER-1="123.45"
                  b"\xf6\xf7\xf8\xf9\xf0" # NUMBER-2="678.90"
                  b"\x00\x00\x12\x34" # NUMBER-3=4660
                  b"\x98\x76\x5d" # NUMBER-4=-987.65                
              b"\x00\x19\x00\x00" # RDW
                  b"\xe9\xd6\xe2" # WORD="ZOS"
                  b"\xf1\xf2\xf3K\xf4\xf6" # NUMBER-1="123.46"
                  b"\xf6\xf7\xf8\xf9\xf0" # NUMBER-2="678.90"
                  b"\x00\x00\x12\x34" # NUMBER-3=4660
                  b"\x98\x76\x5d" # NUMBER-4=-987.65                
          b"\x00\x36\x00\x00" #BDW
              b"\x00\x19\x00\x00" # RDW
                  b"\xe9\xd6\xe2" # WORD="ZOS"
                  b"\xf1\xf2\xf3K\xf4\xf7" # NUMBER-1="123.47"
                  b"\xf6\xf7\xf8\xf9\xf0" # NUMBER-2="678.90"
                  b"\x00\x00\x12\x34" # NUMBER-3=4660
                  b"\x98\x76\x5d" # NUMBER-4=-987.65
              b"\x00\x19\x00\x00" # RDW
                  b"\xe9\xd6\xe2" # WORD="ZOS"
                  b"\xf1\xf2\xf3K\xf4\xf8" # NUMBER-1="123.48"
                  b"\xf6\xf7\xf8\xf9\xf0" # NUMBER-2="678.90"
                  b"\x00\x00\x12\x34" # NUMBER-3=4660
                  b"\x98\x76\x5d" # NUMBER-4=-987.65                
          )
          self.wb= stingray.cobol.EBCDIC_File( "name", self.data, self.schema, RECFM="VB" )
          row_iter= self.wb.sheet( "IGNORED" ).rows()
          row= next( row_iter )
          self.assertEqual( 21, self.schema.lrecl() ) # last field has offset 18, size 3
          self.assertEqual( "ZOS", row.cell( self.attr_word ).value )
          self.assertEqual( decimal.Decimal('123.45'), row.cell( self.attr_num1 ).value )
          self.assertEqual( decimal.Decimal('678.90'), row.cell( self.attr_num2 ).value )
          self.assertEqual( decimal.Decimal('4660'), row.cell( self.attr_comp ).value )
          self.assertEqual( decimal.Decimal('-987.65'), row.cell( self.attr_comp3 ).value )
          row2= next( row_iter )
          self.assertEqual( decimal.Decimal('123.46'), row2.cell( self.attr_num1 ).value )
          row3= next( row_iter )
          self.assertEqual( decimal.Decimal('123.47'), row3.cell( self.attr_num1 ).value )
          row4= next( row_iter )
          self.assertEqual( decimal.Decimal('123.48'), row4.cell( self.attr_num1 ).value )
          with self.assertRaises(StopIteration):
              next(row_iter)

Given a schema and some data, pick apart the row of a fake EBCDIC File.
This file has a RECFM of V, but not Record Descriptor Word (RDW).
The RECFM=N should be able to parse it, however.

::

  class TestEBCDICFile_VariableWithoutRDW( TestEBCDICFile_Fixed ):
      def test_should_get_cells( self ):
          """Data lacks a 4 byte RDW in front of the row."""
          buffer= (
              b"\xe9\xd6\xe2" # WORD="ZOS"
              b"\xf1\xf2\xf3K\xf4\xf5" # NUMBER-1="123.45"
              b"\xf6\xf7\xf8\xf9\xf0" # NUMBER-2="678.90"
              b"\x00\x00\x12\x34" # NUMBER-3=4660
              b"\x98\x76\x5d" # NUMBER-4=-987.65
            
              b"\xe9\xd6\xe2" # WORD="ZOS"
              b"\xf1\xf2\xf3K\xf4\xf5" # NUMBER-1="123.45"
              b"\xf6\xf7\xf8\xf9\xf0" # NUMBER-2="678.90"
              b"\x00\x00\x12\x34" # NUMBER-3=4660
              b"\x98\x76\x5d" # NUMBER-4=-987.65                
          )
          self.data= io.BytesIO( buffer )
          self.wb= stingray.cobol.EBCDIC_File( "name", self.data, self.schema, RECFM="N" )
          row_iter= self.wb.sheet( "IGNORED" ).rows()
          row= next( row_iter )
          self.assertEqual( 21, self.schema.lrecl() ) # last field has offset 18, size 3
          self.assertEqual( "ZOS", row.cell( self.attr_word ).value )
          self.assertEqual( decimal.Decimal('123.45'), row.cell( self.attr_num1 ).value )
          self.assertEqual( decimal.Decimal('678.90'), row.cell( self.attr_num2 ).value )
          self.assertEqual( decimal.Decimal('4660'), row.cell( self.attr_comp ).value )
          self.assertEqual( decimal.Decimal('-987.65'), row.cell( self.attr_comp3 ).value )
          row2= next( row_iter )
          with self.assertRaises(StopIteration):
              row3= next( row_iter )


..  todo:: EBCDIC File V format with Occurs Depending On to show the combination.

EBCDIC file with bad numeric data. This doesn't use the above Mocks. It uses real
Cell definitions and real Usage definitions because the real Usage definitions tolerate bad data.

::

  class TestEBCDICFile_Invalid( unittest.TestCase ):
      def setUp( self ):
          self.parent= MockDDE( level="01", name="PARENT", variably_located= False, )
          self.display= stingray.cobol.defs.UsageDisplay("DISPLAY")
          self.display.numeric= True
          self.comp= stingray.cobol.defs.UsageComp("COMP")
          self.comp3= stingray.cobol.defs.UsageComp3("COMP-3")
          self.attr_word= MockAttribute( 
              name="WORD", offset=0, size=3, create=stingray.cobol.defs.TextCell, dimensionality=None,
              size_scale_precision=Picture("XXX",True,3,0,0,False,None), )
          self.attr_num1= MockAttribute( 
              name="NUMBER-1", offset=3, size=6, create=self.display.create_func, picture="999.99", dimensionality=None, 
              size_scale_precision=Picture("999.99",False,6,0,2,True,"."), )
          self.attr_num2= MockAttribute( 
              name="NUMBER-2", offset=9, size=5, create=self.display.create_func, picture="S999V99", dimensionality=None, 
              size_scale_precision=Picture("999V99",False,6,0,2,True,"V"), )
          self.attr_comp= MockAttribute( 
              name="NUMBER-3", offset=14, size=4, create=self.comp.create_func, picture="S99999999", dimensionality=None, 
              size_scale_precision=Picture("999V99",False,6,0,2,True,"V"), )
          self.attr_comp3= MockAttribute( 
              name="NUMBER-4", offset=18, size=5, create=self.comp3.create_func, picture="S999V99", dimensionality=None, 
              size_scale_precision=Picture("999V99",False,6,0,2,True,"V"), )
          self.schema= MockSchema( 
              ( self.attr_word, self.attr_num1, self.attr_num2, self.attr_comp, self.attr_comp3
              ),
              dde= [ self.parent ],
          )
      def test_should_get_error_cells( self ):
          self.data= io.BytesIO( b"\xe9\xd6\xe2\xf1\xd6\xf3K\xf4\xd6\xf6\xf7\xf8\xd6\xf0\x00\x00\x12\xd6\x00\x00\x00\x00\x00" )
          self.wb= stingray.cobol.EBCDIC_File( "name", self.data, self.schema )
          row= next( self.wb.sheet( "IGNORED" ).rows() )
          self.assertEqual( 23, self.schema.lrecl() )
          self.assertEqual( "ZOS", row.cell( self.attr_word ).value )
          self.assertIsNone( row.cell( self.attr_num1 ).value )
          self.assertEqual( b'\xf1\xd6\xf3K\xf4\xd6', row.cell( self.attr_num1 ).raw )
          self.assertIsNone( row.cell( self.attr_num2 ).value )
          self.assertEqual( b'\xf6\xf7\xf8\xd6\xf0', row.cell( self.attr_num2 ).raw )
          # Should Work: can't easily corrupt USAGE COMPUTATIONAL data.
          self.assertEqual( decimal.Decimal('4822'), row.cell( self.attr_comp ).value )
          self.assertEqual( b'\x00\x00\x12\xd6', row.cell( self.attr_comp ).raw )
          # Should this work? Not clear.
          #self.assertIsNone( row.cell( self.attr_comp3 ).value )
          self.assertEqual( decimal.Decimal('0'), row.cell( self.attr_comp3 ).value )
          self.assertEqual( b'\x00\x00\x00\x00\x00', row.cell( self.attr_comp3 ).raw )

Test Conversions
==================

The various PIC and USAGE format subtleties lead to a number of conversion test cases.
The attributes are a shortcut for the DDE. This includes some explicit
attributes, plus the picture_parser return value. It's
a Picture instance, a tuple of (final, alpha, len(final), scale, precision, signed, decimal)


::

  class TestNumberConversion( unittest.TestCase ):
      def setUp( self ):
          self.wb= stingray.cobol.EBCDIC_File
      def test_should_convert_display(self):
          attr1= types.SimpleNamespace(
              size_scale_precision = Picture('999v99', False, 5, 0, 2, True, "V") )
          n1= self.wb.number_display( b'\xf1\xf2\xf3\xf4\xf5', attr1 )
          self.assertEqual( decimal.Decimal('123.45'), n1 )
          attr2= types.SimpleNamespace(
              size_scale_precision = Picture('99999V', False, 5, 0, 0, True, "V") )
          n2= self.wb.number_display( b'\xf1\xf2\xf3\xf4\xf5', attr2 )
          self.assertEqual( decimal.Decimal('12345'), n2 )
          attr3= types.SimpleNamespace(
              size_scale_precision = Picture('999.99', False, 5, 0, 2, True, ".") )
          n3= self.wb.number_display( b'\xf1\xf2\xf3K\xf4\xf5', attr3 )
          self.assertEqual( decimal.Decimal('123.45'), n3 )
        
      def test_should_convert_display_2(self):
          attr1= types.SimpleNamespace(
              size_scale_precision = Picture('9', False, 1, 0, 0, True, "V") )
          n1= self.wb.number_display( b'\xc4', attr1 )
          self.assertEqual( decimal.Decimal('4'), n1 )
          n2= self.wb.number_display( b'\xd4', attr1 )
          self.assertEqual( decimal.Decimal('-4'), n2 )

      def test_should_convert_comp3(self):
          attr1= types.SimpleNamespace(
              size_scale_precision = Picture('999v99', False, 5, 0, 2, True, "V") )
          n1= self.wb.number_comp3( b'\x12\x34\x98\x76\x5d', attr1 )
          self.assertEqual( decimal.Decimal('-1234987.65'), n1 )
          attr2= types.SimpleNamespace(
              size_scale_precision = Picture('99999V', False, 5, 0, 0, True, "V") )
          n2= self.wb.number_comp3( b'\x12\x34\x98\x76\x5d', attr2 )
          self.assertEqual( decimal.Decimal('-123498765'), n2 )
        

Test Suite and Runner
=====================

In case we want to build up a larger test suite, we avoid doing
any real work unless this is the main module being executed.

::

  import test
  suite= test.suite_maker( globals() )

  if __name__ == "__main__":
      with test.Logger( stream=sys.stdout, level=logging.DEBUG ):
          logging.info( __file__ )
          unittest.TextTestRunner().run(suite())
          #unittest.main( Test_Dimensional_Access() ) # Specific debugging
