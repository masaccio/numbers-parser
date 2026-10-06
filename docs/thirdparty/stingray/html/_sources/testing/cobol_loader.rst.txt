################################
COBOL Schema Loader Unit Tests
################################

These are unit tests of various parts of the COBOL schema loader.  See :ref:`cobol_loader`.

Overheads
=================

::

  """stingray.cobol.loader Unit Tests."""
  import unittest
  import stingray.cobol.loader
  import weakref
  import io

Lexical Scanner
=================

At a low-level, a parser requires a lexical scanner to break the input into
discrete tokens.  In the case of COBOL, there are mercifully few rules to
follow.

A few test cases will demonstrate that line-ending ``.`` as well as
a simple version of COBOL quoting rules are followed.  The full COBOL language
isn't being parsed, just the part of the language used for DDE's.

::

  class TestLexicalScanner( unittest.TestCase ):
      def test_should_scan( self ):
          copy1= """\
        * COPY1.COB
         01  DETAIL-LINE.
             05                              PIC X(7).
             05  QUESTION                    PIC ZZ.
             05                              PIC X(6).
             05  PRINT-YES                   PIC ZZ.
             05                              PIC X(3).
           SKIP2
             05  PRINT-NO                    PIC ZZ.
             05                              PIC X(6).
             05  NOT-SURE                    PIC ZZ.
             05                              PIC X(7).
  """
          result= ['01', 'DETAIL-LINE', '.', 
          '05', 'PIC', 'X(7)', '.', 
          '05', 'QUESTION', 'PIC', 'ZZ', '.', 
          '05', 'PIC', 'X(6)', '.', 
          '05', 'PRINT-YES', 'PIC', 'ZZ', '.', 
          '05', 'PIC', 'X(3)', '.', 
          '05', 'PRINT-NO', 'PIC', 'ZZ', '.', 
          '05', 'PIC', 'X(6)', '.', 
          '05', 'NOT-SURE', 'PIC', 'ZZ', '.', 
          '05', 'PIC', 'X(7)', '.']
          scanner= stingray.cobol.loader.Lexer().scan( copy1 )
          tokens = list( scanner )
          self.assertEqual( len(result), len(tokens) )
          self.assertEqual( result, tokens )

      def test_should_scan_pic( self ):
          copy2= """\
        * COPY2.COB
         01  DETAIL-LINE.
             05  FANCY-FORMAT                PIC 9(7).99.
             05  SSN-FORMAT                  PIC 999-99-9999.
  """
          result= ['01', 'DETAIL-LINE', '.', 
          '05', 'FANCY-FORMAT', 'PIC', '9(7).99', '.', 
          '05', 'SSN-FORMAT', 'PIC', '999-99-9999', '.']
          scanner= stingray.cobol.loader.Lexer().scan( copy2 )
          tokens = list( scanner )
          self.assertEqual( len(result), len(tokens) )
          self.assertEqual( result, tokens )
        
      def test_should_scan_quotes( self ):
          copy3= """\
        * COPY3.COB
         01  SIMPLE-LINE.
             05  VALUE-1                PIC X(10) VALUE 'STRING    '.
             05  VALUE-2                PIC X(10) VALUE "STRING    ".
  """
          result= [ '01', 'SIMPLE-LINE', '.',
          '05', 'VALUE-1', 'PIC', 'X(10)', 'VALUE', "'STRING    '", '.',
          '05', 'VALUE-2', 'PIC', 'X(10)', 'VALUE', '"STRING    "', '.']
          scanner= stingray.cobol.loader.Lexer().scan( copy3 )
          tokens = list( scanner )
          self.assertEqual( len(result), len(tokens) )
          self.assertEqual( result, tokens )

      def test_should_replace_text( self ):
          copy4= """\
        * COPY4.COB
         01  'XY'-SIMPLE-LINE.
             05  'XY'-VALUE-1                PIC X(10) VALUE 'STRING    '.
             05  'XY'-VALUE-2                PIC X(10) VALUE "STRING    ".
  """
          result= [ '01', 'AB-SIMPLE-LINE', '.',
          '05', 'AB-VALUE-1', 'PIC', 'X(10)', 'VALUE', "'STRING    '", '.',
          '05', 'AB-VALUE-2', 'PIC', 'X(10)', 'VALUE', '"STRING    "', '.']
          scanner= stingray.cobol.loader.Lexer(replacing=[("'XY'","AB")]).scan( copy4 )
          tokens = list( scanner )
          self.assertEqual( len(result), len(tokens) )
          self.assertEqual( result, tokens )
        
Long Lexical Scanner
====================

Some copybooks have junk at the left and right on each line of source.
A separate subclass can handle this.

::

  class TestLongLexicalScanner( unittest.TestCase ):
      def setUp( self ):
          self.copy1= """\
        **************************************************************
         01  REPORT-TAPE-DETAIL-RECORD.                                   
             02  RDT-REC-CODE-BYTES.                                      00000130
         EJECT
                 03  RDT-REC-CODE-KEY              PIC X.                 00000140
          """
          self.scanner= stingray.cobol.loader.Lexer_Long_Lines().scan( self.copy1 )
      def test_should_scan( self ):
          result= ['01', 'REPORT-TAPE-DETAIL-RECORD', '.', 
          '02', 'RDT-REC-CODE-BYTES', '.', 
          '03', 'RDT-REC-CODE-KEY', 'PIC', 'X', '.', 
          ]
          tokens = list( self.scanner )
          self.assertEqual( len(result), len(tokens) )
          self.assertEqual( result, tokens )

Picture Parsing
=====================

A picture clause has it's own "sub-language" for describing an elementary
piece of data.  From the picture clause, we extract a number of features.

..  py:attribute:: final
    final picture with ()'s expanded.

..  py:attribute:: alpha
    boolean alpha = True, numeric = False.

..  py:attribute:: length
    len(final)

..  py:attribute:: scale
    count of "P" positions

..  py:attribute:: precision
    digits to the right of the decimal point

..  py:attribute:: signed
    boolean

..  py:attribute:: decimal
    "." or "V" or None

::

  class TestPictureParser( unittest.TestCase ):
      def test_should_expand( self ):
          pic= stingray.cobol.loader.picture_parser( "X(7)" )
          self.assertEqual( "XXXXXXX", pic.final )
          self.assertTrue( pic.alpha )
          self.assertEqual( 7, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 0, pic.precision )
          self.assertFalse( pic.signed )
          self.assertIsNone( pic.decimal )
      def test_should_handle_z( self ):
          pic= stingray.cobol.loader.picture_parser( "ZZ" )
          self.assertEqual( "ZZ", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 2, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 0, pic.precision )
          self.assertFalse( pic.signed )
          self.assertIsNone( pic.decimal )
      def test_should_handle_9( self ):
          pic= stingray.cobol.loader.picture_parser( "999" )
          self.assertEqual( "999", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 3, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 0, pic.precision )
          self.assertFalse( pic.signed )
          self.assertIsNone( pic.decimal )
      def test_should_handle_complex( self ):
          pic= stingray.cobol.loader.picture_parser( "9(5)V99" )
          self.assertEqual( "9999999", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 7, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 2, pic.precision )
          self.assertFalse( pic.signed )
          self.assertEqual( "V", pic.decimal )
      def test_should_handle_signed( self ):
          pic= stingray.cobol.loader.picture_parser( "S9(7)V99" )
          self.assertEqual( "999999999", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 9, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 2, pic.precision )
          self.assertTrue( pic.signed )
          self.assertEqual( "V", pic.decimal )
      def test_should_handle_db( self ):
          pic= stingray.cobol.loader.picture_parser( "DB9(5).99" )
          self.assertEqual( "DB99999.99", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 10, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 2, pic.precision )
          self.assertTrue( pic.signed )
          self.assertEqual( ".", pic.decimal )
      def test_should_handle_signed_and_v( self ):
          pic= stingray.cobol.loader.picture_parser( "S9(4)V" )
          self.assertEqual( "9999", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 4, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 0, pic.precision )
          self.assertTrue( pic.signed )
          self.assertEqual( "V", pic.decimal )
      def test_should_handle_multiple_repeats( self ):
          pic= stingray.cobol.loader.picture_parser( "9(6).9(3)" )
          self.assertEqual( "999999.999", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 10, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 3, pic.precision )
          self.assertFalse( pic.signed )
          self.assertEqual( ".", pic.decimal )
      def test_should_handle_trailing_repeat_1(self):
          pic= stingray.cobol.loader.picture_parser("SV9(5)")
          self.assertEqual( "99999", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 5, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 5, pic.precision )
          self.assertTrue( pic.signed )
          self.assertEqual( "V", pic.decimal )
      def test_should_handle_trailing_repeat_2(self):
          pic= stingray.cobol.loader.picture_parser("SV9(05)")
          self.assertEqual( "99999", pic.final )
          self.assertFalse( pic.alpha )
          self.assertEqual( 5, pic.length )
          self.assertEqual( 0, pic.scale )
          self.assertEqual( 5, pic.precision )
          self.assertTrue( pic.signed )
          self.assertEqual( "V", pic.decimal )

Usage
=========

A Usage object is attached to a DDE to explain how to decode the bytes
that will be found in the record.  There are many cases in COBOL, but we only
really care about three: DISPLAY, COMP and COMP-3.

::

  class TestUsageDisplay( unittest.TestCase ):
      def setUp( self ):
          self.usage = stingray.cobol.defs.UsageDisplay( "DISPLAY" )
          self.picture= stingray.cobol.loader.Picture( "99999", False, 5, 0, 2, True, "V" )
      def test_should_show_size( self ):
          self.usage.setTypeInfo( self.picture )
          self.assertEqual( 5, self.usage.size() )
          self.assertEqual( "DISPLAY", self.usage.source() )

Note the sizing issue for COMP:

    ======================     =====
    Picture Info               Bytes
    ======================     =====
    if  1<=(int+fract)<=4      2
   
    if  5<=(int+fract)<=9      4 

    if 10<=(int+fract)<=18     8
    ======================     =====
   
::

  class TestUsageComp( unittest.TestCase ):
      def setUp( self ):
          self.usage = stingray.cobol.defs.UsageComp( "COMP" )
      def test_should_show_size_999( self ):
          self.picture= stingray.cobol.loader.Picture( "999", False, 3, 0, 0, True, None )
          self.usage.setTypeInfo( self.picture )
          self.assertEqual( 2, self.usage.size() )
          self.assertEqual( "COMP", self.usage.source() )
      def test_should_show_size_S9_4( self ):
          self.picture= stingray.cobol.loader.Picture( "9999", False, 4, 0, 0, True, "V" )
          self.usage.setTypeInfo( self.picture )
          self.assertEqual( 2, self.usage.size() )
          self.assertEqual( "COMP", self.usage.source() )
      def test_should_show_size_S9_5( self ):
          self.picture= stingray.cobol.loader.Picture( "99999", False, 5, 0, 0, True, "V" )
          self.usage.setTypeInfo( self.picture )
          self.assertEqual( 4, self.usage.size() )
          self.assertEqual( "COMP", self.usage.source() )
    
::

  class TestUsageComp3( unittest.TestCase ):
      def setUp( self ):
          self.usage = stingray.cobol.defs.UsageComp3( "COMP-3" )
          self.picture= stingray.cobol.loader.Picture( "9999999", False, 7, 0, 2, True, "V" )
      def test_should_show_size( self ):
          self.usage.setTypeInfo( self.picture )
          self.assertEqual( 4, self.usage.size() )
          self.assertEqual( "COMP-3", self.usage.source() )


Allocation
==========

There are three kinds: group (i.e., a header under a group), successor, 
and redefines.

A Redefines object gets the offset information from another named element.
A Group object 
for an item.  A non-redefined element is "real": it has a proper size and offset.

An item with a REDEFINES
clause is an alias for another element.   The offset comes from the other element.  The size is reported as zero to simplify offset calculations.

Here's a Mock DDE which can have a redefines clause, or be referenced by
a redefines clause.

::

  class MockDDE:
      def __init__( self, **kw ):
          self.children= []
          self.top= weakref.ref(self) # default
          self.__dict__.update( kw )
      def get( self, name ):
          return [c for c in self.children if c.name == name][0]
      def addChild( self, child ):
          self.children.append( child )
          child.parent= weakref.ref(self)
          child.top= self.top
        
There are two non-redefinies cases: Successor and Group. The DDE stands for itself.  The size is
as computed.  The offset is as generated by the :py:func:`cobol.defs.setSizeAndOffset` function.

::

  class TestAllocation_Group( unittest.TestCase ):
      def setUp( self ):
          self.parent= MockDDE( size=123, allocation=None, 
              occurs=stingray.cobol.defs.Occurs(), totalSize=123 )
          self.group = stingray.cobol.defs.Group()
          self.dde= MockDDE( size=123, allocation=self.group, 
              occurs=stingray.cobol.defs.Occurs(), totalSize=123 )
          self.parent.addChild( self.dde )
      def test_should_get_size_and_offset( self ):
          self.group.resolve( self.dde )
          self.assertEqual( 123, self.dde.allocation.totalSize() )
          self.assertEqual( 13, self.dde.allocation.offset(13) )
          self.assertIs( self.dde, self.dde.allocation.dde() )

  class TestAllocation_Successor( unittest.TestCase ):
      def setUp( self ):
          self.parent= MockDDE( size=12, allocation=None, 
              occurs=stingray.cobol.defs.Occurs(), totalSize=12 )
          self.prev= MockDDE( size=5, 
              occurs=stingray.cobol.defs.Occurs(), totalSize=5 )
          self.parent.addChild( self.prev )
          self.successor = stingray.cobol.defs.Successor( self.prev )
          self.dde= MockDDE( size=7, allocation=self.successor, occurs=stingray.cobol.defs.Occurs(), totalSize=7 )
          self.parent.addChild( self.dde )
      def test_should_get_size_and_offset( self ):
          self.successor.resolve( self.dde )
          self.assertEqual( 7, self.dde.allocation.totalSize() )
          self.assertEqual( 5, self.dde.allocation.offset(5) )
          self.assertIs( self.dde, self.dde.allocation.dde() )

Redefines is a reference to another DDE. The other DDE is located by the resolver pass.
The size and offset come from the other element. 

::

  class TestAllocation_Rdefines( unittest.TestCase ):
      def setUp( self ):
          self.parent= MockDDE( name= "TOP", size=123, allocation=None, occurs=stingray.cobol.defs.Occurs(), totalSize=12 )
          self.otherdde= MockDDE( name= "SOME-NAME", size=100, offset=23 )
          self.parent.addChild( self.otherdde )
        
          self.redefines_some_name = stingray.cobol.defs.Redefines( "SOME-NAME" )
          self.dde= MockDDE( name= "REDEF", size=47, allocation=self.redefines_some_name, occurs=stingray.cobol.defs.Occurs() )
          self.parent.addChild( self.dde )
      def test_should_get_size_and_offset( self ):
          """Since this redefines something else, it doesn't contribute any size."""
          self.redefines_some_name.resolve( self.dde )
          self.assertEqual( 0, self.dde.allocation.totalSize() )
          self.assertEqual( 23, self.dde.allocation.offset(13) )
          self.assertIs( self.dde, self.dde.allocation.dde() )


Size And Offset Function
========================

The idea is that this function does a depth-first traversal to accumulate the 
size in each group-level DDE.  

Size and Offset Mocks
----------------------

To test size and offset function, we need a number of mocks.
We need to mock a DDE.   Plus,
we need to mock Usage and Redefines classes, also.

We also mock the iterable protocol of the DDE.

::

  class MockDDE2:
      def __init__( self, **kw ):
          self.occurs= stingray.cobol.defs.Occurs()
          self.allocation= None
          self.children= []
          self.parent= None
          self.level= None
          self.name= "FILLER"
          self.picture= ""
          self.sizeScalePrecision= ()
          self.dimensionality= ()
          self.indent= 0
          self.__dict__.update( kw )
          self.size= self.usage.length
      def __repr__( self ):
          rc= ""
          if isinstance(self.allocation,stingray.cobol.defs.Redefines):
              rc= "REDEFINES {0}".format( self.allocation.name )
          oc= str(self.occurs)
          return "{0} {1} {2} {3} {4}.  {5} {6}".format(
              self.level, self.name, self.picture, rc, oc,
              self.offset, self.size,  )
      def addChild( self, child ):
          child.parent= weakref.ref(self)
          child.top= self.top
          self.children.append( child )
          child.indent= self.indent+1
      def get( self, name ):
          for d in self:
              if d.name == name: return d
      def pathTo( self ):
          return self.name
      def __iter__( self ):
          yield self
          for c in self.children:
              for c_i in c:
                  yield c_i

  class MockUsage:
      class MockCell:
          def __init__( self, buffer, workbook ):
              self.buffer= buffer
              self.workbook= workbook
      def __init__( self, source, **kw ):
          self.source_= source
          self.typeInfo= None
          self.length= 0
          self.__dict__.update( kw )
      def size( self ):
          return self.length
      def setTypeInfo( self, picture ):
          self.typeInfo= picture
          self.length= picture.length
      def source( self ):
          return self.source_
      def create_func( self, raw, workbook, attr ):
          return MockUsage.MockCell(raw, workbook, attr)


  def mock_resolver( top ):
      for aDDE in top:
          if aDDE.allocation:
              aDDE.allocation.resolve( aDDE )
          else: # Not setup by test.
              #aDDE.allocation= stingray.cobol.defs.Group()
              pass


Size and Offset Cases
------------------------

Exercise the size and offset function with a number of cases.

::

  class TestFlatSizeOffset( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE2( level='01', allocation=stingray.cobol.defs.Group(), picture=None, usage=MockUsage("") )
          self.top.top= weakref.ref(self.top)
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=3 ), 
                  allocation=stingray.cobol.defs.Group(), picture="XXX" ) )
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=5 ), 
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="99999" ) )
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=7 ), 
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="9999999" ) )
          mock_resolver( self.top )
      def test_should_set_size( self ):
          stingray.cobol.defs.setSizeAndOffset( self.top )
          self.assertEqual( 15, self.top.size )
          self.assertEqual( 0, self.top.children[0].offset )
          self.assertEqual( 3, self.top.children[0].size )
          self.assertEqual( 3, self.top.children[1].offset )
          self.assertEqual( 5, self.top.children[1].size )
          self.assertEqual( 8, self.top.children[2].offset )
          self.assertEqual( 7, self.top.children[2].size )

::

  class TestNestedSizeOffset( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE2( level='01', allocation=stingray.cobol.defs.Group(), picture=None, usage=MockUsage("") )
          self.top.top= weakref.ref(self.top)
          self.top.parent= None
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=3 ),  
                  allocation=stingray.cobol.defs.Group(), picture="XXX" ) )
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=5 ), 
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="99999" ) )
          sub_group= MockDDE2( level='05', 
              allocation=stingray.cobol.defs.Successor(self.top.children[-1]), usage=MockUsage("") )
          self.top.addChild( sub_group )
          sub_group.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=7 ), 
                  allocation=stingray.cobol.defs.Group(), picture="9999999" ) )
          sub_group.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=9 ), 
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="XXXXXXXXX" ) )
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=11 ), 
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="XXXXXXXXXXX" ) )
          mock_resolver( self.top )
      def test_should_set_size( self ):
          stingray.cobol.defs.setSizeAndOffset( self.top )
          self.assertEqual( 35, self.top.size )
          self.assertEqual( 0, self.top.children[0].offset )
          self.assertEqual( 3, self.top.children[0].size )
          self.assertEqual( 3, self.top.children[1].offset )
          self.assertEqual( 5, self.top.children[1].size )
          self.assertEqual( 8, self.top.children[2].offset )
          self.assertEqual( 16, self.top.children[2].size )
          self.assertEqual( 24, self.top.children[3].offset )
          self.assertEqual( 11, self.top.children[3].size )
          self.assertEqual( 8, self.top.children[2].children[0].offset )
          self.assertEqual( 7, self.top.children[2].children[0].size )
          self.assertEqual( 15, self.top.children[2].children[1].offset )
          self.assertEqual( 9, self.top.children[2].children[1].size )


::

  class TestRedefinesSizeOffset( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE2( level='01', allocation=stingray.cobol.defs.Group(), usage=MockUsage("") )
          self.top.top= weakref.ref(self.top)
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=3 ),  
                  allocation=stingray.cobol.defs.Group(), picture="XXX" ) )
          sub_group_1= MockDDE2( level='05', name="GROUP-1", 
              allocation=stingray.cobol.defs.Successor(self.top.children[-1]), usage=MockUsage("") )
          self.top.addChild( sub_group_1 )
          sub_group_2= MockDDE2( level='05', name="GROUP-2", 
              allocation=stingray.cobol.defs.Redefines(refers_to=sub_group_1, name="GROUP-1",), usage=MockUsage("") )
          self.top.addChild( sub_group_2 )
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=5 ), 
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="XXXXX" ) )

          sub_group_1.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=5 ), 
                  allocation=stingray.cobol.defs.Group(), picture="99999" ) )
          sub_group_1.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=7 ), 
                  allocation=stingray.cobol.defs.Successor(sub_group_1.children[-1]), picture="9999999" ) )
            
          sub_group_2.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=9 ), 
                  allocation=stingray.cobol.defs.Group(), picture="XXXXXXXXX" ) )
          sub_group_2.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=3 ), 
                  allocation=stingray.cobol.defs.Successor(sub_group_2.children[-1]), picture="XXX" ) )
            
          mock_resolver( self.top )
      def test_should_set_size( self ):
          stingray.cobol.defs.setSizeAndOffset( self.top )
          self.assertEqual( 20, self.top.size )
          self.assertEqual( 0, self.top.children[0].offset )
          self.assertEqual( 3, self.top.children[0].size )
          self.assertEqual( 3, self.top.children[1].offset )
          self.assertEqual( 12, self.top.children[1].size )
          self.assertEqual( 3, self.top.children[2].offset )
          self.assertEqual( 12, self.top.children[2].size )
          self.assertEqual( 15, self.top.children[3].offset )
          self.assertEqual( 5, self.top.children[3].size )
          self.assertEqual( 3, self.top.children[1].children[0].offset )
          self.assertEqual( 5, self.top.children[1].children[0].size )
          self.assertEqual( 8, self.top.children[1].children[1].offset )
          self.assertEqual( 7, self.top.children[1].children[1].size )
          self.assertEqual( 3, self.top.children[2].children[0].offset )
          self.assertEqual( 9, self.top.children[2].children[0].size )
          self.assertEqual( 12, self.top.children[2].children[1].offset )
          self.assertEqual( 3, self.top.children[2].children[1].size )

::
           
  class TestOccursSizeOffset( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE2( level='01', allocation=stingray.cobol.defs.Group(), picture=None, usage=MockUsage("") )
          self.top.top= weakref.ref(self.top)
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=3 ),  
                  allocation=stingray.cobol.defs.Group(), picture="XXX" ) )
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=5 ), 
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="99999", 
                  occurs=stingray.cobol.defs.OccursFixed(7) ) )
          mock_resolver( self.top )
      def test_should_set_size( self ):
          stingray.cobol.defs.setSizeAndOffset( self.top )
          self.assertEqual( 38, self.top.totalSize )
          self.assertEqual( 0, self.top.children[0].offset )
          self.assertEqual( 3, self.top.children[0].size )
          self.assertEqual( 3, self.top.children[1].offset )
          self.assertEqual( 35, self.top.children[1].totalSize )
          self.assertEqual( 5, self.top.children[1].size )
          self.assertEqual( 7, self.top.children[1].occurs.number(None) )

::

  class TestGroupOccursSizeOffset( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE2( level='01', allocation=stingray.cobol.defs.Group(), picture=None, usage=MockUsage("") )
          self.top.top= weakref.ref(self.top)
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=3 ),  
                  allocation=stingray.cobol.defs.Group(), picture="XXX" ) )
          sub_group_1= MockDDE2( level='05', name="GROUP-1", 
              allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture=None, 
              occurs=stingray.cobol.defs.OccursFixed(4), usage=MockUsage("") )
          self.top.addChild( sub_group_1 )
          sub_group_1.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=5 ), 
                  allocation=stingray.cobol.defs.Group(), picture="99999", 
                  occurs=stingray.cobol.defs.OccursFixed(7) ) )
          mock_resolver( self.top )
      def test_should_set_size( self ):
          stingray.cobol.defs.setSizeAndOffset( self.top )
          self.assertEqual( 7*4*5+3, self.top.totalSize )
          self.assertEqual( 0, self.top.children[0].offset )
          self.assertEqual( 3, self.top.children[0].size )
          self.assertEqual( 3, self.top.children[1].offset )
          self.assertEqual( 7*4*5, self.top.children[1].totalSize )
          self.assertEqual( 7*5, self.top.children[1].size )
          self.assertEqual( 4, self.top.children[1].occurs.number(None) )
          self.assertEqual( 3, self.top.children[1].children[0].offset )
          self.assertEqual( 35, self.top.children[1].children[0].totalSize )
          self.assertEqual( 5, self.top.children[1].children[0].size )
          self.assertEqual( 7, self.top.children[1].children[0].occurs.number(None) )
      def test_should_be_repeatable( self ):
          stingray.cobol.defs.setSizeAndOffset( self.top )
          self.assertEqual( 7*4*5+3, self.top.totalSize )
          stingray.cobol.defs.setSizeAndOffset( self.top )
          self.assertEqual( 7*4*5+3, self.top.totalSize )
          stingray.cobol.defs.setSizeAndOffset( self.top )
          self.assertEqual( 7*4*5+3, self.top.totalSize )

Resolve Redefines
===========================

We need to test the :py:func:`cobol.loader.resolver` function.

::
 

  class TestRedefinesNameResolver( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE2( level='01', allocation=stingray.cobol.defs.Group(), picture=None, usage=MockUsage("") )
          self.top.top= weakref.ref(self.top)
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=3 ),  
                  allocation=stingray.cobol.defs.Group(), picture="XXX" ) )
          self.sub_group_1= MockDDE2( level='05', name="GROUP-1", 
              allocation=stingray.cobol.defs.Successor(self.top.children[-1]), usage=MockUsage("") )
          self.top.addChild( self.sub_group_1 )
          self.sub_group_2= MockDDE2( level='05', name="GROUP-2", 
              allocation=stingray.cobol.defs.Redefines(refers_to=None, name="GROUP-1",), usage=MockUsage("") )
          self.top.addChild( self.sub_group_2 )
          self.top.addChild(
              MockDDE2( level='05', usage=MockUsage( "", length=5 ),  
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="XXXXX" ) )

          self.sub_group_1.addChild(
              MockDDE2( level='10', usage=MockUsage( "",length=5 ), 
                  allocation=stingray.cobol.defs.Group(), picture="99999" ) )
          self.sub_group_1.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=7 ), 
                  allocation=stingray.cobol.defs.Successor(self.sub_group_1.children[-1]), picture="9999999" ) )
            
          self.sub_group_2.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=9 ), 
                  allocation=stingray.cobol.defs.Group(), picture="XXXXXXXXX" ) )
          self.sub_group_2.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=3 ), 
                  allocation=stingray.cobol.defs.Successor(self.sub_group_2.children[-1]), picture="XXX" ) )
            
          self.top.search= { 
              'GROUP-1': self.sub_group_1, 
              'GROUP-2': self.sub_group_2,
              }
            
      def test_should_resolve( self ):
          stingray.cobol.defs.resolver( self.top )
          self.assertIs( self.sub_group_1, self.sub_group_2.allocation.refers_to )
        
        
Set Dimensionality
==========================

The :py:func:`stingray.cobol.defs.setDimensionality` function pushes the OCCURS information down to each 
child of the OCCURS. This builds the effective dimensionality of the lowest-level
elements.

..  code-block:: cobol

          * COPY3.COB
           01  SURVEY-RESPONSES.
               05  QUESTION-NUMBER         OCCURS 10 TIMES.
                   10  RESPONSE-CATEGORY     OCCURS 3 TIMES.
                       15  ANSWER                          PIC 99.

::

  class Test_Dimensionality( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE2( level='01', name='SURVEY-RESPONSES', 
              allocation=stingray.cobol.defs.Group(), picture=None, offset=0, size=60, 
              occurs=stingray.cobol.defs.Occurs(), usage=MockUsage("") )
          self.top.top= weakref.ref(self.top)
          self.top.parent= None
          self.group_05 = MockDDE2( level='05', name='QUESTION-NUMBER', 
              allocation=stingray.cobol.defs.Group(), usage=MockUsage(""), 
              occurs=stingray.cobol.defs.OccursFixed(10), offset=0, totalSize=60, size=6 )
          self.top.addChild( self.group_05 )
          self.group_10 = MockDDE2( level='10', name='RESPONSE-CATEGORY', 
              allocation=stingray.cobol.defs.Group(), usage=MockUsage(""), 
              occurs=stingray.cobol.defs.OccursFixed(3), offset=0, totalSize=6, size=2 )
          self.group_05.addChild( self.group_10 )
          self.group_15 = MockDDE2( level='15', name='ANSWER', 
              allocation=stingray.cobol.defs.Group(), picture="99", 
              occurs=stingray.cobol.defs.Occurs(), offset=0, totalSize=2, size=2, usage=MockUsage("99") )
          self.group_10.addChild( self.group_15 )
      def test_should_set_dimensions( self ):
          stingray.cobol.defs.setDimensionality( self.top )
          self.assertEqual( 1, self.top.occurs.number(None) )
          self.assertEqual( (), self.top.dimensionality )
          self.assertEqual( 10, self.group_05.occurs.number(None) )
          self.assertEqual( (self.group_05,), self.group_05.dimensionality )
          self.assertEqual( 3, self.group_10.occurs.number(None) )
          self.assertEqual( (self.group_05,self.group_10), self.group_10.dimensionality )
          self.assertEqual( 1, self.group_15.occurs.number(None) )
          self.assertEqual( (self.group_05,self.group_10), self.group_15.dimensionality )

DDE Construction Methods
===========================

DDE class has many methods.  They fit into three functionality categories.

1.  Building the DDE, adding children, visiting.

2.  Getting elements by simple name.  Computing path names.  Getting elements by path name.

3.  Getting a run-time value from a buffer, accounting for indexes.  This 
    is DDE Access and is tested separately.

The first two areas are tested here.

::

  class TestDDEMethods( unittest.TestCase ):
      def setUp( self ):
          self.top= stingray.cobol.defs.DDE( level='01', name='TOP', usage=MockUsage("") )
          self.top.top= weakref.ref(self.top)
          self.group= stingray.cobol.defs.DDE( level='05', name='GROUP', usage=MockUsage("")  )
          self.element= stingray.cobol.defs.DDE( level='10', name='ELEMENT', pic='X(5)', usage=MockUsage("")  )
          self.top.addChild( self.group )
          self.group.addChild( self.element )
      def test_structure( self ):
          self.assertIs( self.group, self.top.children[0] )
          self.assertIs( self.element, self.group.children[0] )
      def test_path_name( self ):
          self.assertEqual( "TOP.GROUP.ELEMENT", self.element.pathTo() )
          self.assertIs( self.element, self.top.getPath("TOP.GROUP.ELEMENT") )
      def test_get( self ):
          g= self.top.get( "GROUP" )
          self.assertIs( self.group, g )
          e= g.get( "ELEMENT" )
          self.assertIs( self.element, e )

Parsing Single DDE
=====================

The parser handles 11 different clauses.  Additionally, it makes a single
DDE element as well as making a composite DDE record.  A parser depends on
a lexical scanner.  However, since our scanner is essentially an iterator,
we don't need to do very much to mock it.

To test just the :py:class:`cobol.loader.RecordFactory` 
in isolation, we need to provide mocks for a large number of dependenices.

..  parsed-literal::

        self.redefines_class= Redefines
        self.non_redefines_class= NonRedefines
        self.display_class= UsageDisplay
        self.comp_class= UsageComp
        self.comp3_class= UsageComp3

Note that there are numerous syntax errors which we do not test for.
Ideally, a DDE clause is used in working software, and passes through a COBOL compiler.  This isn't *lint for COBOL*.

We have three test cases: things the parser finds (and keeps), things
the parser skips gracefully, and things we can't cope with.

::

  class TestParser( unittest.TestCase ):
      def setUp( self ):
          self.parser= stingray.cobol.loader.RecordFactory()
          #self.parser.redefines_class= MockRedefines
          #self.parser.non_redefines_class= MockNonRedefine
          self.parser.display_class= MockUsage
          self.parser.comp_class= MockUsage
          self.parser.comp3_class= MockUsage
      def test_should_parse_group( self ):
          source= ( "01", "GROUP", "." ) 
          dde= next(self.parser.dde_iter(iter(source)))
          self.assertEqual( '01', dde.level )
          self.assertEqual( 'GROUP', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertIsNone( dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
      def test_should_parse_filler( self ):
          source= ( "05", "PIC", "X(10)", "." ) 
          dde= next(self.parser.dde_iter(iter(source)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'FILLER', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "X(10)", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
      def test_should_parse_name( self ):
          source= ( "05", "ELEMENTARY", "PIC", "X(10)", "." ) 
          dde= next(self.parser.dde_iter(iter(source)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'ELEMENTARY', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "X(10)", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
      def test_should_parse_occurs( self ):
          src1= ( "05", "ELEMENTARY-OCCURS", "PIC", "S9999", "OCCURS", "5", "TIMES", "." ) 
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'ELEMENTARY-OCCURS', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 5, dde.occurs.number(None) )
          self.assertEqual( "S9999", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
          src2= ( "05", "GROUP-OCCURS", "OCCURS", "7", "TIMES", "INDEXED", "BY", "IRRELEVANT", "." ) 
          dde= next(self.parser.dde_iter(iter(src2)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'GROUP-OCCURS', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 7, dde.occurs.number(None) )
          self.assertIsNone( dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
      def test_should_parse_picture( self ):
          """Details of picture clause tested separately."""
          src1= ( "05", "ELEMENTARY-PIC", "PIC", "S9999.999", "." ) 
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'ELEMENTARY-PIC', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "S9999.999", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
      def test_should_parse_redefines( self ):
          src1= ( "05", "ELEMENTARY-REDEF", "PIC", "S9999", "REDEFINES", "SOME-NAME", "." ) 
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'ELEMENTARY-REDEF', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "S9999", dde.picture )
          self.assertEqual( "SOME-NAME", dde.allocation.name )
          self.assertEqual( "", dde.usage.source() )
          src2= ( "05", "GROUP-REDEF", "OCCURS", "7", "TIMES", "REDEFINES", "ANOTHER-NAME", "." ) 
          dde= next(self.parser.dde_iter(iter(src2)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'GROUP-REDEF', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 7, dde.occurs.number(None) )
          self.assertIsNone( dde.picture )
          self.assertEqual( "ANOTHER-NAME", dde.allocation.name )
          self.assertEqual( "", dde.usage.source() )
      def test_should_parse_usage( self ):
          src1= ( "05", "USE-DISPLAY", "PIC", "S9999", "USAGE", "DISPLAY", "." ) 
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'USE-DISPLAY', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "S9999", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "DISPLAY", dde.usage.source() )
          src2= ( "05", "USE-COMP-3", "PIC", "S9(7)V99", "USAGE", "COMP-3", "." ) 
          dde= next(self.parser.dde_iter(iter(src2)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'USE-COMP-3', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "S9(7)V99", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "COMP-3", dde.usage.source() )
          src3= ( "05", "USE-COMP-3-ALT", "PIC", "S9(9)V99", "COMP-3", "." ) 
          dde= next(self.parser.dde_iter(iter(src3)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'USE-COMP-3-ALT', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "S9(9)V99", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "COMP-3", dde.usage.source() )
      def test_should_parse_depending_on_1( self ):
          src1= ( "05", "DEP-1", "OCCURS", "1", "TO", "5", "TIMES", "DEPENDING", "ON", "ODO-1", "." ) 
          self.parser.lex= iter(src1)
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'DEP-1', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( "ODO-1", dde.occurs.name ) # number(None) )
          self.assertEqual( None, dde.picture )
          self.assertIsNone( dde.allocation )
      def test_should_parse_depending_on_2( self ):
          src1= ( "05", "DEP-2", "OCCURS", "TO", "5", "TIMES", "DEPENDING", "ON", "ODO-2", "." ) 
          self.parser.lex= iter(src1)
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'DEP-2', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( "ODO-2", dde.occurs.name ) # number(None) )
          self.assertEqual( None, dde.picture )
          self.assertIsNone( dde.allocation )

Syntax which is silently skipped.

::


  class TestParserSkip( unittest.TestCase ):
      def setUp( self ):
          self.parser= stingray.cobol.loader.RecordFactory()
      def test_should_skip_blank( self ):
          src1= ( "05", "BLANK-ZERO-1", "PIC", "X(10)", "BLANK", "ZERO", "." ) 
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'BLANK-ZERO-1', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "X(10)", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
          self.assertIsNone( dde.initValue )
          src2= ( "05", "BLANK-ZERO-2", "PIC", "X(10)", "BLANK", "WHEN", "ZEROES", "." ) 
          dde= next(self.parser.dde_iter(iter(src2)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'BLANK-ZERO-2', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "X(10)", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
          self.assertIsNone( dde.initValue )
      def test_should_skip_justified( self ):
          src1= ( "05", "JUST-RIGHT-1", "PIC", "X(10)", "JUST", "RIGHT", "." ) 
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'JUST-RIGHT-1', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "X(10)", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
          self.assertIsNone( dde.initValue )
          src2= ( "05", "JUST-RIGHT-2", "PIC", "X(10)", "JUSTIFIED", "RIGHT", "." ) 
          dde= next(self.parser.dde_iter(iter(src2)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'JUST-RIGHT-2', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "X(10)", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
          self.assertIsNone( dde.initValue )
      def test_should_skip_value( self ):
          src1= ( "05", "VALUE-1", "PIC", "X(8)", "VALUE", "'10 CHARS'", "." ) 
          dde= next(self.parser.dde_iter(iter(src1)))
          self.assertEqual( '05', dde.level )
          self.assertEqual( 'VALUE-1', dde.name )
          self.assertEqual( 0, len(dde.children) )
          self.assertEqual( 1, dde.occurs.number(None) )
          self.assertEqual( "X(8)", dde.picture )
          self.assertIsNone( dde.allocation )
          self.assertEqual( "", dde.usage.source() )
          self.assertIsNone( dde.initValue )

..  todo:: Test EXTERNAL, GLOBAL as Skipped Words, too.

A few clauses may be relevant for some kinds of DDE's.  These
can impact the encoding of the bytes.  We don't parse
them, however, because they are rarely seen in the wild.

::

  class TestParserException( unittest.TestCase ):
      def setUp( self ):
          self.parser= stingray.cobol.loader.RecordFactory()
      def test_should_fail_renames( self ):
          src1= ( "66", "BLANK-ZERO-1", "RENAMES", "SOME-NAME", "." ) 
          self.assertWarns( UserWarning, next, self.parser.dde_iter(iter(src1)) )
          # Alternative RENAMES
          # self.assertRaises( stingray.cobol.defs.UnsupportedError, ... 
      def test_should_fail_sign( self ):
          src1= ( "05", "SIGN-1", "PIC", "X(10)", "LEADING", "SIGN", "." ) 
          self.parser.lex= iter(src1)
          self.assertRaises( stingray.cobol.defs.UnsupportedError, next, self.parser.dde_iter(iter(src1)) )
          src2= ( "05", "SIGN-2", "PIC", "X(10)", "TRAILING", "." ) 
          self.parser.lex= iter(src2)
          self.assertRaises( stingray.cobol.defs.UnsupportedError, next, self.parser.dde_iter(iter(src2)) )
          src3= ( "05", "SIGN-3", "PIC", "X(10)", "SIGN", "IS", "SEPARATE", "." ) 
          self.parser.lex= iter(src3)
          self.assertRaises( stingray.cobol.defs.UnsupportedError, next, self.parser.dde_iter(iter(src3)) )
      def test_should_fail_synchronized( self ):
          src1= ( "05", "SYNC-1", "PIC", "X(10)", "SYNCHRONIZED", "." ) 
          self.parser.lex= iter(src1)
          self.assertRaises( stingray.cobol.defs.UnsupportedError, next, self.parser.dde_iter(iter(src1)) )

Parsing Complete DDE
=====================

Parsing a complete DDE is a multi-step dance.  First, parse all the clauses.
Then apply some standard functions to resolve the REDEFINES clauses
as well as compute sizes and offsets.  

To test just the :py:class:`cobol.loader.RecordFactory` 
in isolation, we need to provide mocks for a large number of dependenices.

..  parsed-literal::

        # Built during the parsing
        self.redefines_class= Redefines
        self.non_redefines_class= NonRedefines
        self.display_class= UsageDisplay
        self.comp_class= UsageComp
        self.comp3_class= UsageComp3



We use dependency injection here to tease apart Redefines and Usage classes.
We can use existing mocks for this.

Here's a test of the "main" method for parsing a record.
Note that :py:meth:`stingray.cobol.loader.RecordFactory.makeRecord` is
iterable, so we make a list and take the first element.

We could as easily  ``next()`` it to get the first element.


Test the parser

::

  class TestCompleteParser( unittest.TestCase ):
      def setUp( self ):
          self.parser= stingray.cobol.loader.RecordFactory()
          self.parser.display_class= MockUsage
          self.parser.comp_class= MockUsage
          self.parser.comp3_class= MockUsage
      def test_should_parse( self ):
          copy1= """
                * COPY1.COB
                 01  DETAIL-LINE.
                     05                              PIC X(7).
                     05  QUESTION                    PIC ZZ.
                     05                              PIC X(6).
                     05  PRINT-YES                   PIC ZZ.
          """
          self.lex= ['01', 'DETAIL-LINE', '.', 
          '05', 'PIC', 'X(7)', '.', 
          '05', 'QUESTION', 'PIC', 'ZZ', '.', 
          '05', 'PIC', 'X(6)', '.', 
          '05', 'PRINT-YES', 'PIC', 'ZZ', '.']
          dde= list(self.parser.makeRecord( iter( self.lex ) ))[0]
          self.assertEqual( "01", dde.top().level )
          self.assertEqual( "DETAIL-LINE", dde.top().name )
          self.assertEqual( 4, len(dde.top().children) )
          self.assertEqual( "", dde.top().usage.source() )
          c0= dde.top().children[0]
          self.assertEqual( "05", c0.level )
          self.assertEqual( "FILLER", c0.name )
          self.assertEqual( "X(7)", c0.picture )
          self.assertEqual( "", c0.usage.source() )
          self.assertEqual( 1, c0.occurs.number(None) )
          c1= dde.top().children[1]
          self.assertEqual( "05", c1.level )
          self.assertEqual( "QUESTION", c1.name )
          self.assertEqual( "ZZ", c1.picture )
          self.assertEqual( "", c1.usage.source() )
          self.assertEqual( 1, c1.occurs.number(None) )
          c2= dde.top().children[2]
          self.assertEqual( "05", c2.level )
          self.assertEqual( "FILLER", c2.name )
          self.assertEqual( "X(6)", c2.picture )
          self.assertEqual( "", c2.usage.source() )
          self.assertEqual( 1, c2.occurs.number(None) )
          c3= dde.top().children[3]
          self.assertEqual( "05", c3.level )
          self.assertEqual( "PRINT-YES", c3.name )
          self.assertEqual( "ZZ", c3.picture )
          self.assertEqual( "", c3.usage.source() )
          self.assertEqual( 1, c3.occurs.number(None) )
          self.assertEqual( 5, len(list(dde) ) )

Schema Maker
=====================

A DDE is tranformed into a flat schema from the highly-nested DDE structure
via the :py:func:`stingray.cobol.loader.make_schema` function.

::

  class TestNestedSchemaMaker( unittest.TestCase ):
      def setUp( self ):
          self.top= MockDDE2( level='01', name="TOP", 
              allocation=stingray.cobol.defs.Group(), picture=None, size=35, offset=0, usage=MockUsage( "" ) )
          self.top.top= weakref.ref(self.top)
          self.top.addChild(
              MockDDE2( level='05', name="FIRST", usage=MockUsage( "", length=3 ),  
                  allocation=stingray.cobol.defs.Group(), picture="XXX", size=3, offset=0 ) )
          self.top.addChild(
              MockDDE2( level='05', name="SECOND", usage=MockUsage( "", length=5 ), 
                  allocation=stingray.cobol.defs.Successor(self.top.children[-1]), picture="99999", size=5, offset=3 ) )
          sub_group= MockDDE2( level='05', name="GROUP", 
              allocation=stingray.cobol.defs.Group(), picture=None, size=16, offset=8, usage=MockUsage( "" ) )
          self.top.addChild( sub_group )
          sub_group.addChild(
              MockDDE2( level='10', name="INNER", usage=MockUsage( "", length=7 ), 
                  allocation=stingray.cobol.defs.Group(), picture="9999999", size=8, offset=7 ) )
          sub_group.addChild(
              MockDDE2( level='10', usage=MockUsage( "", length=9 ), 
                  allocation=stingray.cobol.defs.Successor(sub_group.children[-1]), picture="XXXXXXXXX", size=9, offset=15) )
          self.top.addChild(
              MockDDE2( level='05', name="LAST", usage=MockUsage( "", length=11 ), 
                  allocation=stingray.cobol.defs.Successor(sub_group.children[-1]), picture="XXXXXXXXXXX", size=11, offset=24 ) )
      def test_should_set_size( self ):
          stingray.cobol.defs.setSizeAndOffset( self.top )
          schema = stingray.cobol.loader.make_schema( [self.top] )
          #DEBUG# print( schema )
          self.assertEqual( "TOP", schema[0].name )
          self.assertEqual( 0, schema[0].position )
          self.assertEqual( 0, schema[0].offset )
          self.assertEqual( 35, schema[0].size )
          self.assertEqual( "FIRST", schema[1].name )
          self.assertEqual( 1, schema[1].position )
          self.assertEqual( 0, schema[1].offset )
          self.assertEqual( 3, schema[1].size )
          self.assertEqual( "SECOND", schema[2].name )
          self.assertEqual( 2, schema[2].position )
          self.assertEqual( 3, schema[2].offset )
          self.assertEqual( 5, schema[2].size )
          self.assertEqual( "LAST", schema[6].name )
          self.assertEqual( 6, schema[6].position )
          self.assertEqual( 24, schema[6].offset )
          self.assertEqual( 11, schema[6].size )

Another Variation on Loading DDE's
==================================

Here's a class which does some loading. The processing is a bit convoluted, 
but this is a possible use case, and must work.

::

  class SchemaLoaderClass (stingray.cobol.loader.COBOLSchemaLoader):
      lexer_class= stingray.cobol.loader.Lexer_Long_Lines

      def load( self ):
          fixed_dde_list= list()
          unused_dde_list= list()
          dde_list=list(self.parser.makeRecord( self.lexer.scan(self.source) ))
          hdr_trl_record_copybook="""\
  000010 01  HEADER-TRAILER-RECORD.
  000020            05  HDR-REC-TYPE                        PIC X(6).
  000030            05  FILLER                              PIC X(5).
  000040            05  HDR-TRL                             PIC X(8).
  000050            05  FILLER                              PIC X(1).
  000060            05  HDR-DATE1                           PIC X(10).
  000070            05  FILLER                              PIC X(1).
  000080            05  HDR-JULIAN-DATE1                    PIC X(7).
  000090            05  FILLER                              PIC X(1).
  000100            05  HDR-SOME-NUMBER1                    PIC X(8).
  000110            05  FILLER                              PIC X(1).
  000120            05  HDR-DATE2                           PIC X(10).
  000130            05  FILLER                              PIC X(1).
  000140            05  HDR-SOME-NUMBER2                    PIC X(7).
  000150            05  FILLER                              PIC X(1).
  000140            05  HDR-SOME-NUMBER3                    PIC X(4).
  000160            05  HDR-REST-OF-LINE            PIC X(30).
  """
          dde_list+=list(self.parser.makeRecord( self.lexer.scan(hdr_trl_record_copybook) ))
          for dde in dde_list:
              if dde.name =='HEADER-TRAILER-RECORD':
                  fixed_dde_list+= [dde]
              elif dde.name == 'ANOTHER-RECORD':
                  base_dde= [dde]
              else:
                  unused_dde_list+= [dde]
          schema= stingray.cobol.loader.make_schema( base_dde )
          fixed_schema= stingray.cobol.loader.make_schema( fixed_dde_list )
          return schema, fixed_schema
        
Confirm that this class really works

::

  class TestSchemaLoaderClass(unittest.TestCase):
      def test_should_load(self):
          another_record ="""\
        01  ANOTHER-RECORD.
            05  FILLER PIC X(10).
          """
          example = io.StringIO(another_record)
          slc= SchemaLoaderClass(example)
          schema, fixed_schema= slc.load()
          self.assertTrue( any( f.name == "ANOTHER-RECORD" for f in schema ) )
          self.assertTrue( any( f.name == "HEADER-TRAILER-RECORD" for f in fixed_schema ) )


Test Suite and Runner
=====================

In case we want to build up a larger test suite, we avoid doing
any real work unless this is the main module being executed.

::

  import test
  suite= test.suite_maker( globals() )

  if __name__ == "__main__":
      print( __file__ )
      unittest.TextTestRunner(verbosity=1).run(suite())
