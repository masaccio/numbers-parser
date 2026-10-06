..  #!/usr/bin/env python3

.. _`workbook_xlsx`:


XLSX or XLSM Workbook
----------------------

::

  import logging
  import pprint
  import xml.etree.cElementTree as dom
  import re
  import zipfile
  from collections import defaultdict

  from stingray.workbook.base import Workbook
  import stingray.sheet
  import stingray.cell

..  py:module:: workbook.xlsx

..  py:class:: XLSX_Workbook

    Extract sheets, rows and cells from an XLSX format file.

    We're opening a ZIP archive and parsing the various XML documents
    that we find therein.

    The :py:class:`ElementTree` incremental parser provides
    parse "events" for specific tags, allowing for lower-memory parsing of
    the sometimes large XML documents.

    See http://effbot.org/zone/element-iterparse.htm

    The class as a whole defines some handy constants like XML namespaces
    and a pattern for parsing Cell ID's to separate the letters from the numbers.

    In addition to the superclass attributes, some additional unique
    attributes are introduced here.
       
    ..  py:attribute:: zip_archive
   
        A zip archive for this file.

    ..  py:attribute:: strings_dict
   
        The strings in this workbook
       
::

  class XLSX_Workbook( Workbook ):
      """ECMA Standard XLSX or XLSM documents.
      Locate sheets and rows within a given sheet.

      See http://www.ecma-international.org/publications/standards/Ecma-376.htm
      """
      # Relevant subset of namespaces used
      XLSX_NS = {
      "main":"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel":"http://schemas.openxmlformats.org/package/2006/relationships",
      }
      cell_id_pat = re.compile( r"(\D+)(\d+)" )
      def __init__( self, name, file_object=None ):
          """Prepare the workbook for reading.
          :param name: File name
          :param file_object: Optional file-like object.  If omitted, the named file is opened.
          """
          super().__init__( name, file_object )
          self.zip_archive= zipfile.ZipFile( file_object or name, "r" )
          self._prepare()

The are two preparation steps required for reading these files.  First, the
sheets must be located.  This involves resolving internal rID numbers.
Second, the shared strings need to be loaded into memory.

::

      def _prepare( self ):
          self._locate_sheets()
          self._get_shared_strings()

Locate all sheets involves building a :py:data:`name_to_id` mapping and  and :py:data:`id_to_member` mapping.  This allows is to map the
user-oriented name to an id and the id to the XLSX zipfile member.

::

      def _locate_sheets( self ):
          """Locate the name to id mapping and the id to member mapping.
          """
          # 1a. Open "workbook.xml" member.
          workbook_zip= self.zip_archive.getinfo("xl/workbook.xml")
          workbook_doc= dom.parse( self.zip_archive.open(workbook_zip) )
          # 1b. Get a dict of sheet names and their rIdx values.
          key_attr_id= 'name'
          val_attr_id= dom.QName( self.XLSX_NS['r'], 'id' )
          self.name_to_id = dict(
              ( s.attrib[key_attr_id], s.attrib[val_attr_id] )
              for s in workbook_doc.findall("*/main:sheet", namespaces=self.XLSX_NS)
          )
          logging.debug( self.name_to_id )

          # 2a. Open the "_rels/workbook.xml.rels" member
          rels_zip= self.zip_archive.getinfo("xl/_rels/workbook.xml.rels")
          rels_doc= dom.parse( self.zip_archive.open(rels_zip) )
          # 2b. Get a dict of rIdx to Target member name
          logging.debug( dom.tostring( rels_doc.getroot() ) )
          key_attr_id= 'Id'
          val_attr_id= 'Target'
          self.id_to_member = dict(
              ( r.attrib[key_attr_id], r.attrib[val_attr_id] )
              for r in rels_doc.findall("rel:Relationship", namespaces=self.XLSX_NS)
          )
          logging.debug( self.id_to_member )

Get Shared Strings walks a fine line.  Ideally, we'd like to parse
the document and simply use ``itertext`` to gather all of the text
within a given string instance (:samp:`<si>`) tag.  **However.**

In practice, these documents can be so huge that they don't fit
in memory comfortably.  We rely on incremental parsing via the ``iterparse`` function.

::

      def _get_shared_strings( self ):
          """Build ``strings_dict`` with all shared strings.
          """
          self.strings_dict= defaultdict(str)
          count= 0
          text_tag= dom.QName( self.XLSX_NS['main'], "t" )
          string_tag= dom.QName( self.XLSX_NS['main'], "si" )
          # 1. Open the "xl/sharedStrings.xml" member
          sharedStrings_zip= self.zip_archive.getinfo("xl/sharedStrings.xml")
          for event, element in dom.iterparse(
              self.zip_archive.open( sharedStrings_zip ), events=('end',) ):
              logging.debug( event, element.tag )
              if element.tag == text_tag:
                  self.strings_dict[ count ]+= element.text
              elif element.tag == string_tag:
                  count += 1
              element.clear()
          logging.debug( self.strings_dict )

The shared strings may be too massive for in-memory incremental parsing.
We can create a temporary extract file to handle this case. Here's
the kind of code we might use.

..  parsed-literal::

    with tempfile.TemporaryFile( ) as temp:
        self.zip_archive.extract( sharedStrings_mbr, temp.filename )
        for event, element in dom.iterparse( temp.filename ):
            *process event and element*

..  py:method:: XLSX_Workbook.sheets( )

    Return the list of sheets for this workbook.


::

      def sheets( self ):
          return self.name_to_id.keys()

Translate a col-row pair from :samp:`({letter}, {number})` 
to proper 0-based Python index of :samp:`({row}, {col})`.

::

      @staticmethod
      def make_row_col( col_row_pair ):
          col, row = col_row_pair
          cn = 0
          for char in col_row_pair[0]:
              cn = cn*26 + (ord(char)-ord("A")+1)
          return int(row), cn-1

We can build an eager :py:class:`sheet.Row` or a  :py:class:`sheet.LazyRow` from the available data.
The eager :py:class:`sheet.Row` is built from :py:class:`cell.Cell` objects.  
The :py:class:`sheet.LazyRow` delegates the creation
of :py:class:`cell.Cell` objects to :py:meth:`Workbook.row_get`.

This uses an incremental parser, also.  There are four kinds of tags that
have to be located.

-   :samp:`<row>{row}</row>`, end event.  Finish (and yield) the row of cells.
    Since XLSX is sparse, missing empty cells must be filled in.

-   :samp:`<c t="{type}" r="{id}">{cell}</c>`.

    -   Start event for ``c``.  Get the cell type and id.  Empty the value accumulator.

    -   End event for ``c``.  Save the accumulated value.  This allows the cell to have
        mixed content model.

-   :samp:`<v>{value}</v>`, end event. Use the :py:meth:`cell` method to track down
    enough information to build the Cell instance.

..  py:method:: XLSX_Workbook.rows_of( sheet )

    Iterate through rows of the given sheet.


::

      def rows_of( self, sheet ):
          """Iterator over rows as a list of Cells for a named worksheet."""
          # 1. Map user name to member.
          rId = self.name_to_id[sheet.name]
          self.sheet_member_name = self.id_to_member[rId]
          # 2. Open member.
          sheet_zip= self.zip_archive.getinfo("xl/"+self.sheet_member_name)
          self.row= {}
          # 3. Assemble each row, allowing for missing cells.
          row_tag= dom.QName(self.XLSX_NS['main'], "row")
          cell_tag= dom.QName(self.XLSX_NS['main'], "c")
          value_tag= dom.QName(self.XLSX_NS['main'], "v")
          format_tag= dom.QName(self.XLSX_NS['main'], "f")
        
          for event, element in dom.iterparse(
              self.zip_archive.open(sheet_zip), events=('start','end') ):
              logging.debug( element.tag, repr(element.text) )
              if event=='end' and element.tag == row_tag:
                  # End of row: fill in missing cells
                  if self.row.keys():
                      data= stingray.sheet.Row( sheet, *(
                          self.row.get(i, stingray.cell.EmptyCell('', self))
                          for i in range(max(self.row.keys())+1) ) )
                      yield data
                  else:
                      yield stingray.sheet.Row( sheet )
                  self.row= {}
                  element.clear()
              elif event=='end' and element.tag == cell_tag:
                  # End of cell: consolidate the final string
                  self.row[self.row_col[1]] = self.value
                  self.value= stingray.cell.EmptyCell( '', self )
              elif event=='start' and element.tag == cell_tag:
                  # Start of cell: collect a string in pieces.
                  self.cell_type= element.attrib.get('t',None)
                  self.cell_id = element.attrib['r']
                  id_match = self.cell_id_pat.match( self.cell_id )
                  self.row_col = self.make_row_col( id_match.groups() )
                  self.value= stingray.cell.EmptyCell( '', self )
              elif event=='end' and element.tag == value_tag:
                  # End of a value; what type was it?
                  self.value= self.cell( element )

              elif event=='end' and element.tag == format_tag:
                  pass # A format string
              else:
                  pass
                  logging.debug( "Ignoring", end="" ) # Numerous bits of structure exposed.
                  logging.debug( dom.tostring(element) )

..  py:method:: XLSX_Workbook.row_get( row, attribute )

    Low-level get of a particular attribute from the given row.


::

      def row_get( self, row, attribute ):
          """Create a Cell from the row's data."""
          return row[attribute.position]

..  py:method:: XLSX_Workbook.cell( row, element )

    Build a subclass of :py:class:`cell.Cell` from the current value tag content plus the
    containing cell type information.

::

      def cell( self, element ):
          """Create a proper :py:class:`cell.Cell` subclass from cell and value information."""
          logging.debug( self.cell_type, self.cell_id, element.text )
          if self.cell_type is None or self.cell_type == 'n':
              try:
                  return stingray.cell.NumberCell( float(element.text), self )
              except ValueError:
                  print( self.cell_id, element.attrib, element.text )
                  return None
          elif self.cell_type == "s":
              try:
                  # Shared String?
                  return stingray.cell.TextCell( self.strings_dict[int(element.text)], self )
              except ValueError:
                  # Inline String?
                  logging.debug( self.cell_id, element.attrib, element.text )
                  return stingray.cell.TextCell( element.text, self )
              except KeyError:
                  # Not a valid shared string identifier?
                  logging.debug( self.cell_id, element.attrib, element.text )
                  return stingray.cell.TextCell( element.text, self )
          elif self.cell_type == "b":
              return stingray.cell.BooleanCell( float(element.text), self )
          elif self.cell_type == "d":
              return stingray.cell.FloatDateCell( float(element.text), self )
          elif self.cell_type == "e":
              return stingray.cell.ErrorCell( element.text, self )
          else:
              # 'str' (formula), 'inlineStr' (string), 'e' (error)
              print( self.cell_type, self.cell_id, element.attrib, element.text )
              logging.debug( self.strings_dict.get(int(element.text)) )
              return None
