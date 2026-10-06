..  #!/usr/bin/env python3

.. _`workbook_number09`:


Apple iWorks Numbers '09 Workbook
-----------------------------------

The Stingray model of sheet/row/cell structure does not
easily fit the Numbers sheet/table/row/cell structure.
How can we handle the extra layer of names introduced by 
Numbers?

Option 1: navigation hierarchy.

    Workbook ➞ new layer (Numbers "Workspace") ➞ Sheet (Numbers "Table") ➞ Row ➞ Cell

Option 2: navigation hierarchy.

    Combine (Workspace,Table) into a 2-tuple, and call this a "sheet" name when working
    with Numbers documents.
   
    This will fit with Stingray acceptably. 
   
The imports required to process this kind of file.

::

  import logging
  import pprint
  import xml.etree.cElementTree as dom
  import zipfile
  import datetime
  import decimal

  from stingray.workbook.base import Workbook
  import stingray.sheet
  import stingray.cell

..  py:module:: workbook.numbers09
       
The iWork Numbers 09 format is a Zip file with an XML document inside it.
There may be slight variations between native Numbers '09 and Numbers '13 doing
a "save as" in Numbers '09 format. It's not clear; we haven't done
exhaustive checking.

Numbers '13 is entirely different. See :ref:`workbook_number13`.

..  py:class:: Numbers09_Workbook

    Extract sheets, rows and cells from a Numbers '09 format file.
       
    The ``.numbers`` "file" is a ZIP file.
       
    The :file:`index.xml` element the interesting part of the archive.

    In addition to the superclass attributes, some additional unique
    attributes are introduced here.
       
    ..  py:attribute:: zip_archive
   
        A zip archive for this file.
       
    ..  py:attribute:: workspace
   
        The "workspaces": pages with tables inside them.

::

  class Numbers09_Workbook( Workbook ):
      """Mac OS X Numbers Workbook for iWork 09.
      """
      NUMBERS_NS = {
      "ls":"http://developer.apple.com/namespaces/ls",
      "sf":"http://developer.apple.com/namespaces/sf",
      "sfa":"http://developer.apple.com/namespaces/sfa",
      }
      row_debug= False
      def __init__( self, name, file_object=None ):
          """Prepare the workbook for reading.
          :param name: File name
          :param file_object: Optional file-like object. Ignored for v3.2 numbers files.
          """
          super().__init__( name, file_object )
          self.zip_archive= zipfile.ZipFile( file_object or name, "r" )
          self._prepare()

As preparation for reading these files, we locate all the sheet names
and all the number styles.

::

      def _prepare( self ):
          """Locate sheets/tables and styles."""
          root= dom.parse( self.zip_archive.open('index.xml') ).getroot()
          self._locate_sheets(root)
          self._get_styles(root)

Locating all the sheets is a matter of doing an XPath search for
:samp:`workspace-array/workspace` and getting the ``workspace-name`` attribute
from the  :samp:`<table name="{name}">` tags.

Within each workspace we have to find :samp:`page-info/tabular-info/tabular-model` to 
get the tables within the workspaces.

::

      def _locate_sheets( self, root ):
          """Create ``workspace_table`` map from name to workspace and table."""
          self.workspace= dict()

          ws_name_attr= dom.QName( self.NUMBERS_NS["ls"], 'workspace-name' )
          name_attr= dom.QName( self.NUMBERS_NS["sf"], 'name' )
          workspace_array= root.find("ls:workspace-array", namespaces=self.NUMBERS_NS )
          for workspace in workspace_array.findall('.//ls:workspace', namespaces=self.NUMBERS_NS ):
              # Populate tables within this workspace.
              tables= dict()
              page_info = workspace.find('ls:page-info', namespaces=self.NUMBERS_NS)
              for tabular_info in page_info.findall('.//sf:tabular-info', namespaces=self.NUMBERS_NS):
                  tabular_model = tabular_info.find( 'sf:tabular-model', namespaces=self.NUMBERS_NS)
                  tables[ tabular_model.get(name_attr) ] = tabular_model
              self.workspace[ workspace.get(ws_name_attr) ]= workspace, tables

Locate a "data source" within the XML document. Create ``Cell`` instances.

::

      def _datasource( self, grid ):
          """The data source for cell values within a grid.
          This yields each individual cell value, transformed into
          string, Decimal, datetime.
          """            
          datasource = grid.find('.//sf:datasource', namespaces=self.NUMBERS_NS)
          for cell_doc in datasource:
              yield self.cell( cell_doc )
          # or return map( self.cell, datasource )

..  py:method:: Numbers09_Workbook.cell( cell )

    Create a ``Cell`` instance from the decoded data.

::

      def cell( self, cell ):
          logging.debug( dom.tostring(cell) )

          date_tag= dom.QName( self.NUMBERS_NS["sf"], 'd' )
          date_attr= dom.QName( self.NUMBERS_NS["sf"], 'cell-date' )
          formula_tag= dom.QName( self.NUMBERS_NS["sf"], 'f' )
          s_attr= dom.QName( self.NUMBERS_NS["sf"], 's' )
          v_attr= dom.QName( self.NUMBERS_NS["sf"], 'v' )
          general_tag= dom.QName( self.NUMBERS_NS["sf"], 'g' )
          number_tag= dom.QName( self.NUMBERS_NS["sf"], 'n' )
          text_tag= dom.QName( self.NUMBERS_NS["sf"], 't' )
          o_tag= dom.QName( self.NUMBERS_NS["sf"], 'o' )
          span_tag= dom.QName( self.NUMBERS_NS["sf"], 's' )
          bool_tag= dom.QName( self.NUMBERS_NS["sf"], 'b' )
          popup_menu_tag= dom.QName( self.NUMBERS_NS["sf"], 'pm' )
          IDREF_attr= dom.QName( self.NUMBERS_NS["sfa"], 'IDREF' )
          ID_attr= dom.QName( self.NUMBERS_NS["sfa"], 'ID' )
          fs_attr= dom.QName( self.NUMBERS_NS["sf"],"fs")

          if cell.tag == date_tag: 
              seconds= int(cell.attrib[date_attr])
              epoch= datetime.datetime(2001, 1, 1)
              delta= datetime.timedelta( seconds=seconds )
              theDate= epoch + delta
              return stingray.cell.DateCell( theDate, self )
            
          elif cell.tag == formula_tag: # formula or error.
              s= cell.get(s_attr)
              fo= cell.find('sf:fo', namespaces=self.NUMBERS_NS)
              # Numeric Result? What about non-numeric results?
              r= cell.find('sf:r', namespaces=self.NUMBERS_NS)
              if r:
                  # Result:
                  rn= r.find('sf:rn', namespaces=self.NUMBERS_NS)
                  try:
                      value_txt= rn.attrib[v_attr]
                      value= self._to_decimal( value_txt, s )
                  except KeyError as ex:
                      #self._cell_warning("Formula with no value", cell)
                      value= self._to_decimal( '0', s )
                  return stingray.cell.NumberCell( value, self )
              else:
                  # Error: 
                  #self._cell_warning("Formula error", cell)
                  value= "#Error in {0}".format(fo.get(fs_attr))
                  return stingray.cell.ErrorCell( value, self )
                
          elif cell.tag == general_tag: # General?
              return stingray.cell.EmptyCell( '', self )
          elif cell.tag == number_tag: # Number
              value= self._decode_number( cell )
              return stingray.cell.NumberCell( value, self )
          elif cell.tag == o_tag: #??
              self._cell_warning("Unknown cell type", cell)
              return stingray.cell.EmptyCell( '', self )
          elif cell.tag == span_tag: # Span?
              self._cell_warning("Unknown cell type", cell)
              return stingray.cell.EmptyCell( '', self )
          elif cell.tag == text_tag: # Text
              value= self._decode_text( cell )
              return stingray.cell.TextCell( value, self )
          elif cell.tag == bool_tag: # Boolean
              value= self._decode_number( cell )
              return stingray.cell.BooleanCell( value, self )
          elif cell.tag == popup_menu_tag: # popup menu
              # TODO:: Better Xpath query: ``menu-choices/*[@ID='name']``
              value= None # In case we can't find anything.
              selected= cell.find('sf:proxied-cell-ref', namespaces=self.NUMBERS_NS)
              name= selected.get(IDREF_attr)
              mc= cell.find('sf:menu-choices', namespaces=self.NUMBERS_NS)
              for t in mc:
                  if t.get(ID_attr) == name:
                      # t's tag cold end in Could be "t", or "n".
                      if t.tag.endswith('t'): # Text
                          value= self._decode_text( t )
                          return stingray.cell.TextCell( value, self )
                      elif t.tag.endswith('n'): # Number
                          value= self._decode_number( t )
                          return stingray.cell.NumberCell( value, self )
                      else:
                          raise Exception( "Unknown popup menu {0}".format(dom.tostring(cell)))
          else:
              raise Exception( "Unknown cell {0}".format( dom.tostring(cell) ) )

Some lower-level conversions. 

::

      def _to_decimal( self, value_txt, style_id ):
          """Convert a given numeric value_text using the named style.

          TODO: From the style, get the number of decimal places, use that to
          build a string version of the float value.
          """
          fdp_attr= dom.QName( self.NUMBERS_NS["sf"], 'format-decimal-places' )
          fs_attr= dom.QName( self.NUMBERS_NS["sf"], 'format-string' )
          cell_style= self.cell_style.get(style_id)
          #print( "TO_DECIMAL", value_txt, style_id, "=", cell_style )

          fs= None # cell_style.get(fs_attr) # Doesn't seem correct
          fdp= None # cell_style.get(fdp_attr) # Doesn't seem correct
        
          # Transform fs into proper Python format, otherwise, use the number of 
          # decimal places.
        
          if fs is not None:
              fmt= self._rewrite_fmt( fs )
              #print( "Decimal: {{0:{0}}}.format({1}) = ".format( fmt, value_txt ), end="" )
              value= decimal.Decimal( "{:{fmt}}".format(float(value_txt), fmt=fmt) )
              #print( value )
              return value
          elif fdp is not None:
              #fmt= "{{0:.{0}f}}".format(fdp)
              value= decimal.Decimal( "{:.{fdp}f}".format(float(value_txt), fdp=fdp) )
              #print( "Decimal: {0}.format({1}) = {2!r}".format( fmt, value_txt, value ) )
              return value
          else:
              value= decimal.Decimal( value_txt )
              #print( "Decimal: {0} = {1!r}".format( value_txt, value ) )
          return value

      def _decode_text( self, cell ):
          """Decode a <t> tag's value."""
          sfa_s_attr= dom.QName( self.NUMBERS_NS["sfa"], 's' )
          ct= cell.find( 'sf:ct', namespaces=self.NUMBERS_NS )
          value= ct.get(sfa_s_attr)
          if value is None:
              value= "\n".join( cell.itertext() )
          return value

      def _decode_number( self, cell ):
          """Decode a <n> tag's value, applying the style."""
          s_attr= dom.QName( self.NUMBERS_NS["sf"], 's' )
          v_attr= dom.QName( self.NUMBERS_NS["sf"], 'v' )
          s= cell.get(s_attr)
          cell_style= self.cell_style.get(s)
          try:
              value_txt= cell.attrib[v_attr]
              value= self._to_decimal( value_txt, s )
          except KeyError as ex:
              #self._cell_warning("Number with no value", cell)
              value= self._to_decimal( '0', s )
          return value


The styles are also important because we can use them to parse the numbers more
precisely.

::

      def _get_styles( self, root ):
          """Get the styles."""
          ID_attr= dom.QName( self.NUMBERS_NS["sfa"], 'ID' )
          ident_attr= dom.QName( self.NUMBERS_NS["sf"], 'ident' )
          parent_ident_attr= dom.QName( self.NUMBERS_NS["sf"], 'parent-ident' )

          self.cell_style= {}
          for cs in root.findall('.//sf:cell-style', namespaces=self.NUMBERS_NS):
              #print( "STYLE", dom.tostring(cs) )
              ID= cs.get(ID_attr)
              ident= cs.get(ident_attr)
              parent_ident= cs.get(parent_ident_attr)
              property_number_format= cs.find('.//sf:SFTCellStylePropertyNumberFormat', namespaces=self.NUMBERS_NS)
              if property_number_format is None:
                  if parent_ident is not None:
                      self.cell_style[ID]= self.cell_style[parent_ident]
              else:
                  number_format= property_number_format.find('sf:number-format', namespaces=self.NUMBERS_NS)
                  if number_format is None:
                      if parent_ident is not None:
                          self.cell_style[ID]= self.cell_style[parent_ident]
                  else:
                      self.cell_style[ID]= number_format.attrib
                      if ident is not None:
                          self.cell_style[ident]= number_format.attrib
                  #print( ID, self.cell_style.get(ID,None) )
        
Rewrite a number format from Numbers to Python

::

      def _rewrite_fmt( self, format_string ):
          """Parse the mini-language: '#,##0.###;-#,##0.###' is an example.
          This becomes "{:10,.3f}"
          """
          positive, _, negative = format_string.partition(";")
          fmt= negative or positive
          digits= len(fmt)
          comma= "," if "," in fmt else ""
          whole, _, frac= fmt.partition(".")
          precision= len(frac)
          return "{digits}{comma}.{precision}f".format(
              digits= digits, comma=comma, precision=precision )
                    
..  py:method:: Numbers09_Workbook.sheets( )

    Return a list of "sheets" (actually underlying tables.)

    The "sheets" are ``[ (`` *workspace*\ `,` *table* ``), ... ]`` pairs.

    Picking a sheet involves matching a two-part name: (workspace, table).

::

      def sheets( self ):
          """Build "sheet" names from workspace/table"""
          sheet_list= []
          for w_name in self.workspace:
              ws, tables = self.workspace[w_name]
              for t_name in tables:
                  sheet_list.append( (w_name, t_name) )
          return sheet_list

..  py:method:: Numbers09_Workbook.rows_of( sheet )

    Iterator through all rows of a sheet.

::

      def rows_of( self, sheet ):
          """Iterator over rows.

          Two parallel traversals:

          Internal iterator over grid/datasource/* has d, t, n, pm, g, o and s
              yields individual cell values.

          Iterator over grid/rows/grid-row may have ``nc``, number of columns in that row.
              Each grid-row fetches a number of cell values to assemble a row.
              Row's may be variable length (sigh) but padded to the number of columns
              specified in the grid.
            
          :param sheet: a Sheet object to retrieve rows from.
          """
          self.log.debug( "rows of {0}: {1}".format(sheet, sheet.name) )
          ws_name, t_name = sheet.name
          ws, tables= self.workspace[ws_name]
          tabular_model= tables[t_name]
        
          grid= tabular_model.find( 'sf:grid', namespaces=self.NUMBERS_NS )
          numrows_attr= dom.QName( self.NUMBERS_NS["sf"], 'numrows' )
          numcols_attr= dom.QName( self.NUMBERS_NS["sf"], 'numcols' )
          numrows = int(grid.attrib[numrows_attr])
          numcols = int(grid.attrib[numcols_attr])
        
          nc_attr= dom.QName( self.NUMBERS_NS["sf"], 'nc' )
        
          datasource= iter( self._datasource(grid) )
        
          rows = grid.find('sf:rows', namespaces=self.NUMBERS_NS)
          for n, r in enumerate(rows.findall( 'sf:grid-row', namespaces=self.NUMBERS_NS )):
              #print( "ROW", dom.tostring(r) )
              self.debug_row= n
              # Is this really relevant for Numbers '09?
              nc= int(r.get(nc_attr,numcols)) 
              try:
                  row= [ next(datasource) for self.debug_col in range(nc) ]
              except StopIteration as e:
                  pass # Last row will exhaust the datasource.
              if len(row) == numcols:
                  yield row
              else:
                  yield row + (numcols-nc)*[None]
