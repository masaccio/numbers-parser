..  #!/usr/bin/env python3

.. _`workbook_number13`:


Apple iWorks Numbers '13 Workbook
----------------------------------

See :ref:`workbook_number09` for a note on the naming of sheets.

Here are the modules we need to import to process these Workbooks.

::

  import logging
  import pprint
  import zipfile
  import datetime
  import decimal
  import os
  import struct

  from stingray.workbook.base import Workbook
  import stingray.sheet
  import stingray.cell
  from stingray.snappy import Snappy
  from stingray.protobuf import Archive_Reader, Message

..  py:module:: workbook.numbers13
       
..  todo:: Additional Numbers13_Workbook Feature
   
    Translate Formula and Formula error to Text
   
The iWork 13 format is a directory with an ``index.zip`` file. The ZIP contains
a number of ``.IWA`` files. Each ``.IWA`` is compressed using the Snappy protocol.
The uncompressed data is messages in Protobuf format.

We could depend on other Snappy and Protobuf implementations. We provide
our own fall-back implementation, instead.
See :ref:`other_modules` for more information on these two supporting modules.

..  py:class:: Numbers13_Workbook

    Extract sheets, rows and cells from a Numbers '13 format file.
       
    The ``.numbers`` "file" is a directory bundle or package.
       
    The :file:`index.zip` file is the interesting part of the bundle.
   
    In addition to the superclass attributes, some additional unique
    attributes are introduced here.
       
    ..  py:attribute:: archive
   
        The snappy archive for this file.


::

  class Numbers13_Workbook( Workbook ):
      """Mac OS X Numbers Workbook for iWork '13.
      """
      def __init__( self, name, file_object=None ):
          """Prepare the workbook for reading.
        
          :param name: File name
          :param file_object: Ignored for iWork13 files.
        
          We might be able to use the file's internal handle to open 
          it as a proper directory. But we don't.
          """
          super().__init__( name, None )
          self.archive= self._load_archive( name )
   
Read the archive to get the serialized messages. This method deserializes
all of the protobuf-encoded messages. It's a shabby stand-in for proper protobuf
processing.

One thing we *could* do is to refactor this method into a bunch of methods, 
each of which is tied to a specific class of message. That would parallel the
way protobuf really works.
                    
::

      @staticmethod
      def _load_archive( filename ):
          """Extract all the protobuf-serialized Archived messages.
          We don't actually need to read **all** of them.
          We really only need Index/Document.iwa, Index/CalculationEngine.iwa, and
          all Index/Tables/*.iwa. But that's almost everything.
        
          :param filename: File name
          """
          log= logging.getLogger( "load_archive" )
          snappy= Snappy()
          reader= Archive_Reader()
          archive=dict()
          with zipfile.ZipFile( os.path.join( filename, "index.zip" ) ) as index:
              for n in index.namelist():
                  log.info( n )
                  with index.open( n ) as member:
                      data= snappy.decompress( member )
                      log.debug( "{0} bytes".format(len(data)) )
                      for id, m in reader.archive_iter( data ):
                          log.debug( "{0:4d}: {1}".format( id, m ) )
                          archive[id]= m
                          if m.name_ == "TN.DocumentArchive":
                              m.sheets= [Message("Reference", sheet) for sheet in m[1]]
                              log.debug( "{0} {1}".format(m.name_, m.sheets) )
                          elif m.name_ == "TN.SheetArchive": 
                              m.name= bytes(m[1][0]).decode("UTF-8")
                              m.drawable_infos= [Message('Reference', ref) for ref in m[2]]
                              log.debug( "{0} {1} {2}".format(m.name_, m.name, m.drawable_infos) )
                          elif m.name_ == "TST.TableInfoArchive": 
                              m.super= Message( 'DrawableArchive', m[1][0] )
                              m.tableModel= Message('Reference', m[2][0])
                              log.debug( "{0} {1} {2}".format(m.name_, m.super, m.tableModel) )
                          elif m.name_ == "TST.TableModelArchive": 
                              m.data_store= Message( "DataStore", m[4][0] )
                              m.table_id= bytes(m[1][0])
                              m.table_name= bytes(m[8][0]).decode("UTF-8")
                              log.debug( "{0} {1} {2}".format(m.name_, m.table_name, m.data_store) )
                              m.data_store.tiles= [Message("TileStorage", tile) for tile in m.data_store[3]]
                              log.debug( "{0} {1}".format(m.name_,m.data_store.tiles) )
                              for ts in m.data_store.tiles:
                                  ts.tiles= [Message("TileStorage.Tile", tile) for tile in ts[1]]
                                  for tile in ts.tiles:
                                      tile.id= tile[1][0]
                                      tile.ref= Message('Reference', tile[2][0])
                                      log.debug( "{0} {1} {2} {3}".format(m.name_, tile, tile.id, tile.ref) )
                              m.data_store.stringTable= [Message('Reference', string) for string in m.data_store[4]]
                              m.data_store.formulaTable= [Message('Reference', formula) for formula in m.data_store[6]]
                              m.data_store.formulaErrorTable= [Message('Reference', error) for error in m.data_store[12]]
                              log.debug( "DataStore stringTable {0}, formulaTable {1}, formulaErrorTable {2}".format( 
                                  m.data_store.stringTable, m.data_store.formulaTable,
                                  m.data_store.formulaErrorTable) )
                          elif m.name_ == "TST.TableDataList":
                              m.listType= m[1][0]
                              m.nextListID= m[2][0]
                              m.entries= [Message('ListEntry', entry) for entry in m[3]]
                              log.debug( "{0} {1} {2}".format( m.name_, m.listType, m.nextListID ) )
                              for entry in m.entries:
                                  entry.key= entry[1][0]
                                  try:
                                      entry.string= bytes(entry[3][0])
                                  except IndexError:
                                      entry.string= b''
                                  try:
                                      entry.formula= Message('FormulaArchive', entry[5][0])
                                      entry.formula.AST_node_array= Message("ASTNodeArrayArchive", entry.formula[1][0])
                                      entry.formula.AST_node_array.AST_node= [Message("AST_node", n) for n in entry.formula.AST_node_array[1]]
                                  except IndexError:
                                      entry.formula= None
                                  log.debug( "ListEntry {0} {1} {2}".format( entry, entry.string, entry.formula ) )
                          elif m.name_ == "TST.Tile":
                              m.rowInfos= [ Message("TileRowInfo", row_info) for row_info in m[5] ]
                              for row_info in m.rowInfos:
                                  row_info.tileRowIndex= row_info[1][0]
                                  row_info.cellCount= row_info[2][0]
                                  row_info.cellStorageBuffer= row_info[3][0]
                                  row_info.cellOffsets= row_info[4][0]
                                  log.debug( "tileRowIndex {0} cellCount {1}".format(row_info.tileRowIndex, row_info.cellCount) )
                                  format= "<{0}h".format(len(row_info.cellOffsets)//2)
                                  offsets= struct.unpack( format, bytes(row_info.cellOffsets) )
                                  log.debug( "{0} {1}".format( m.name_, offsets ) )
                                  row_info.cells= dict()
                                  count= row_info.cellCount
                                  for col, offset in enumerate( offsets ):
                                      if offset == -1: 
                                          log.debug( "{0} {1} {2}".format( m.name_, col, offset ) )
                                          row_info.cells[col]= (None,None)
                                          continue
                                      # A little needless copying to simplify the elif sequence.
                                      cell_raw= row_info.cellStorageBuffer[offset:offset+32]
                                      version, celltype = struct.unpack( "<hbx", bytes(cell_raw[:4]) )
                                      if celltype == 2: # Number
                                          data= struct.unpack( "<IIIdi", bytes(cell_raw[4:28]) )
                                      elif celltype == 3: # Text (in the associated TableDataList)
                                          data= struct.unpack( "<IIIi", bytes(cell_raw[4:20]) )
                                      elif celltype == 5: # Date
                                          data= struct.unpack( "<IIIdi", bytes(cell_raw[4:28]) )
                                      elif celltype == 6: # Boolean
                                          data= struct.unpack( "<IIId", bytes(cell_raw[4:24]) )
                                      elif celltype == 8: # Error
                                          data= struct.unpack( "<hhII", bytes(cell_raw[4:16]) )
                                      else: 
                                          raise Exception( "Unsupported {0}".format(celltype) )
                                      log.debug( "{0} {1} {2} {3} {4}".format( m.name_, col, offset, celltype, data ) )
                                      row_info.cells[col]= (celltype,data[3])
                                      count -= 1
                                      if count == 0: break
                            
          return archive

Once we've decoded the archive, we can fetch sheets, tables, rows and cells.

We start with document.iwa and get the message with an id of 1.
This is the TN.DocumentArchive and it lists the TN.SheetArchive by id.

Each TN.SheetArchive has drawable_infos references to TST.TableInfoArchive.
TST.TableInfoArchive has reference to TST.TableModelArchive (and a super reference to SheetArchive.)
TST.TableModelArchive references a DataStore.

TST.DataStore.stringTable references a TST.TableDataList. This has string literals
  or formulas (or styles) referenced by cells.

TST.DataStore.tiles references a TST.Tile.
  The Tile has some bytes which contain the cells as hard-to-parse raw structures.
  This includes text cells with references to the data lists string literals.
  It includes number, date and boolean cells with numeric-looking data.

..  py:method:: Numbers13_Workbook.sheets( )

    Return a list of "sheets" (actually underlying tables.)

    The "sheets" are ``[ (`` *workspace*\ `,` *table* ``), ... ]`` pairs.

    Picking a sheet involves matching a two-part name: (workspace, table).

::

      def sheets( self ):
          sheet_list= []
          document = self.archive[1]
          for sheet_ref in document.sheets:
              sheet= self.archive[sheet_ref[1][0]]
              for table_ref in sheet.drawable_infos:
                  tableinfo= self.archive[table_ref[1][0]]
                  tablemodel= self.archive[tableinfo.tableModel[1][0]]
                  sheet_list.append(  (sheet.name, tablemodel.table_name) )
          return sheet_list
    
The proto files include this.

..  parsed-literal::

    enum CellType {
      genericCellType = 0;
      spanCellType = 1;
      numberCellType = 2;
      textCellType = 3;
      formulaCellType = 4;
      dateCellType = 5;
      boolCellType = 6;
      durationCellType = 7;
      formulaErrorCellType = 8;
      automaticCellType = 9;
    }


::

      def _cell( self, cell, strings, formulae, errors ):
          """Given a decoded row cell message, pluck out the relevant data.
          Look into the collection of strings or formulae for the data object.
          :param cell: cell message
          :param strings: dict of strings
          :param formulae: dict of formulae 
          :param errors: dict of errors
          :returns: Cell object
          """
          celltype, value = cell
          if celltype == 0: # What does "generic" mean?
              v= stingray.cell.NumberCell( value, self )
          elif celltype == 1 or celltype is None:
              v= stingray.cell.EmptyCell( "", self )
          elif celltype == 2:
              v= stingray.cell.NumberCell( value, self )
          elif celltype in (3, 9): # Text and Rich Text
              v= stingray.cell.TextCell( strings[value].decode("UTF-8"), self )
          elif celltype == 5:
              seconds= int(value)
              epoch= datetime.datetime(2001, 1, 1)
              delta= datetime.timedelta( seconds=seconds )
              theDate= epoch + delta
              v= stingray.cell.DateCell( theDate, self )
          elif celltype == 6:
              v= stingray.cell.BooleanCell( value, self )
          elif celltype == 7: # Actually a duration
              v= stingray.cell.NumberCell( value, self )
          elif celltype in (4, 8):
              ast_0= formulae[value][0]
              v= stingray.cell.ErrorCell( "Formula {0}".format(ast_0[1]), self )
          else:
              raise Exception( "Unknown cell type {0!r}".format(cell) )
          self.log.debug( "_cell {0} = {1}".format( cell, v) ) 
          return v

..  py:method:: Numbers13_Workbook.rows_of( sheet )

    Iterator through all rows.

::

      def rows_of( self, sheet ):
          """Iterator over rows.

          :param sheet: a Sheet object to retrieve rows from.
          """
          self.log.debug( "rows of {0}: {1}".format(sheet, sheet.name) )
          document = self.archive[1]
          for sheet_ref in document.sheets:
              sheet_msg= self.archive[sheet_ref[1][0]]
              for table_ref in sheet_msg.drawable_infos:
                  tableinfo= self.archive[table_ref[1][0]]
                  tablemodel= self.archive[tableinfo.tableModel[1][0]]
                  if (sheet_msg.name, tablemodel.table_name) != sheet.name:
                      continue
                  for row in self._row_iter( tablemodel ):
                      yield row
                  break
                    
::

      def _row_iter( self, tablemodel ):
          str_values= dict()
          for str_ref in tablemodel.data_store.stringTable:
              strs= self.archive[str_ref[1][0]]
              for entry in strs.entries:
                  str_values[entry.key]= entry.string
          form_values= dict()
          for form_ref in tablemodel.data_store.formulaTable:
              form= self.archive[form_ref[1][0]]
              for entry in form.entries:
                  form_values[entry.key]= entry.formula.AST_node_array.AST_node
          error_values= dict()
          for error_ref in tablemodel.data_store.formulaErrorTable:
              error= self.archive[error_ref[1][0]]
              # TODO:  decode error details
              #for entry in error.entries:
              #    error_values[entry.key]= entry.formula # entry.string?
          for t in tablemodel.data_store.tiles:
              for tt_ref in t.tiles:
                  tile= self.archive[tt_ref.ref[1][0]]
                  for row in tile.rowInfos:
                      yield [ self._cell(row.cells[col], str_values, form_values, error_values) 
                          for col in row.cells ] 
        
