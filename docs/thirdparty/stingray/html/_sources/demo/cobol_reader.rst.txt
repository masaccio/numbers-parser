..    #!/usr/bin/env python3

..  _`demo_cobol`:

##########################################################
Reading COBOL Files
##########################################################

For sample data, we're using data found here:
http://wonder.cdc.gov/wonder/sci_data/codes/fips/type_txt/cntyxref.asp

The data files are in two zip archives: http://wonder.cdc.gov/wonder/sci_data/datasets/zipctyA.zip 
and http://wonder.cdc.gov/wonder/sci_data/datasets/zipctyB.zip

Each of these archives contains five large files, with 2,310,000 rows of data, plus a header. 
The 10th file has 2,037,944 rows of data plus a header.

The member names of the ZIP archive are zipcty1 to zipcty5
and zipcty6 to zipcty10.

We'll work with two small subsets in the sample directory.

Here are the two record layouts.

..    parsed-literal::

    COUNTY CROSS-REFERENCE FILE - COBOL EXAMPLE


            BLOCK CONTAINS 0 RECORDS
            LABEL RECORDS ARE STANDARD
            RECORD CONTAINS 53 CHARACTERS
            RECORDING MODE IS F
            DATA RECORDS ARE
                   COUNTY-CROSS-REFERENCE-RECORD.

        01  COUNTY-CROSS-REFERENCE-RECORD.
            05   ZIP-CODE                                 PIC X(05).
            05   UPDATE-KEY-NO                            PIC X(10).
            05   ZIP-ADD-ON-RANGE.
                 10  ZIP-ADD-ON-LOW-NO.
                      15  ZIP-SECTOR-NO                   PIC X(02).
                      15  ZIP-SEGMENT-NO                  PIC X(02).
                 10  ZIP-ADD-ON-HIGH-NO.
                      15  ZIP-SECTOR-NO                   PIC X(02).
                      15  ZIP-SEGMENT-NO                  PIC X(02).
            05   STATE-ABBREV                             PIC X(02).
            05   COUNTY-NO                                PIC X(03).
            05   COUNTY-NAME                              PIC X(25).

..    parsed-literal::


    COPYRIGHT HEADER RECORD - COBOL EXAMPLE


               BLOCK CONTAINS 0 RECORDS
               LABEL RECORDS ARE STANDARD
               RECORD CONTAINS 53 CHARACTERS
               RECORDING MODE IS F
               DATA RECORDS ARE
                   COPYRIGHT-HEADER RECORD.

          01  COPYRIGHT-HEADER-RECORD.
              05  FILLER                                     PIC  X(05).
              05  FILE-VERSION-YEAR                          PIC  X(02).
              05  FILE-VERSION-MONTH                         PIC  X(02).
              05  COPYRIGHT-SYMBOL                           PIC  X(11).
              05  TAPE-SEQUENCE-NO                           PIC  X(03).
              05  FILLER                                     PIC  X(30).

Implementation
==============

The actual COBOL code is in :file:`demo/zipcty.cob`. This file has both record
layouts. Two ``01`` level items with a redefines.

Here are the imports we'll use. We'll need the COBOL loader to read
the source  :file:`demo/zipcty.cob` schema and create a COBOL file we
can process. 

We'll use ``types.SimpleNamespace`` to build objects from the source data.

We might use ``pprint`` to dump values.

::

    import stingray.cobol.loader
    import stingray.cobol
    import pprint
    import types
    
When working with unknown files, we sometimes need to preview a raw dump of the
records.

::
    
    def raw_dump( schema, sheet ):
        for row in sheet.rows():
            stingray.cobol.dump( schema, row )
        
This is a handy expedient for debugging.

As suggested in :ref:`developer`, here are two builder functions.
The ``header_builder()`` function creates a header object from the 
first row of each :file:`zipcty*` file. 

::

    def header_builder(row, schema):
        return types.SimpleNamespace( 
            file_version_year= row.cell(schema['FILE-VERSION-YEAR']).to_str(),
            file_version_month= row.cell(schema['FILE-VERSION-MONTH']).to_str(),
            copyright_symbol= row.cell(schema['COPYRIGHT-SYMBOL']).to_str(),
            tape_sequence_no= row.cell(schema['TAPE-SEQUENCE-NO']).to_str(),
        )
  
The ``detail_builder()`` function creates a detail object from the subsequent
rows of each :file:`zipcty*` file. 

Because the names within the COBOL layout are not unique at the bottom-most
element level, we must use path names. The default path names include all
levels of the DDE. More clever path name components might be useful here.

COBOL uses an "of" to work up the hierarchy looking for a unique name.

Maybe we could build a fluent interface ``schema['ZIP-SECTOR-NO'].of('ZIP-ADD-ON-LOW-NO')``.
  
::

    def detail_builder(row, schema):
        return types.SimpleNamespace( 
            zip_code= row.cell(schema['ZIP-CODE']).to_str(),
            update_key_no= row.cell(schema['UPDATE-KEY-NO']).to_str(),
            low_sector= row.cell(schema['COUNTY-CROSS-REFERENCE-RECORD.ZIP-ADD-ON-RANGE.ZIP-ADD-ON-LOW-NO.ZIP-SECTOR-NO']).to_str(),
            low_segment= row.cell(schema['COUNTY-CROSS-REFERENCE-RECORD.ZIP-ADD-ON-RANGE.ZIP-ADD-ON-LOW-NO.ZIP-SEGMENT-NO']).to_str(),
            high_sector= row.cell(schema['COUNTY-CROSS-REFERENCE-RECORD.ZIP-ADD-ON-RANGE.ZIP-ADD-ON-HIGH-NO.ZIP-SECTOR-NO']).to_str(),
            high_segment= row.cell(schema['COUNTY-CROSS-REFERENCE-RECORD.ZIP-ADD-ON-RANGE.ZIP-ADD-ON-HIGH-NO.ZIP-SEGMENT-NO']).to_str(),
            state_abbrev= row.cell(schema['STATE-ABBREV']).to_str(),
            county_no= row.cell(schema['COUNTY-NO']).to_str(),
            county_name= row.cell(schema['COUNTY-NAME']).to_str(),
        )

Here's the ``process_sheet()`` function which applies the builders to the various
rows in each sheet. Currently, all that happens is a print of the object that
was built.

Note that we've transformed the schema from a simple, flat list into
a dictionary keyed by field name. For COBOL processing, this is essential, since
the numeric order of fields isn't often sensible.

Also note that we've put two versions of each name into the schema dictionary.

-   The lowest level name.

-   The entire path from ``01`` level down to the lowest level name.

[For spreadsheets, where columns are numbered, the positional information may be
useful.]

::

    def process_sheet( sheet ):
        schema_dict= dict( (a.name, a) for a in sheet.schema )
        schema_dict.update( dict( (a.path, a) for a in sheet.schema ) )
        
        counts= { 'read': 0 }
        row_iter= sheet.rows()
        row= next(row_iter)
        header= header_builder( row, schema_dict )
        print( header )
        
        for row in row_iter:
            data= detail_builder( row, schema_dict )
            print( data )
            counts['read'] += 1
        return counts

The top-level script must do two things:

1.  Parse the :file:`"zipcty.cob"` data definition to create a schema.

2.  Open a data file as a :py:class:`cobol.Character_File`. This presumes
    the file is all character (no COMP-3) and already translated into ASCII.
    
    The :py:func:`process_sheet` is applied to each file.

Here's the script.

::

    with open("sample/zipcty.cob", "r") as cobol:
        schema= stingray.cobol.loader.COBOLSchemaLoader( cobol ).load()
        #pprint.pprint( schema )
    for filename in 'sample/zipcty1', 'sample/zipcty2':
        with stingray.cobol.Character_File( filename, schema=schema ) as wb:
            sheet= wb.sheet( filename )
            #counts= process_sheet( sheet )
            #pprint.pprint( counts )
            raw_dump( schema, sheet )
            
Running the demo
================

We can run this program like this:

..  code-block:: bash

    python3 demo/cobol_reader.py
    
The output looks like this.

..  parsed-literal::

    namespace(copyright_symbol=' (C)USPS', file_version_month='09', file_version_year='88', tape_sequence_no='001')
    namespace(county_name='WESTCHESTER', county_no='119', high_sector='00', high_segment='01', low_sector='00', low_segment='01', state_abbrev='NY', update_key_no='0000000001', zip_code='00401')
    namespace(county_name='NORFOLK', county_no='021', high_sector='52', high_segment='66', low_sector='52', low_segment='66', state_abbrev='MA', update_key_no='0000462001', zip_code='02186')
    namespace(county_name='HARTFORD', county_no='003', high_sector='49', high_segment='01', low_sector='49', low_segment='01', state_abbrev='CT', update_key_no='0000924001', zip_code='06111')
    namespace(county_name='UNION', county_no='039', high_sector='22', high_segment='08', low_sector='22', low_segment='08', state_abbrev='NJ', update_key_no='0001386001', zip_code='07901')
    namespace(county_name='BRONX', county_no='005', high_sector='17', high_segment='05', low_sector='17', low_segment='05', state_abbrev='NY', update_key_no='0001848001', zip_code='10463')
    {'read': 5}
    namespace(copyright_symbol=' (C)USPS', file_version_month='09', file_version_year='88', tape_sequence_no='002')
    namespace(county_name='SUFFOLK', county_no='103', high_sector='25', high_segment='43', low_sector='25', low_segment='43', state_abbrev='NY', update_key_no='0002310001', zip_code='11789')
    namespace(county_name='CHAUTAUQUA', county_no='013', high_sector='97', high_segment='71', low_sector='97', low_segment='71', state_abbrev='NY', update_key_no='0002772001', zip_code='14767')
    namespace(county_name='FRANKLIN', county_no='055', high_sector='90', high_segment='33', low_sector='90', low_segment='33', state_abbrev='PA', update_key_no='0003234001', zip_code='17201')
    namespace(county_name='MONTGOMERY', county_no='091', high_sector='28', high_segment='22', low_sector='28', low_segment='22', state_abbrev='PA', update_key_no='0003696001', zip_code='19438')
    namespace(county_name='WASHINGTON', county_no='043', high_sector='53', high_segment='05', low_sector='53', low_segment='05', state_abbrev='MD', update_key_no='0004158001', zip_code='21740')
    {'read': 5}

Working with Archives
=====================

We don't need to unpack the archives to work with them.
We can open a ZipFile member and process that. 
This can be a helpful optimization when small extracts are pulled from ZIP archives.

The trick is this.

When we open the file with ``stingray.cobol.Character_File( filename, schema=schema )``
we can pass the file object created by ``ZipFile.open()`` as the second argument.

``stingray.cobol.Character_File( filename, file_object=archive.open(filename), schema=schema )``



