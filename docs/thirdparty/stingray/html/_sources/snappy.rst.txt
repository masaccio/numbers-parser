..  #!/usr/bin/env python3

.. _`snappy`:

###############################################################
Snappy Module -- Unpacking iWork 13 files.
###############################################################

This is not a full implementation of the Snappy compression protocol.
It's a minimal implementation, enough to unpack iWork '13 files.

..  py:module:: snappy 

The iWork '13 use of Snappy
===============================================

https://github.com/obriensp/iWorkFileFormat

https://github.com/obriensp/iWorkFileFormat/blob/master/Docs/index.md

    "Components are serialized into .iwa (iWork Archive) files, 
    a custom format consisting of a Protobuf stream wrapped in a Snappy stream.

    "Snappy Compression

    "Snappy is a compression format created by Google aimed at providing decent 
    compression ratios at high speeds. IWA files are stored in Snappy's framing format,
    though they do not adhere rigorously to the spec. 
    In particular, they do not include the required Stream Identifier chunk, 
    and compressed chunks do not include a CRC-32C checksum.

    "The stream is composed of contiguous chunks prefixed by a 4 byte header. 
    The first byte indicates the chunk type, which in practice is always 0 for iWork, 
    indicating a Snappy compressed chunk. 
    The next three bytes are interpreted as a 24-bit little-endian integer 
    indicating the length of the chunk. 
    The 4 byte header is not included in the chunk length.

Snappy
===============================================

Here's some more information on the snappy protocol.

http://en.wikipedia.org/wiki/Snappy_(software)

https://code.google.com/p/snappy/

https://code.google.com/p/snappy/source/browse/trunk/format_description.txt

Implementation
===============

Module docstring.

::

  """Read snappy-compressed IWA files used for Numbers '13 workbooks.

  This is a variation on the "official" snappy protocol. The CRC checksums
  are not used by iWork '13. This is not a full implementation, just
  a decoder for iWork snappy-compressed IWA files.
    
  See https://code.google.com/p/snappy/
  """

Some overheads

::

  import logging
  import sys

..  py:function:: bytes_int( seq )

    Decode a sequence of bytes into an integer.
   
    :param seq: sequence of bytes; the entire sequence is consumed.
   
    :returns: integer

::

  def bytes_int( seq ):
      """8-bit encoded integer as sequence of 1 to 4 bytes, little-endian.
      """
      shift= 0 
      v= 0
      for b in seq:
          v += b<<shift 
          shift += 8
      return v

..  py:function:: varint( stream )

    Decode varint-encoded sequence of bytes.
   
    :param seq: sequence of bytes; consume bytes to decode the int.
   
    :returns: integer


::

  def varint( stream ):
      """7-bit encoded integer as a sequence of bytes, little-endian.
      MSB is used to indicate if more bytes are part of this value.

      >>> varint( iter([0xfe, 0xff, 0x7f]) )
      2097150
      """
      b= next(stream)
      shift= 0
      v = (b & 0x7F) # <<shift to be pedantic
      while b & 0x80 != 0:
          b= next(stream)
          shift += 7
          v += (b & 0x7F)<<shift
      return v

The snappy protocol has two levels. 

-   The LZ77 decoder which expands the tags to create the data.
    There are four kinds of tags.
   
    -   0b00: literal
   
                    Literals are uncompressed data stored directly in the byte stream.
                    The literal length is stored differently depending on the length
                    of the literal:

                    -  For literals up to and including 60 bytes in length, the upper
                       six bits of the tag byte contain (len-1). The literal follows
                       immediately thereafter in the bytestream.
              
                    -  For longer literals, the (len-1) value is stored after the tag byte,
                       little-endian. The upper six bits of the tag byte describe how
                       many bytes are used for the length; 60, 61, 62 or 63 for
                       1-4 bytes, respectively. The literal itself follows after the
                       length.
                      
    -   0b01: Copy with 1-byte offset

                        These elements can encode lengths between [4..11] bytes and offsets
                        between [0..2047] bytes. (len-4) occupies three bits and is stored
                        in bits [2..4] of the tag byte. The offset occupies 11 bits, of which the
                        upper three are stored in the upper three bits ([5..7]) of the tag byte,
                        and the lower eight are stored in a byte following the tag byte.
                       
    -   0b10: Copy with a 2-byte offset
   
                        These elements can encode lengths between [1..64] and offsets from
                        [0..65535]. (len-1) occupies six bits and is stored in the upper
                        six bits ([2..7]) of the tag byte. The offset is stored as a
                        little-endian 16-bit integer in the two bytes following the tag byte.

    -   0b11: Copy with a 4-byte offset
   
                        These are like the copies with 2-byte offsets (see previous subsection),
                        except that the offset is stored as a 32-bit integer instead of a
                        16-bit integer (and thus will occupy four bytes).


-   The higher-level framing protocol. 
   
    -   type "0" (Compressed Data) frame with a  three-byte length.
   
    -   Other types are possible in principle. Numbers '13 doesn't use them.

..  py:class:: Snappy

Implement the two-level snappy protocol used by Numbers '13.

-   The LZ77 decoder which expands the tags to create the data.

-   The higher-level framing protocol
    with just one kind of frame, type "0" (Compressed Data) with a 
    three-byte length.


::

  class Snappy:
      def __init__( self ):
          self.log= logging.getLogger( self.__class__.__qualname__ )

..  py:method:: Snappy.lz77( frame )

    The LZ77 decoder. This locates the **varint** size header.  That's followed by 
    a sequence of tags.  The literal tag has data. The other three tags repeat
    previously output bytes.
   
    We're building a ``bytearray`` buffer from the input. This means copying
    literals into the buffer. It also means copying part of the buffer into the buffer
    to add one of the three kinds of copies.

    Because of the framing protocol, we're limited to a buffer of only 64K bytes.
       
::

      def lz77( self, frame ):
          """Decode one frame of a Snappy LZ77-encoded stream.

          Get the tags, data and emit the resulting uncompressed bytes for this frame.

          There are four types of tags:

          0b00 - Literal - the balance of the tag specifies the length.
          0b01 - Copy 1-byte offset - repeat previous bytes from the output buffer.
          0b10 - Copy 2-byte offset - repeat previous bytes
          0b11 - Copy 4-byte offset - repeat previous bytes

          :param frame: One frame from a Snappy file.
          :returns: buffer of bytes for this frame.
          """
          buffer= bytearray()
          # DEBUG: print(f"FRAME: {frame!r}")
          stream= iter( frame )
          # The size of the uncompressed data in this frame.
          size= varint( stream ) 
          self.log.debug( "  LZ77 size {0}".format(size) )
          # Build the uncompressed buffer.
          while len(buffer) < size:
              hdr= int(next(stream))
              tag_upper, element_type = hdr >> 2, hdr & 0b11
    
              if element_type == 0b00: # Literal
                  if tag_upper < 60:
                      size_elt= tag_upper
                  else:
                      size_elt= bytes_int( next(stream) for i in range(tag_upper - 59) )
                  bytes= [next(stream) for b in range(size_elt+1)]
                  self.log.debug( 
                      "{0:08b} {1} {2} = {3!r}".format(
                      hdr, element_type, size_elt, bytes) )
                  buffer.extend( bytes )
              else: # Some kind of copy
                  # Copy -- gather bytes based on offset, stow into buffer based on length
                  if element_type == 0b01: # Copy with 1-byte offset
                      length, offset_hi = tag_upper & 0b111, (tag_upper & 0b111000)>>3
                      offset_lo= next(stream)
                      offset= (offset_hi<<8)+offset_lo
                      self.log.debug( 
                          "{0:08b} {1:8b} {2} {3} = {4!r} {5!r}".format( 
                          hdr, offset_lo, element_type, length, (offset_hi, offset_lo), offset) )
                      length += 4
                  elif element_type == 0b10: # Copy with 2-byte offset
                      offset= bytes_int( next(stream) for i in range(2) )
                      length= tag_upper
                      self.log.debug( "{0:08b} {1} {2} {3}".format(hdr, element_type, length, offset) )
                      length += 1
                  elif element_type == 0b11: # Copy with 4-byte offset
                      offset= bytes_int( next(stream) for i in range(4) )
                      length= tag_upper
                      self.log.debug( 
                          "{0:08b} {1} {2} {3}".format(
                          hdr, element_type, length, offset) )
                      length += 1
                  else:
                      raise Exception( "Logic Problem" )
                  # Extend buffer with the copied bytes. 
                  # Handle RLE feature, if necessary.
                  copy= buffer[-offset:]
                  if offset < length:
                      repeat= copy[:]
                      while len(copy) < length:
                          copy += repeat
                  buffer.extend( copy[:length] )
          assert len(buffer) == size, "len(buffer) {0} != size {1}".format(len(buffer),size)
          return buffer

..  py:method:: Snappy.decompress( file_object )

    The Framing protocol required to decode. The frames contain up to 64K of compressed
    data. This defines a sequence of windows over the stream of data.
       
::

      def decompress( self, file_object ):
          """Decompress a snappy file object. Locate each frame in the snappy 
          framing protocol that's used by iWork (not precisely as specified
          by Google.) For each frame, do the LZ77 expansion on the frame's bytes to
          build the uncompressed data.

          Frames have a 4-byte header. Byte 0 is frame type, only type 0 (Compressed Data)
          is supported. Bytes 1-3 are a 24-bit size for the frame. 
          Practically, it's limited to 65536 bytes.
        
          The CRC32 is omitted for iWork files

          ..  todo:: yield iterable byte stream for use in higher-levels of the protocol.

              It's not *required* to materialize the entire data buffer as a single object.
              The intent of the framing is to limit the size of the buffer required.
    
          Note that we could provide ``file_object`` file directly to ``lz77()`` function because
          lz77 protocol starts with the target uncompressed size at the front of the frame.
          We don't **actually** need to read the frame here.
          """
          data= bytearray()
          header= file_object.read(4)
          while header:
              # The Snappy framing format: type 0 (Compressed Data) with a 24-bit size.
              # The CRC32 is omitted for iWork files
              type_frame, size_frame = header[0], bytes_int(header[1:4])
              assert type_frame == 0, "Unsupported Snappy Frame {0}".format(type_frame)
              self.log.debug( "Frame type {0} size {1}".format( type_frame, size_frame ) )
              frame= file_object.read( size_frame )
              data.extend( self.lz77( frame ) )
              header= file_object.read(4)
          return data
