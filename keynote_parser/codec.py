from __future__ import print_function
from __future__ import absolute_import
from builtins import zip
from builtins import str
from builtins import object
import sys
import yaml
import struct
import snappy
import traceback

from .mapping import NAME_CLASS_MAP, ID_NAME_MAP

from google.protobuf.internal.encoder import _VarintBytes
from google.protobuf.internal.decoder import _DecodeVarint32
from google.protobuf.json_format import MessageToDict
from .protobuf_patch import ParseDict

from .generated.TSPArchiveMessages_pb2 import ArchiveInfo


class IWAFile(object):
    def __init__(self, chunks):
        self.chunks = chunks

    @classmethod
    def from_file(cls, filename):
        with open(filename) as f:
            return cls.from_buffer(f.read())

    @classmethod
    def from_buffer(cls, data):
        chunks = []
        while data:
            chunk, data = IWACompressedChunk.from_buffer(data)
            chunks.append(chunk)

        return cls(chunks)

    @classmethod
    def from_dict(cls, _dict):
        return cls([
            IWACompressedChunk.from_dict(chunk)
            for chunk in _dict['chunks']
        ])

    def to_dict(self):
        return {
            "chunks": [chunk.to_dict() for chunk in self.chunks]
        }

    def to_buffer(self):
        return b''.join([
            chunk.to_buffer() for chunk in self.chunks
        ])


class IWACompressedChunk(object):
    def __init__(self, archives):
        self.archives = archives

    @classmethod
    def _decompress_all(cls, data):
        while data:
            header = data[:4]

            first_byte = header[0]
            if not isinstance(first_byte, int):
                first_byte = ord(first_byte)

            if first_byte != 0x00:
                raise ValueError(
                    "IWA chunk does not start with 0x00! (found %x)" %
                    first_byte)

            unpacked = struct.unpack_from('<I', bytes(header[1:]) + b'\x00')
            length = unpacked[0]
            chunk = data[4:4 + length]
            data = data[4 + length:]

            try:
                yield snappy.uncompress(chunk)
            except Exception:
                # Try to see if this data isn't compressed in the first place.
                # If this data is still compressed, parsing it as Protobuf
                # will almost definitely fail anyways.
                yield chunk

    @classmethod
    def from_buffer(cls, data):
        data = b''.join(cls._decompress_all(data))
        archives = []
        while data:
            archive, data = IWAArchiveSegment.from_buffer(data)
            archives.append(archive)
        return cls(archives), None

    @classmethod
    def from_dict(cls, _dict):
        return cls([IWAArchiveSegment.from_dict(d) for d in _dict['archives']])

    def to_dict(self):
        return {
            "archives": [
                archive.to_dict() for archive in self.archives
            ]
        }

    def to_buffer(self):
        payload = snappy.compress(b''.join([
            archive.to_buffer() for archive in self.archives
        ]))
        return b'\x00' + struct.pack('<I', len(payload))[:3] + payload


class IWAArchiveSegment(object):
    def __init__(self, header, objects):
        self.header = header
        self.objects = objects

    def __eq__(self, other):
        return self.header == other.header and self.objects == other.objects

    @classmethod
    def from_buffer(cls, buf):
        archive_info, payload = get_archive_info_and_remainder(buf)
        if not repr(archive_info):
            raise ValueError(
                "Segment doesn't seem to start with an ArchiveInfo!")

        payloads = []

        n = 0
        for message_info in archive_info.message_infos:
            try:
                klass = ID_NAME_MAP[message_info.type]
            except KeyError:
                raise NotImplementedError(
                    "Don't know how to parse Protobuf message type " +
                    str(message_info.type))
            try:
                output = klass.FromString(payload[n:n + message_info.length])
            except Exception as e:
                raise ValueError(
                    "Failed to deserialize %s payload of length %d: %s" % (
                        klass, message_info.length, e))
            payloads.append(output)
            n += message_info.length

        return cls(archive_info, payloads), payload[n:]

    @classmethod
    def from_dict(cls, _dict):
        return cls(
            dict_to_header(_dict['header']),
            [dict_to_message(o) for o in _dict['objects']])

    def to_dict(self):
        return {
            "header": header_to_dict(self.header),
            "objects": [message_to_dict(message) for message in self.objects],
        }

    def to_buffer(self):
        # Each message_info as part of the header needs to be updated
        # so that its length matches the object contained within.
        for obj, message_info in zip(self.objects, self.header.message_infos):
            object_length = len(obj.SerializeToString())
            provided_length = message_info.length
            if object_length != provided_length:
                message_info.length = object_length
        return b''.join([
            _VarintBytes(self.header.ByteSize()),
            self.header.SerializeToString()
        ] + [
            obj.SerializeToString() for obj in self.objects
        ])


def message_to_dict(message):
    output = MessageToDict(message)
    output['_pbtype'] = type(message).DESCRIPTOR.full_name
    return output


def header_to_dict(message):
    output = message_to_dict(message)
    for message_info in output['messageInfos']:
        del message_info['length']
    return output


def dict_to_message(_dict):
    _type = _dict['_pbtype']
    del _dict['_pbtype']
    return ParseDict(_dict, NAME_CLASS_MAP[_type]())


def dict_to_header(_dict):
    for message_info in _dict['messageInfos']:
        # set a dummy length value that we'll overwrite later
        message_info['length'] = 0
    return dict_to_message(_dict)


def get_archive_info_and_remainder(buf):
    msg_len, new_pos = _DecodeVarint32(buf, 0)
    n = new_pos
    msg_buf = buf[n:n + msg_len]
    n += msg_len
    return ArchiveInfo.FromString(msg_buf), buf[n:]


if __name__ == "__main__":
    for filename in sys.argv[1:]:
        try:
            iwa_file = IWAFile.from_file(filename)
            print(yaml.safe_dump(
                iwa_file.to_dict(),
                default_flow_style=False))
        except Exception as e:
            print("FAILED", traceback.format_exc(e))
