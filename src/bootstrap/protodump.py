"""
Super hacky script to parse compiled Protobuf definitions out of
one or more binary files in a directory tree.

Requires `pip install 'protobuf>=3.20.0rc1'`.
Example usage:
 python3 protodump.py /Applications/SomeAppBundle.app ./proto_files_go_here/

(c) Peter Sobot (@psobot), March 13, 2022
Inspired by Sean Patrick O'Brien (@obriensp)'s 2013 "proto-dump":
https://github.com/obriensp/proto-dump
"""

from pathlib import Path

from google.protobuf import descriptor_pb2
from google.protobuf.descriptor_pool import DescriptorPool
from google.protobuf.internal import api_implementation
from google.protobuf.internal.decoder import SkipField, _DecodeVarint
from google.protobuf.message import DecodeError
from termcolor import cprint


class ProtoFile(object):
    def __init__(self, data, pool):
        self.data = data
        self.pool = pool
        self.file_descriptor_proto = descriptor_pb2.FileDescriptorProto.FromString(data)
        self.path = self.file_descriptor_proto.name
        self.imports = list(self.file_descriptor_proto.dependency)
        self.attempt_to_load()

    def __hash__(self):
        return hash(self.data)

    def __eq__(self, other):
        return isinstance(other, ProtoFile) and self.data == other.data

    def attempt_to_load(self):
        # This method will fail if this file is missing dependencies (imports)
        try:
            return self.pool.Add(self.file_descriptor_proto)
        except Exception:
            return None

    @property
    def descriptor(self):
        return self.attempt_to_load()

    def __repr__(self):
        return '<%s: path="%s">' % (self.__class__.__name__, self.path)

    @property
    def source(self):
        if self.descriptor:
            return self.descriptor.GetDebugString()
        return None


def read_until_null_tag(data):
    position = 0
    while position < len(data):
        try:
            tag, position = _DecodeVarint(data, position)
        except Exception:
            return position

        if tag == 0:
            # Found a null tag, so we're done
            return position

        try:
            new_position = SkipField(data, position, len(data), bytes([tag]))
        except (AttributeError, DecodeError):
            return position
        if new_position == -1:
            return position
        position = new_position


def extract_proto_from_file(filename, descriptor_pool):
    with open(filename, "rb") as f:
        data = f.read()
    offset = 0

    proto_marker = b".proto"

    while True:
        # Look for ".proto"
        suffix_position = data.find(proto_marker, offset)
        if suffix_position == -1:
            break

        marker_start = data.rfind(b"\x0A", offset, suffix_position)
        if marker_start == -1:
            # Doesn't look like a proto descriptor
            offset = suffix_position + len(proto_marker)
            continue

        try:
            name_length, new_pos = _DecodeVarint(data, marker_start)
        except Exception:
            # Expected a VarInt here, so if not, continue
            offset = suffix_position + len(proto_marker)
            continue

        # Length = 1 byte for the marker (0x0A)
        # + length of the varint + length of the descriptor name
        expected_length = 1 + (new_pos - marker_start) + name_length + 7
        current_length = (suffix_position + len(proto_marker)) - marker_start

        # Huge margin of error here - my calculations above are probably just wrong.
        if current_length > expected_length + 30:
            offset = suffix_position + len(proto_marker)
            continue

        # Split the data starting at the marker byte and try to read it as a
        # protobuf stream. Descriptors are stored as c strings in the .pb.cc files.
        # They're null-terminated, but can also contain embedded null bytes. Since we
        # can't search for the null-terminator explicitly, we parse the string manually
        # until we reach a protobuf tag which equals 0 (identifier = 0, wiretype =
        # varint), signalling the final null byte of the string. This works because
        # there are no 0 tags in a real FileDescriptorProto stream.
        descriptor_length = read_until_null_tag(data[marker_start:]) - 1
        descriptor_data = data[marker_start : marker_start + descriptor_length]
        try:
            proto_file = ProtoFile(descriptor_data, descriptor_pool)
            if (
                proto_file.path.endswith(".proto")
                and proto_file.path != "google/protobuf/descriptor.proto"
            ):
                yield proto_file
        except Exception:
            pass

        offset = marker_start + descriptor_length


def find_missing_dependencies(all_files, source_file):
    matches = [f for f in all_files if f.path == source_file]
    if not matches:
        return {source_file}

    missing = set()
    for match in matches:
        if not match.attempt_to_load():
            missing.update(set(match.imports))

    to_return = set()
    for dep in missing:
        to_return.update(find_missing_dependencies(all_files, dep))

    return to_return


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description=(
            "Read all files in a given directory and scan each file for protobuf definitions,"
            " printing usable .proto files to a given directory."
        )
    )
    parser.add_argument("input_path", help="Input path to scan. May be a file or directory.")
    parser.add_argument("output_path", help="Output directory to dump .protoc files to.")

    args = parser.parse_args()

    if api_implementation.Type() != "cpp":
        raise NotImplementedError(
            "This script requires the Protobuf installation to use the C++ implementation. Please"
            " reinstall Protobuf with C++ support."
        )

    global_descriptor_pool = DescriptorPool()

    all_filenames = [str(path) for path in Path(args.input_path).rglob("*") if not path.is_dir()]

    cprint(
        f"Bootstrap: scanning {len(all_filenames):,}"
        + f" files under {args.input_path} for protobuf definitions",
        "green",
    )

    proto_files_found = set()
    for path in all_filenames:
        for proto in extract_proto_from_file(path, global_descriptor_pool):
            proto_files_found.add(proto)

    missing_deps = set()
    for found in proto_files_found:
        if not found.attempt_to_load():
            missing_deps.update(find_missing_dependencies(proto_files_found, found.path))

    if missing_deps:
        cprint(
            "Warning: unable to print out all Protobuf definitions;"
            + f" {len(missing_deps):,} proto files could not be found:\n{missing_deps}",
            "red",
        )
    else:
        for proto_file in proto_files_found:
            Path(args.output_path).mkdir(parents=True, exist_ok=True)
            with open(Path(args.output_path) / proto_file.path, "w") as f:
                source = proto_file.source
                if source:
                    f.write(source)
                else:
                    cprint(f"Warning: no source available for {proto_file}", "red")
        cprint(
            f"Bootstrap: wrote {len(proto_files_found):,} proto files to {args.output_path}",
            "green",
        )


if __name__ == "__main__":
    main()
