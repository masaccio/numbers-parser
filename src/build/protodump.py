"""Super hacky script to parse compiled Protobuf definitions out of one or more binary files in a directory tree.

Requires `pip install 'protobuf>=3.20.0rc1'`.
Example usage:
 python3 protodump.py /Applications/SomeAppBundle.app ./proto_files_go_here/

(c) Peter Sobot (@psobot), March 13, 2022
Inspired by Sean Patrick O'Brien (@obriensp)'s 2013 "proto-dump": https://github.com/obriensp/proto-dump
"""

import sys
from collections import defaultdict
from pathlib import Path

from google.protobuf import descriptor_pb2
from google.protobuf.descriptor_pool import DescriptorPool
from google.protobuf.internal.decoder import SkipField, _DecodeVarint
from google.protobuf.message import DecodeError
from tqdm import tqdm

PROTO_TYPES = {
    1: "double",
    2: "float",
    3: "int64",
    4: "uint64",
    5: "int32",
    6: "fixed64",
    7: "fixed32",
    8: "bool",
    9: "string",
    12: "bytes",
    13: "uint32",
    15: "sfixed32",
    16: "sfixed64",
    17: "sint32",
    18: "sint64",
}


def to_proto_file(fds: descriptor_pb2.FileDescriptorSet) -> str:
    if len(fds.file) != 1:
        msg = "Only one file per fds."
        raise NotImplementedError(msg)
    f = fds.file[0]
    lines = [
        'syntax = "proto2";',
        "",
    ]

    for dependency in f.dependency:
        lines.append(f'import "{dependency}";')

    lines.append(f"package {f.package};")
    lines.append("")

    def generate_enum_lines(f, lines: list[str], indent: int = 0):
        prefix = "  " * indent
        for enum in f.enum_type:
            lines.append(prefix + f"enum {enum.name} " + "{")
            for value in enum.value:
                lines.append(prefix + f"  {value.name} = {value.number};")
            lines.append(prefix + "}")

    def generate_field_line(field, in_oneof: bool = False) -> str:
        line = []
        if field.label == 1:
            if not in_oneof:
                line.append("optional")
        elif field.label == 2:
            line.append("required")
        elif field.label == 3:
            line.append("repeated")
        else:
            msg = "Unknown field label type!"
            raise NotImplementedError(msg)

        if field.type in PROTO_TYPES:
            line.append(PROTO_TYPES[field.type])
        elif field.type in (11, 14):  # MESSAGE
            line.append(field.type_name)
        else:
            msg = f"Unknown field type {field.type}!"
            raise NotImplementedError(msg)

        line.append(field.name)
        line.append("=")
        line.append(str(field.number))
        options = []
        if field.default_value:
            options.append(f"default = {field.default_value}")
        if field.options.deprecated:
            options.append("deprecated = true")
        if field.options.packed:
            options.append("packed = true")
        # TODO: Protobuf supports other options in square brackets!
        # Add support for them here to make this feature-complete.
        if options:
            line.append(f"[{', '.join(options)}]")
        return f"  {' '.join(line)};"

    def generate_extension_lines(message, lines: list[str], indent: int = 0):
        prefix = "  " * indent
        extensions_grouped_by_extendee = defaultdict(list)
        for extension in message.extension:
            extensions_grouped_by_extendee[extension.extendee].append(extension)
        for extendee, extensions in extensions_grouped_by_extendee.items():
            lines.append(prefix + f"extend {extendee} {{")
            for extension in extensions:
                lines.append(prefix + generate_field_line(extension))
            lines.append(prefix + "}")

    def generate_message_lines(f, lines: list[str], indent: int = 0):
        prefix = "  " * indent

        submessages = f.message_type if hasattr(f, "message_type") else f.nested_type

        for message in submessages:
            # if message.name == "ContainedObjectsCommandArchive":
            #     breakpoint()
            lines.append(prefix + f"message {message.name} " + "{")

            generate_enum_lines(message, lines, indent + 1)
            generate_message_lines(message, lines, indent + 1)

            for field in message.field:
                if not field.HasField("oneof_index"):
                    lines.append(prefix + generate_field_line(field))

            # ...then the oneofs:
            next_prefix = "  " * (indent + 1)
            for oneof_index, oneof in enumerate(message.oneof_decl):
                lines.append(next_prefix + f"oneof {oneof.name} {{")
                for field in message.field:
                    if field.HasField("oneof_index") and field.oneof_index == oneof_index:
                        lines.append(next_prefix + generate_field_line(field, in_oneof=True))
                lines.append(next_prefix + "}")

            if len(message.extension_range):
                if len(message.extension_range) > 1:
                    msg = "Not sure how to handle multiple extension ranges!"
                    raise NotImplementedError(msg)
                start, end = (
                    message.extension_range[0].start,
                    min(message.extension_range[0].end, 536870911),
                )
                lines.append(next_prefix + f"extensions {start} to {end};")

            generate_extension_lines(message, lines, indent + 1)
            lines.append(prefix + "}")
            lines.append("")

    generate_enum_lines(f, lines)
    generate_message_lines(f, lines)
    generate_extension_lines(f, lines)

    return "\n".join(lines)


class ProtoFile:
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
        except Exception as e:
            if "duplicate file name" in str(e):
                return self.pool.FindFileByName(e.args[0].split("duplicate file name")[1].strip())
            return None

    @property
    def descriptor(self):
        return self.attempt_to_load()

    def __repr__(self):
        return f'<{self.__class__.__name__}: path="{self.path}">'

    @property
    def source(self):
        if self.descriptor:
            fds = descriptor_pb2.FileDescriptorSet()
            fds.file.append(descriptor_pb2.FileDescriptorProto())
            fds.file[0].ParseFromString(self.descriptor.serialized_pb)
            return to_proto_file(fds)
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
    return None


def extract_proto_from_file(filename, descriptor_pool):
    with open(filename, "rb") as f:
        data = f.read()
    offset = 0

    PROTO_MARKER = b".proto"

    while True:
        # Look for ".proto"
        suffix_position = data.find(PROTO_MARKER, offset)
        if suffix_position == -1:
            break

        marker_start = data.rfind(b"\x0a", offset, suffix_position)
        if marker_start == -1:
            # Doesn't look like a proto descriptor
            offset = suffix_position + len(PROTO_MARKER)
            continue

        try:
            name_length, new_pos = _DecodeVarint(data, marker_start)
        except Exception:
            # Expected a VarInt here, so if not, continue
            offset = suffix_position + len(PROTO_MARKER)
            continue

        # Length = 1 byte for the marker (0x0A) + length of the varint + length of the descriptor name
        expected_length = 1 + (new_pos - marker_start) + name_length + 7
        current_length = (suffix_position + len(PROTO_MARKER)) - marker_start

        # Huge margin of error here - my calculations above are probably just wrong.
        if current_length > expected_length + 30:
            offset = suffix_position + len(PROTO_MARKER)
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
        ),
    )
    parser.add_argument("input_path", help="Input path to scan. May be a file or directory.")
    parser.add_argument("output_path", help="Output directory to dump .protoc files to.")

    args = parser.parse_args()
    GLOBAL_DESCRIPTOR_POOL = DescriptorPool()

    all_filenames = [str(path) for path in Path(args.input_path).rglob("*") if not path.is_dir()]

    print(
        f"Scanning {len(all_filenames):,} files under {args.input_path} for protobuf definitions...",
    )

    proto_files_found = set()
    for path in tqdm(all_filenames):
        for proto in extract_proto_from_file(path, GLOBAL_DESCRIPTOR_POOL):
            proto_files_found.add(proto)

    print(f"Found what look like {len(proto_files_found):,} protobuf definitions.")

    missing_deps = set()
    for found in proto_files_found:
        if not found.attempt_to_load():
            missing_deps.update(find_missing_dependencies(proto_files_found, found.path))

    for found in proto_files_found:
        if not found.attempt_to_load():
            missing_deps.add(found)

    if missing_deps:
        print(
            f"Unable to print out all Protobuf definitions; {len(missing_deps):,} proto files could"
            f" not be found:\n{missing_deps}",
        )
        sys.exit(1)
    else:
        for proto_file in tqdm(proto_files_found):
            Path(args.output_path).mkdir(parents=True, exist_ok=True)
            with open(Path(args.output_path) / proto_file.path, "w") as f:
                source = proto_file.source
                if source:
                    f.write(source)
                else:
                    print(f"Warning: no source available for {proto_file}")
        print(f"Done! Wrote {len(proto_files_found):,} proto files to {args.output_path}.")


if __name__ == "__main__":
    main()
