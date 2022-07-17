import os
import json
import regex
import sys

from array import array
from argparse import ArgumentParser
from base64 import b64decode
from binascii import hexlify
from compact_json import Formatter


from numbers_parser.file import read_numbers_file
from numbers_parser import _get_version
from numbers_parser.iwafile import IWAFile
from numbers_parser.exceptions import FileFormatError


def ensure_directory_exists(prefix, path):
    """Ensure that a path's directory exists."""
    parts = os.path.split(path)
    try:
        os.makedirs(os.path.join(*([prefix] + list(parts[:-1]))))
    except OSError:
        pass


def prettify_uuids(obj):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, dict) or isinstance(v, list):
                prettify_uuids(v)
            elif k == "lower" or k == "upper":
                obj[k] = "0x{0:0{1}X}".format(int(v), 16)
            elif k in ["uuidW0", "uuidW1", "uuidW2", "uuidW3"]:
                obj[k] = "0x{0:0{1}X}".format(v, 8)
    elif isinstance(obj, list):
        for v in obj:
            if isinstance(v, dict) or isinstance(v, list):
                prettify_uuids(v)


def prettify_cell_storage(obj):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, dict) or isinstance(v, list):
                prettify_cell_storage(v)
            elif k == "cellStorageBuffer" or k == "cellStorageBufferPreBnc":
                if sys.version_info.minor >= 8:
                    obj[k] = str(hexlify(b64decode(obj[k]), sep=":"))
                else:
                    obj[k] = str(hexlify(b64decode(obj[k])))
                obj[k] = obj[k].replace("b'", "").replace("'", "")
            elif k == "cellOffsets" or k == "cellOffsetsPreBnc":
                offsets = array("h", b64decode(obj[k])).tolist()
                obj[k] = ",".join([str(x) for x in offsets])
                obj[k] = regex.sub(r"(?:,-1)+$", ",[...]", obj[k])
    elif isinstance(obj, list):
        for v in obj:
            if isinstance(v, dict) or isinstance(v, list):
                prettify_cell_storage(v)


def process_file(filename, blob, output_dir, args):
    filename = regex.sub(r".*\.numbers/", "", filename)
    ensure_directory_exists(output_dir, filename)
    target_path = os.path.join(output_dir, filename)
    if isinstance(blob, IWAFile):
        target_path = target_path.replace(".iwa", "")
        target_path += ".json"
        with open(target_path, "w") as out:
            data = blob.to_dict()
            if args.hex_uuids or args.pretty:
                prettify_uuids(data)
            if args.pretty_storage or args.pretty:
                prettify_cell_storage(data)
            if args.compact_json or args.pretty:
                formatter = Formatter()
                formatter.indent_spaces = 2
                formatter.max_inline_complexity = 50
                formatter.max_compact_list_complexity = 50
                formatter.max_inline_length = 160
                formatter.max_compact_list_complexity = 2
                formatter.simple_bracket_padding = True
                formatter.nested_bracket_padding = False
                formatter.always_expand_depth = 10
                pretty_json = formatter.serialize(data)
                out.write(pretty_json)
            else:
                json.dump(data, out, sort_keys=True, indent=2)
    elif not filename.endswith("/"):
        with open(target_path, "wb") as out:
            out.write(blob)


def main():
    parser = ArgumentParser()
    parser.add_argument("document", help="Apple Numbers file(s)", nargs="*")
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument("--hex-uuids", action="store_true", help="print UUIDs as hex")
    parser.add_argument(
        "--pretty-storage", action="store_true", help="pretty print cell storage"
    )
    parser.add_argument(
        "--compact-json", action="store_true", help="Format JSON compactly as possible"
    )
    parser.add_argument(
        "--pretty", action="store_true", help="Enable all prettifying options"
    )
    parser.add_argument("--output", "-o", help="directory name to unpack into")
    args = parser.parse_args()
    if args.version:
        print(_get_version())
    elif args.output is not None and len(args.document) > 1:
        print(
            "unpack-numbers: error: output directory only valid with a single document",
            file=sys.stderr,
        )
        sys.exit(1)
    elif len(args.document) == 0:
        parser.print_help()
    else:
        for document in args.document:
            output_dir = args.output or document.replace(".numbers", "")
            try:
                read_numbers_file(
                    document,
                    file_handler=lambda filename, blob: process_file(
                        filename, blob, output_dir, args
                    ),
                )
            except FileFormatError as e:  # pragma: no cover
                print(f"{document}:", str(e), file=sys.stderr)
                sys.exit(1)


if __name__ == "__main__":
    # execute only if run as a script
    main()
