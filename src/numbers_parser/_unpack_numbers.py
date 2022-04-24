# vi: ft=python

import argparse
import os
import json
import re
import sys

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


def convert_uuids_to_hex(obj):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, dict) or isinstance(v, list):
                convert_uuids_to_hex(v)
            elif k == "lower" or k == "upper":
                obj[k] = "0x{0:0{1}X}".format(int(v), 16)
            elif k in ["uuidW0", "uuidW1", "uuidW2", "uuidW3"]:
                obj[k] = "0x{0:0{1}X}".format(v, 8)
    elif isinstance(obj, list):
        for v in obj:
            if isinstance(v, dict) or isinstance(v, list):
                convert_uuids_to_hex(v)


def process_file(filename, blob, output_dir, hex_uuids):
    filename = re.sub(r".*\.numbers/", "", filename)
    ensure_directory_exists(output_dir, filename)
    target_path = os.path.join(output_dir, filename)
    if isinstance(blob, IWAFile):
        target_path = target_path.replace(".iwa", "")
        target_path += ".txt"
        with open(target_path, "w") as out:
            data = blob.to_dict()
            if hex_uuids:
                convert_uuids_to_hex(data)
            print(json.dumps(data, sort_keys=True, indent=4), file=out)
    elif not filename.endswith("/"):
        with open(target_path, "wb") as out:
            out.write(blob)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("document", help="Apple Numbers file(s)", nargs="*")
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument("--hex-uuids", action="store_true", help="print UUIDs as hex")
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
                        filename, blob, output_dir, args.hex_uuids
                    ),
                )
            except FileFormatError as e:
                print(f"{document}:", str(e), file=sys.stderr)
                sys.exit(1)


if __name__ == "__main__":
    # execute only if run as a script
    main()
