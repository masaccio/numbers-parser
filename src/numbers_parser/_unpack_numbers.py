# vi: ft=python

import argparse
import os
import json
import re
import sys

from numbers_parser.unpack import read_numbers_file
from numbers_parser import _get_version
from numbers_parser.iwafile import IWAFile


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


def process_file(contents, filename, output_dir, hex_uuids):
    filename = re.sub(".*\.numbers/", "", filename)
    ensure_directory_exists(output_dir, filename)
    target_path = os.path.join(output_dir, filename)
    if isinstance(contents, IWAFile):
        target_path = target_path.replace(".iwa", "")
        target_path += ".txt"
        with open(target_path, "w") as out:
            data = contents.to_dict()
            if hex_uuids:
                convert_uuids_to_hex(data)
            print(json.dumps(data, sort_keys=True, indent=4), file=out)
    else:
        with open(target_path, "wb") as out:
            out.write(contents)


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
            read_numbers_file(
                document,
                handler=lambda contents, filename: process_file(
                    contents, filename, output_dir, args.hex_uuids
                ),
                store_objects=False,
            )


if __name__ == "__main__":
    # execute only if run as a script
    main()
