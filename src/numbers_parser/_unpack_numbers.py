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


def process_file(contents, filename, output_dir):
    filename = re.sub(".*\.numbers/", "", filename)
    ensure_directory_exists(output_dir, filename)
    target_path = os.path.join(output_dir, filename)
    if isinstance(contents, IWAFile):
        target_path = target_path.replace(".iwa", "")
        target_path += ".txt"
        with open(target_path, "w") as out:
            print(json.dumps(contents.to_dict(), sort_keys=True, indent=4), file=out)
    else:
        with open(target_path, "wb") as out:
            out.write(contents)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("document", help="Apple Numbers file(s)", nargs="*")
    parser.add_argument("-V", "--version", action="store_true")
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
    else:
        for document in args.document:
            output_dir = args.output or document.replace(".numbers", "")
            read_numbers_file(
                document,
                handler=lambda contents, filename: process_file(contents, filename, output_dir),
                store_objects=False,
            )


if __name__ == "__main__":
    # execute only if run as a script
    main()
