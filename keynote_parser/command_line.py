import os
import yaml
import argparse

from codec import IWAFile
from file_utils import pack, unpack


def main():
    parser = argparse.ArgumentParser(
        description="pack or unpack .key files for Apple Keynote.")
    parser.add_argument(
        "path",
        help="a .key file or directory containing an unpacked .key file")
    args = parser.parse_args()

    if args.path.endswith('.key'):
        unpack(args.path)
    elif args.path.endswith('.iwa'):
        iwa_file = IWAFile.from_file(args.path)
        print yaml.safe_dump(
            iwa_file.to_dict(),
            default_flow_style=False)
    elif args.path.endswith('.yaml'):
        iwa_file = IWAFile.from_yaml(args.path)
        print iwa_file.to_buffer(),
    elif os.path.isdir(args.path):
        pack(args.path)
    else:
        parser.print_help()
