import argparse

import numbers_parser

numbers_parser.__command_line_invocation__ = True  # noqa

from numbers_parser import __version__, __supported_numbers_version__
from numbers_parser.bundle_utils import (
    warn_once_on_newer_numbers,
    get_installed_numbers_version,
)
from numbers_parser.file_utils import process


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-v", "--version", action="version", version=__version__)
    parser.add_argument("input", help="a .numbers file")
    parser.add_argument("--output", "-o", help="a directory name to unpack into")
    args = parser.parse_args()
    warn_once_on_newer_numbers()
    process(args.input, args.output or args.input.replace(".numbers", ""))


if __name__ == "__main__":
    main()
