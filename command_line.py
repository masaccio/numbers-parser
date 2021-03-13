import argparse

import numbers_parser

numbers_parser.__command_line_invocation__ = True  # noqa

from numbers_parser import __version__, __supported_numbers_version__
from collections import Counter
from numbers_parser.bundle_utils import (
    warn_once_on_newer_numbers,
    get_installed_numbers_version,
)
from numbers_parser.file_utils import process


def unpack_command(input, output=None, **kwargs):
    process(
        input,
        output or input.replace(".numbers", ""),
    )


def pack_command(input, output=None, **kwargs):
    process(
        input, output or (input + ".numbers"),
    )


def ls_command(input, **kwargs):
    process(input, "-")


def cat_command(input, filename, **kwargs):
    process(
        input,
        "-",
        subfile=filename,
        raw=kwargs.get("raw"),
    )


def main():
    installed_version = get_installed_numbers_version()
    install_warning = ""
    if installed_version and __supported_numbers_version__ < installed_version:
        install_warning = (
            " (Installed Numbers version %s not yet supported.)" % installed_version
        )
    parser = argparse.ArgumentParser(
        description=(
            "manipulate Apple Numbers .numbers files. version %s, supports Numbers versions up to %s.%s"
        )
        % (__version__, __supported_numbers_version__, install_warning)
    )
    parser.add_argument("-v", "--version", action="version", version=__version__)

    subparsers = parser.add_subparsers()

    parser_unpack = subparsers.add_parser("unpack")
    parser_unpack.add_argument("input", help="a .numbers file")
    parser_unpack.add_argument("--output", "-o", help="a directory name to unpack into")
    parser_unpack.set_defaults(func=unpack_command)

    parser_pack = subparsers.add_parser("pack")
    parser_pack.add_argument("input", help="a directory of an unpacked .numbers file")
    parser_pack.add_argument(
        "--output", "-o", help="a numbers file name to unpack into"
    )
    parser_pack.set_defaults(func=pack_command)

    parser_ls = subparsers.add_parser("ls")
    parser_ls.add_argument("input", help="a .numbers file")
    parser_ls.set_defaults(func=ls_command)

    parser_cat = subparsers.add_parser("cat")
    parser_cat.add_argument("input", help="a .numbers file")
    parser_cat.add_argument(
        "filename",
        help="a file within that .numbers file to cat, decoding .iwa to .yaml",
    )
    parser_cat.add_argument(
        "--raw",
        action="store_true",
        help="always return the original file with no decoding",
    )
    parser_cat.set_defaults(func=cat_command)

    args = parser.parse_args()
    if hasattr(args, "func"):
        warn_once_on_newer_numbers()
        args.func(**vars(args))
    else:
        parser.print_help()
        print()
        warn_once_on_newer_numbers()


if __name__ == "__main__":
    main()
