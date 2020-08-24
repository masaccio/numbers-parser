from __future__ import print_function
from __future__ import absolute_import
import argparse

import keynote_parser

keynote_parser.__command_line_invocation__ = True  # noqa

from keynote_parser import __version__, __supported_keynote_version__
from collections import Counter
from .bundle_utils import warn_once_on_newer_keynote, get_installed_keynote_version
from .file_utils import process
from .replacement import Replacement, parse_json


def parse_replacements(**kwargs):
    json_filename = kwargs.get('replacements')
    if json_filename:
        return parse_json(json_filename)
    else:
        return []


def unpack_command(input, output=None, **kwargs):
    process(input, output or input.replace('.key', ''), replacements=parse_replacements(**kwargs))


def pack_command(input, output=None, **kwargs):
    process(input, output or (input + ".key"), replacements=parse_replacements(**kwargs))


def ls_command(input, **kwargs):
    process(input, '-')


def cat_command(input, filename, **kwargs):
    process(
        input,
        '-',
        subfile=filename,
        replacements=parse_replacements(**kwargs),
        raw=kwargs.get('raw'),
    )


def replace_command(input, **kwargs):
    output = kwargs.get('output', None) or input
    replacements = parse_replacements(**kwargs)
    find, _replace = kwargs.get('find'), kwargs.get('replace')
    if find and _replace:
        replacements.append(Replacement(find, _replace))
    if not replacements:
        print("WARNING: No replacements passed. No change.")
        return
    for ((old, new), count) in list(Counter(process(input, output, replacements)).items()):
        if count == 1:
            print("Replaced %s with %s." % (repr(old), repr(new)))
        else:
            print("Replaced %s with %s %d times." % (repr(old), repr(new), count))


def add_replacement_arg(parser):
    parser.add_argument("--replacements", help="apply replacements from a json or yaml file")


def main():
    installed_version = get_installed_keynote_version()
    install_warning = ""
    if installed_version and __supported_keynote_version__ < installed_version:
        install_warning = (
            " (Installed Keynote version %s not yet supported.)" % installed_version
        )
    parser = argparse.ArgumentParser(
        description=(
            "manipulate Apple Keynote .key files. version %s, supports Keynote versions up to %s.%s"
        )
        % (__version__, __supported_keynote_version__, install_warning)
    )
    parser.add_argument('-v', '--version', action='version', version=__version__)

    subparsers = parser.add_subparsers()

    parser_unpack = subparsers.add_parser('unpack')
    parser_unpack.add_argument("input", help="a .key file")
    parser_unpack.add_argument("--output", "-o", help="a directory name to unpack into")
    add_replacement_arg(parser_unpack)
    parser_unpack.set_defaults(func=unpack_command)

    parser_pack = subparsers.add_parser('pack')
    parser_pack.add_argument("input", help="a directory of an unpacked .key file")
    parser_pack.add_argument("--output", "-o", help="a keynote file name to unpack into")
    add_replacement_arg(parser_pack)
    parser_pack.set_defaults(func=pack_command)

    parser_ls = subparsers.add_parser('ls')
    parser_ls.add_argument("input", help="a .key file")
    parser_ls.set_defaults(func=ls_command)

    parser_cat = subparsers.add_parser('cat')
    parser_cat.add_argument("input", help="a .key file")
    parser_cat.add_argument(
        "filename", help="a file within that .key file to cat, decoding .iwa to .yaml"
    )
    parser_cat.add_argument(
        "--raw", action='store_true', help="always return the original file with no decoding"
    )
    add_replacement_arg(parser_cat)
    parser_cat.set_defaults(func=cat_command)

    parser_replace = subparsers.add_parser('replace')
    parser_replace.add_argument("input", help="a .key file")
    parser_replace.add_argument("--output", "-o", help="a .key file to output to")
    parser_replace.add_argument("--find", help="a pattern to search for in text")
    parser_replace.add_argument("--replace", help="a string to replace with")
    add_replacement_arg(parser_replace)
    parser_replace.set_defaults(func=replace_command)

    args = parser.parse_args()
    if hasattr(args, 'func'):
        warn_once_on_newer_keynote()
        args.func(**vars(args))
    else:
        parser.print_help()
        print()
        warn_once_on_newer_keynote()


if __name__ == "__main__":
    main()
