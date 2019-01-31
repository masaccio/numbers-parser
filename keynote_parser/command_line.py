from __future__ import print_function
from __future__ import absolute_import
import argparse

from collections import Counter
from .file_utils import pack, unpack, ls, cat, replace
from .replacement import Replacement, parse_json


def parse_replacements(**kwargs):
    json_filename = kwargs.get('replacements')
    if json_filename:
        return parse_json(json_filename)
    else:
        return []


def unpack_command(path, output=None, **kwargs):
    unpack(
        path,
        target_dir=output,
        replacements=parse_replacements(**kwargs))


def pack_command(path, output=None, **kwargs):
    pack(
        path,
        target_file=output,
        replacements=parse_replacements(**kwargs))


def ls_command(path, **kwargs):
    ls(path)


def cat_command(path, filename, **kwargs):
    cat(
        path,
        filename,
        replacements=parse_replacements(**kwargs),
        raw=kwargs.get('raw'))


def replace_command(path, **kwargs):
    output = kwargs.get('output', None) or path
    replacements = parse_replacements(**kwargs)
    find, _replace = kwargs.get('find'), kwargs.get('replace')
    if find and _replace:
        replacements.append(Replacement(find, _replace))
    if not replacements:
        print("WARNING: No replacements passed. No change.")
        return
    print(replacements)
    for ((old, new), count) in list(Counter(
            replace(path, output, replacements)).items()):
        if count == 1:
            print("Replaced %s with %s." % (repr(old), repr(new)))
        else:
            print("Replaced %s with %s %d times." % (
                repr(old), repr(new), count))


def add_replacement_arg(parser):
    parser.add_argument(
        "--replacements",
        help="apply replacements from a json or yaml file")


def main():
    parser = argparse.ArgumentParser(
        description="manipulate Apple Keynote .key files")
    subparsers = parser.add_subparsers()

    parser_unpack = subparsers.add_parser('unpack')
    parser_unpack.add_argument("path", help="a .key file")
    parser_unpack.add_argument(
        "--output",
        help="a directory name to unpack into")
    add_replacement_arg(parser_unpack)
    parser_unpack.set_defaults(func=unpack_command)

    parser_pack = subparsers.add_parser('pack')
    parser_pack.add_argument(
        "path",
        help="a directory of an unpacked .key file")
    parser_pack.add_argument(
        "--output",
        help="a keynote file name to unpack into")
    add_replacement_arg(parser_pack)
    parser_pack.set_defaults(func=pack_command)

    parser_ls = subparsers.add_parser('ls')
    parser_ls.add_argument("path", help="a .key file")
    parser_ls.set_defaults(func=ls_command)

    parser_cat = subparsers.add_parser('cat')
    parser_cat.add_argument("path", help="a .key file")
    parser_cat.add_argument(
        "filename",
        help="a file within that .key file to cat, decoding .iwa to .yaml")
    parser_cat.add_argument(
        "--raw",
        action='store_true',
        help="always return the original file with no decoding")
    add_replacement_arg(parser_cat)
    parser_cat.set_defaults(func=cat_command)

    parser_replace = subparsers.add_parser('replace')
    parser_replace.add_argument("path", help="a .key file")
    parser_replace.add_argument("--output", help="a .key file to output to")
    parser_replace.add_argument(
        "--find",
        help="a pattern to search for in text")
    parser_replace.add_argument(
        "--replace",
        help="a string to replace with")
    add_replacement_arg(parser_replace)
    parser_replace.set_defaults(func=replace_command)

    args = parser.parse_args()
    args.func(**vars(args))


if __name__ == "__main__":
    main()
