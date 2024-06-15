import contextlib
import json
import logging
import os
import re
import sys
from argparse import ArgumentParser
from array import array
from base64 import b64decode
from binascii import hexlify
from dataclasses import dataclass
from pathlib import Path

import regex
from compact_json import Formatter

from numbers_parser import __name__ as numbers_parser_name
from numbers_parser import _get_version
from numbers_parser.constants import SUPPORTED_NUMBERS_VERSIONS
from numbers_parser.exceptions import FileError, FileFormatError, UnsupportedError
from numbers_parser.iwafile import IWAFile
from numbers_parser.iwork import IWork, IWorkHandler
from numbers_parser.numbers_uuid import NumbersUUID

logger = logging.getLogger(numbers_parser_name)


@dataclass
class NumbersUnpacker(IWorkHandler):
    hex_uuids: bool = False
    pretty_storage: bool = False
    pretty: bool = False
    compact_json: bool = False
    output_dir: str = None

    def store_file(self, filename: str, blob: bytes) -> None:
        """Store a profobuf archive."""
        filename = regex.sub(r".*\.numbers/", "", str(filename))
        self.ensure_directory_exists(filename)
        target_path = os.path.join(self.output_dir, filename)
        if isinstance(blob, IWAFile):
            target_path = target_path.replace(".iwa", "")
            target_path += ".json"
            with open(target_path, "w") as out:
                data = blob.to_dict()
                if self.hex_uuids or self.pretty:
                    self.prettify_uuids(data)
                if self.pretty_storage or self.pretty:
                    self.prettify_cell_storage(data)
                if self.compact_json or self.pretty:
                    formatter = Formatter()
                    formatter.indent_spaces = 2
                    formatter.max_inline_complexity = 100
                    formatter.max_compact_list_complexity = 100
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

    def ensure_directory_exists(self, path: str):
        """Ensure that a path's directory exists."""
        parts = os.path.split(path)
        with contextlib.suppress(OSError):
            os.makedirs(os.path.join(*([self.output_dir, *list(parts[:-1])])))

    def prettify_uuids(self, obj: object):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if isinstance(v, dict):
                    try:
                        obj[k] = str(NumbersUUID(v))
                    except UnsupportedError:
                        self.prettify_uuids(v)
                elif isinstance(v, list):
                    self.prettify_uuids(v)
        else:  # list
            for i, v in enumerate(obj):
                if isinstance(v, dict):
                    try:
                        obj[i] = str(NumbersUUID(v))
                    except UnsupportedError:
                        self.prettify_uuids(v)
                elif isinstance(v, list):
                    self.prettify_uuids(v)

    def prettify_cell_storage(self, obj: object):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if isinstance(v, (dict, list)):
                    self.prettify_cell_storage(v)
                elif k in ["cell_storage_buffer", "cell_storage_buffer_pre_bnc"]:
                    obj[k] = str(hexlify(b64decode(obj[k]), sep=":"))
                    obj[k] = obj[k].replace("b'", "").replace("'", "")
                elif k in ["cell_offsets", k == "cell_offsets_pre_bnc"]:
                    offsets = array("h", b64decode(obj[k])).tolist()
                    obj[k] = ",".join([str(x) for x in offsets])
                    obj[k] = regex.sub(r"(?:,-1)+$", ",[...]", obj[k])
        else:  # list
            for v in obj:
                if isinstance(v, (dict, list)):
                    self.prettify_cell_storage(v)

    def allowed_format(self, extension: str) -> bool:
        """bool: Return ``True`` if the filename extension is supported by the handler."""
        return extension == ".numbers"

    def allowed_version(self, version: str) -> bool:
        """bool: Return ``True`` if the document version is allowed."""
        version = re.sub(r"(\d+)\.(\d+)\.\d+", r"\1.\2", version)
        return version in SUPPORTED_NUMBERS_VERSIONS


def main():
    parser = ArgumentParser()
    parser.add_argument("document", help="Apple Numbers file(s)", nargs="*")
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument("--hex-uuids", action="store_true", help="print UUIDs as hex")
    parser.add_argument("--pretty-storage", action="store_true", help="pretty print cell storage")
    parser.add_argument(
        "--compact-json",
        action="store_true",
        help="Format JSON compactly as possible",
    )
    parser.add_argument("--pretty", action="store_true", help="Enable all prettifying options")
    parser.add_argument("--output", "-o", help="directory name to unpack into")
    parser.add_argument("--debug", default=False, action="store_true", help="Enable debug logging")
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
        hdlr = logging.StreamHandler()
        hdlr.setFormatter(logging.Formatter("%(levelname)s:%(name)s:%(message)s"))
        logger.addHandler(hdlr)
        if args.debug:
            logger.setLevel("DEBUG")
        else:
            logger.setLevel("ERROR")
        for document in args.document:
            output_dir = args.output or document.replace(".numbers", "")
            try:
                iwork = IWork(
                    handler=NumbersUnpacker(
                        hex_uuids=args.hex_uuids,
                        pretty=args.pretty,
                        pretty_storage=args.pretty_storage,
                        compact_json=args.compact_json,
                        output_dir=output_dir,
                    ),
                )
                iwork.open(Path(document))
            except (FileFormatError, FileError) as e:
                print(f"{document}:", str(e), file=sys.stderr)
                sys.exit(1)


if __name__ == "__main__":
    # execute only if run as a script
    main()
