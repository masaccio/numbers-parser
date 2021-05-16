# vi: ft=python

import argparse
import os
import json
import sys

from contextlib import contextmanager
from glob import glob
from numbers_parser.codec import IWAFile
from numbers_parser import _get_version
from zipfile import ZipFile


def ensure_directory_exists(prefix, path):
    """Ensure that a path's directory exists."""
    parts = os.path.split(path)
    try:
        os.makedirs(os.path.join(*([prefix] + list(parts[:-1]))))
    except OSError:
        pass


def file_reader(path):
    if not os.path.isdir(path) and path.endswith(".numbers"):
        return zip_file_reader(path)
    else:
        return directory_reader(path)


def zip_file_reader(path):
    zipfile = ZipFile(path, "r")
    for _file in zipfile.filelist:
        _file.filename = _file.filename.encode("cp437").decode("utf-8")
    iterator = sorted(zipfile.filelist, key=lambda x: x.filename)
    for zipinfo in iterator:
        if zipinfo.filename.endswith("/"):
            continue
        with zipfile.open(zipinfo) as handle:
            yield (zipinfo.filename, handle)


def directory_reader(path):
    # Python <3.5 doesn't support glob with recursive, so this will have to do.
    iterator = set(sum([glob(path + ("/**" * i)) for i in range(10)], []))
    iterator = sorted(iterator)
    for filename in iterator:
        if os.path.isdir(filename):
            continue
        rel_filename = filename.replace(path + "/", "")
        with open(filename, "rb") as handle:
            yield (rel_filename, handle)


@contextmanager
def dir_file_sink(target_dir):
    def accept(filename, contents):
        ensure_directory_exists(target_dir, filename)
        target_path = os.path.join(target_dir, filename)
        if isinstance(contents, IWAFile):
            target_path = target_path.replace(".iwa", "")
            target_path += ".txt"
            with open(target_path, "w") as out:
                print(json.dumps(contents.to_dict(), sort_keys=True, indent=4), file=out)
        else:
            with open(target_path, "wb") as out:
                if isinstance(contents, IWAFile):
                    out.write(contents.to_buffer())
                else:
                    out.write(contents)

    yield accept


def process_file(filename, handle, sink):
    contents = None
    if ".iwa" in filename:
        contents = handle.read()
        file = IWAFile.from_buffer(contents, filename)

        file_has_changed = False

        if file_has_changed:
            data = file.to_dict()
            sink(filename, IWAFile.from_dict(data))
        else:
            sink(filename, file)
        return

    if filename.startswith("Data/"):
        file_has_changed = False

        if file_has_changed:
            return
    sink(filename, contents or handle.read())


def process(input_path, output_path, subfile=None):
    with dir_file_sink(output_path) as sink:
        for filename, handle in file_reader(input_path):
            try:
                process_file(filename, handle, sink)
            except Exception as e:
                raise ValueError("Failed to process file %s due to: %s" % (filename, e))


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
            process(document, args.output or document.replace(".numbers", ""))

if __name__ == "__main__":
    # execute only if run as a script
    main()
