from __future__ import print_function
from __future__ import absolute_import

import os
import sys

try:
    from yaml import CLoader as Loader, CDumper as Dumper
except ImportError:
    from yaml import Loader, Dumper
import yaml

from contextlib import contextmanager

from PIL import Image
from io import BytesIO

from glob import glob
from tqdm import tqdm
from zipfile import ZipFile

from .codec import IWAFile
from .unicode_utils import fix_unicode


def ensure_directory_exists(prefix, path):
    """Ensure that a path's directory exists."""
    parts = os.path.split(path)
    try:
        os.makedirs(os.path.join(*([prefix] + list(parts[:-1]))))
    except OSError:
        pass


def file_reader(path, progress=True):
    if path.endswith('.key'):
        return zip_file_reader(path, progress)
    else:
        return directory_reader(path, progress)


def zip_file_reader(path, progress=True):
    zipfile = ZipFile(path, "r")
    for _file in zipfile.filelist:
        _file.filename = _file.filename.encode("cp437").decode("utf-8")
    iterator = sorted(zipfile.filelist, key=lambda x: x.filename)
    if progress:
        iterator = tqdm(iterator)
    for zipinfo in iterator:
        if zipinfo.filename.endswith('/'):
            continue
        if progress:
            iterator.set_description("Reading {}...".format(zipinfo.filename))
        with zipfile.open(zipinfo) as handle:
            yield (zipinfo.filename, handle)


def directory_reader(path, progress=True):
    # Python <3.5 doesn't support glob with recursive, so this will have to do.
    iterator = set(sum([glob(path + ('/**' * i)) for i in range(10)], []))
    iterator = sorted(iterator)
    if progress:
        iterator = tqdm(iterator)
    for filename in iterator:
        if os.path.isdir(filename):
            continue
        rel_filename = filename.replace(path + '/', '')
        if progress:
            iterator.set_description("Reading {}...".format(rel_filename))
        with open(filename, 'rb') as handle:
            yield (rel_filename, handle)


def file_sink(path, raw=False, subfile=None):
    if path == "-":
        if subfile:
            return cat_sink(subfile, raw)
        else:
            return ls_sink()
    if path.endswith('.key'):
        return zip_file_sink(path)
    return dir_file_sink(path, raw=raw)


@contextmanager
def dir_file_sink(target_dir, raw=False):
    def accept(filename, contents):
        ensure_directory_exists(target_dir, filename)
        target_path = os.path.join(target_dir, filename)
        if isinstance(contents, IWAFile) and not raw:
            target_path += ".yaml"
        with open(target_path, 'wb') as out:
            if isinstance(contents, IWAFile):
                if raw:
                    out.write(contents.to_buffer())
                else:
                    yaml.dump(
                        contents.to_dict(),
                        out,
                        default_flow_style=False,
                        encoding="utf-8",
                        Dumper=Dumper,
                    )
            else:
                out.write(contents)

    accept.uses_stdout = False
    yield accept


@contextmanager
def ls_sink():
    def accept(filename, contents):
        print(filename)

    accept.uses_stdout = True
    yield accept


@contextmanager
def cat_sink(subfile, raw):
    def accept(filename, contents):
        if filename == subfile:
            if isinstance(contents, IWAFile):
                if raw:
                    sys.stdout.buffer.write(contents.to_buffer())
                else:
                    print(
                        yaml.dump(
                            contents.to_dict(),
                            default_flow_style=False,
                            encoding="utf-8",
                            Dumper=Dumper,
                        ).decode('ascii')
                    )
            else:
                sys.stdout.buffer.write(contents)

    accept.uses_stdout = True
    yield accept


@contextmanager
def zip_file_sink(output_path):
    files_to_write = {}

    def accept(filename, contents):
        files_to_write[filename] = contents

    accept.uses_stdout = False

    yield accept

    print("Writing to %s..." % output_path)
    with ZipFile(output_path, 'w') as zipfile:
        for filename, contents in tqdm(
            iter(list(files_to_write.items())), total=len(files_to_write)
        ):
            if isinstance(contents, IWAFile):
                zipfile.writestr(filename, contents.to_buffer())
            else:
                zipfile.writestr(filename, contents)


def process_file(filename, handle, sink, replacements=[], raw=False, on_replace=None):
    contents = None
    if '.iwa' in filename and not raw:
        contents = handle.read()
        if filename.endswith('.yaml'):
            file = IWAFile.from_dict(
                yaml.load(fix_unicode(contents.decode('utf-8')), Loader=Loader)
            )
            filename = filename.replace('.yaml', '')
        else:
            file = IWAFile.from_buffer(contents, filename)

        file_has_changed = False
        for replacement in replacements:
            # Replacing in a file is expensive, so let's
            # avoid doing so if possible.
            if replacement.should_replace(file):
                file_has_changed = True
                break

        if file_has_changed:
            data = file.to_dict()
            for replacement in replacements:
                data = replacement.perform_on(data, on_replace=on_replace)
            sink(filename, IWAFile.from_dict(data))
        else:
            sink(filename, file)
        return

    if filename.startswith("Data/"):
        file_has_changed = False
        for replacement in replacements:
            find_parts = replacement.find.split(".")
            if len(find_parts) != 2:
                continue
            repl_filepart, repl_ext = find_parts
            data_filename = filename.replace("Data/", "")
            if data_filename.startswith(repl_filepart):
                # Scale this file to the appropriate size
                image = Image.open(handle)
                with open(replacement.replace, 'rb') as f:
                    read_image = Image.open(f)
                    with BytesIO() as output:
                        read_image.thumbnail(image.size, Image.ANTIALIAS)
                        read_image.save(output, image.format)
                        sink(filename, output.getvalue())
                file_has_changed = True
                break

        if file_has_changed:
            return
    sink(filename, contents or handle.read())


def process(input_path, output_path, replacements=[], subfile=None, raw=False):
    completed_replacements = []

    def on_replace(replacement, old, new):
        completed_replacements.append((old, new))

    with file_sink(output_path, subfile=subfile) as sink:
        if not sink.uses_stdout:
            print("Reading from %s..." % input_path)
        for filename, handle in file_reader(input_path, not sink.uses_stdout):
            try:
                process_file(filename, handle, sink, replacements, raw, on_replace)
            except Exception as e:
                raise ValueError("Failed to process file %s due to: %s" % (filename, e))

    return completed_replacements
