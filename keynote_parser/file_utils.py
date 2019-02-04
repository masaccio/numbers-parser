from __future__ import print_function
from __future__ import absolute_import
import os
import yaml

from PIL import Image
from io import BytesIO

from glob import glob
from tqdm import tqdm
from zipfile import ZipFile

from .codec import IWAFile


def ensure_directory_exists(prefix, path):
    """Ensure that a path's directory exists."""
    parts = os.path.split(path)
    try:
        os.makedirs(os.path.join(*([prefix] + list(parts[:-1]))))
    except OSError:
        pass


def unpack(filepath, target_dir=None, replacements=[]):
    """Unpack a .key file into a directory, writing iwa files as yaml files."""
    if not target_dir:
        target_dir = filepath.replace('.key', '')
    with ZipFile(filepath, 'r') as zipfile:
        for zipinfo in tqdm(zipfile.filelist):
            if zipinfo.filename.endswith('/'):
                continue
            with zipfile.open(zipinfo) as f:
                file_contents = f.read()
                ensure_directory_exists(target_dir, zipinfo.filename)
                target_path = os.path.join(target_dir, zipinfo.filename)
                if zipinfo.filename.endswith('iwa'):
                    try:
                        data = IWAFile.from_buffer(file_contents).to_dict()
                        for replacement in replacements:
                            data = replacement.perform_on(data)
                        content = yaml.safe_dump(
                            data, default_flow_style=False)
                        with open(target_path + '.yaml', 'w') as out:
                            out.write(content)
                        continue
                    except Exception as e:
                        print("Failed to unpack %s" % zipinfo.filename)
                        print(e)
                with open(target_path, 'wb') as out:
                    out.write(file_contents)


def ls(filepath):
    with ZipFile(filepath, 'r') as zipfile:
        for zipinfo in sorted(zipfile.filelist, key=lambda x: x.filename):
            if zipinfo.filename.endswith('/'):
                continue
            print(zipinfo.filename)


def cat(filepath, filename, replacements=[], raw=False):
    with ZipFile(filepath, 'r') as zipfile:
        with zipfile.open(filename) as f:
            file_contents = f.read()
            if filename.endswith('iwa') and not raw:
                data = IWAFile.from_buffer(file_contents).to_dict()
                for replacement in replacements:
                    data = replacement.perform_on(data)
                content = yaml.safe_dump(
                    data, default_flow_style=False)
                print(content)
            else:
                print(file_contents)


def replace(filepath, output_path, replacements=[]):
    files_to_write = {}
    completed_replacements = []

    def on_replace(replacement, old, new):
        completed_replacements.append((old, new))

    print("Reading from %s..." % filepath)
    with ZipFile(filepath, 'r') as zipfile:
        for zipinfo in tqdm(zipfile.filelist, desc='Finding...'):
            if zipinfo.filename.endswith('iwa'):
                with zipfile.open(zipinfo, 'r') as f:
                    file_contents = f.read()
                file = IWAFile.from_buffer(file_contents)
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
                        data = replacement.perform_on(
                            data,
                            on_replace=on_replace)
                    files_to_write[zipinfo.filename] = \
                        IWAFile.from_dict(data).to_buffer()
                    continue
            if zipinfo.filename.startswith("Data/"):
                file_has_changed = False
                for replacement in replacements:
                    repl_filepart, repl_ext = replacement.find.split(".")
                    data_filename = zipinfo.filename.replace("Data/", "")
                    if data_filename.startswith(repl_filepart):
                        # Scale this file to the appropriate size
                        with zipfile.open(zipinfo, 'r') as f:
                            image = Image.open(f)
                        with open(replacement.replace, 'rb') as f:
                            read_image = Image.open(f)
                            with BytesIO() as output:
                                read_image \
                                    .thumbnail(image.size, Image.ANTIALIAS)
                                read_image.save(output, image.format)
                                files_to_write[zipinfo.filename] = \
                                    output.getvalue()
                        file_has_changed = True
                        break
                if file_has_changed:
                    continue
            with zipfile.open(zipinfo, 'r') as f:
                files_to_write[zipinfo.filename] = f.read()
    print("Writing to %s..." % output_path)
    with ZipFile(output_path, 'w') as zipfile:
        for filename, contents in tqdm(iter(list(files_to_write.items())),
                                       total=len(files_to_write),
                                       desc='Replacing...'):
            zipfile.writestr(filename, contents)

    return completed_replacements


def pack(filepath, target_file=None, replacements=[]):
    """Pack a directory into a .key file, writing yaml files as iwa files."""
    if target_file is None:
        target_file = filepath + '.out.key'
    with ZipFile(target_file, 'w') as zipfile:
        all_files = glob(filepath + '/**/**') + glob(filepath + '/**')
        for filename in tqdm(all_files):
            if os.path.isdir(filename):
                continue
            with open(filename) as f:
                file_contents = f.read()
                zip_filename = filename.replace(filepath + '/', '')
                if zip_filename.endswith('.yaml'):
                    try:
                        data = yaml.load(file_contents)
                        for replacement in replacements:
                            data = replacement.perform_on(data)
                        data = IWAFile.from_dict(data).to_buffer()
                        zipfile.writestr(
                            zip_filename.replace('.yaml', ''),
                            data)
                        continue
                    except Exception:
                        print("Failed to pack %s" % filename)
                        raise
                zipfile.writestr(zip_filename, file_contents)
