import os
import yaml

from glob import glob
from tqdm import tqdm
from zipfile import ZipFile

from codec import IWAFile


def ensure_directory_exists(prefix, path):
    """Ensure that a path's directory exists."""
    parts = os.path.split(path)
    try:
        os.makedirs(os.path.join(*([prefix] + list(parts[:-1]))))
    except OSError:
        pass


def unpack(filepath, target_dir=None):
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
                        content = yaml.safe_dump(
                            IWAFile.from_buffer(file_contents).to_dict(),
                            default_flow_style=False)
                        with open(target_path + '.yaml', 'w') as out:
                            out.write(content)
                        continue
                    except Exception as e:
                        print "Failed to unpack %s" % zipinfo.filename
                        print e
                with open(target_path, 'w') as out:
                    out.write(file_contents)


def pack(filepath):
    """Pack a directory into a .key file, writing yaml files as iwa files."""
    with ZipFile(filepath + '.out.key', 'w') as zipfile:
        all_files = glob(filepath + '/**/**') + glob(filepath + '/**')
        for filename in tqdm(all_files):
            if os.path.isdir(filename):
                continue
            with open(filename) as f:
                file_contents = f.read()
                zip_filename = filename.replace(filepath + '/', '')
                if zip_filename.endswith('.yaml'):
                    try:
                        file_contents = IWAFile.from_dict(
                            yaml.load(file_contents)).to_buffer()
                        zipfile.writestr(
                            zip_filename.replace('.yaml', ''),
                            file_contents)
                        continue
                    except Exception:
                        print "Failed to pack %s" % filename
                        raise
                zipfile.writestr(zip_filename, file_contents)
