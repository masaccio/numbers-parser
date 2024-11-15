import argparse

from numbers_parser import Document


def dump_format(s):
    pass


def dump_custom_formats_for_doc(doc):
    objects = doc._model.objects
    custom_format_id = doc._model.find_refs("CustomFormatListArchive")[0]
    custom_formats = objects[custom_format_id].custom_formats
    for custom_format in sorted(custom_formats, key=lambda x: x.name):
        print(custom_format)


parser = argparse.ArgumentParser()
parser.add_argument(
    "numbers",
    nargs="*",
    metavar="numbers-filename",
    help="Numbers folders/files to dump",
)
args = parser.parse_args()

for filename in args.numbers:
    doc = Document(filename)
    dump_custom_formats_for_doc(doc)
