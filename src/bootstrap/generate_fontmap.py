import sys
import Cocoa

if len(sys.argv) != 2:
    raise (ValueError(f"Usage: {sys.argv[0]} fontmap.py"))

mapping_py = sys.argv[1]

with open(mapping_py, "w") as fh:
    manager = Cocoa.NSFontManager.sharedFontManager()
    font_families = list(manager.availableFontFamilies())

    print("FONT_NAME_TO_FAMILY = {", file=fh)
    for family in sorted(font_families):
        fonts = manager.availableMembersOfFontFamily_(family)
        for font in fonts:
            print(f'    "{font[0]}": "{family}",', file=fh)
    print('    "Calibri": "Calibri",', file=fh)
    print('    "Cambria": "Cambria",', file=fh)
    print("}", file=fh)
