import sys

import Cocoa

if len(sys.argv) != 2:
    msg = f"Usage: {sys.argv[0]} fontmap.py"
    raise (ValueError(msg))

mapping_py = sys.argv[1]
with open(mapping_py, "w") as fh:
    print("GENERATED_FONT_MAP = {", file=fh)

    manager = Cocoa.NSFontManager.sharedFontManager()
    font_families = list(manager.availableFontFamilies())
    for family in sorted(font_families):
        fonts = manager.availableMembersOfFontFamily_(family)
        for font in fonts:
            (name, style, _, traits) = font
            print(f'   "{name}": {{', file=fh)
            print(f'        "name": "{name}",', file=fh)
            print(f'        "family": "{family}",', file=fh)
            print(f'        "style": "{style}",', file=fh)
            print(f'        "bold": {bool(traits & 2)},', file=fh)
            print(f'        "italic": {bool(traits & 1)},', file=fh)
            print("    },", file=fh)
    print(
        """   "Calibri": {
        "name": "Calibri",
        "family": "Calibri",
        "style": "Regular",
        "bold": False,
        "italic": False,
    },
   "Cambria": {
        "name": "Cambria",
        "family": "Cambria",
        "style": "Regular",
        "bold": False,
        "italic": False,
    },""",
        file=fh,
    )
    print("}", file=fh)
