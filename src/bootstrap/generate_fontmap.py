import Cocoa

manager = Cocoa.NSFontManager.sharedFontManager()
font_families = list(manager.availableFontFamilies())

print("FONT_NAME_MAP = {")
for family in sorted(font_families):
    fonts = manager.availableMembersOfFontFamily_(family)
    for font in fonts:
        print(f'    "{font[0]}": "{family}",')
print("}")
