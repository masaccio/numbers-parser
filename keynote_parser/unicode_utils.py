"""
Tools to handle yaml Unicode incompatibilities between Python 2 and 3.

tl;dr: when writing Yaml with emojis in it from Python 2 and reading from
Python 3 (or vice versa),  are incompatibilities with the unicode handling
that result in runtime exceptions. It's Bad.

This file aims to manually do some pre-parsing on the Yaml strings before
decoding to ensure that whichever version of Python is doing the decoding
can successfully parse the string.
"""

import re
import sys

PY2_SURROGATE_PAIR_RE = re.compile(r'\\u([A-Fa-f0-9]{4})\\u([A-Fa-f0-9]{4})')
PY3_MULTIBYTE_RE = re.compile(r'\\U([A-Fa-f0-9]{8})')


def from_surrogate_pair(high, low):
    high, low = int(high, 16), int(low, 16)
    value = (high - 0xD800) * 0x400 + (low - 0xDC00) + 0x10000
    if value > 0:
        return value


def to_surrogate_pair(input):
    input = int(input, 16)
    high = int((input - 0x10000) / 0x400) + 0xD800
    low = (input - 0x10000) % 0x400 + 0xDC00
    return high, low


def to_py3_compatible(input):
    """Convert an input string containing Unicode surrogate pairs to UTF-16"""
    for high, low in PY2_SURROGATE_PAIR_RE.findall(input):
        character = from_surrogate_pair(high, low)
        if not character:
            continue
        input = input.replace(
            "\\u%s\\u%s" % (high, low),
            "\\U%08x" % character)
        input = input.replace(
            "\\U%s\\U%s" % (high, low),
            "\\U%08x" % character)
    return input


def to_py2_compatible(input):
    """Convert an input string containing UTF-16 to Unicode surrogate pairs"""
    for multibyte_char in PY3_MULTIBYTE_RE.findall(input):
        high, low = to_surrogate_pair(multibyte_char)
        input = input.replace(
            "\\u%s" % multibyte_char,
            "\\u%04x\\u%04x" % (high, low))
        input = input.replace(
            "\\U%s" % multibyte_char,
            "\\u%04x\\u%04x" % (high, low))
    return input


def fix_unicode(input):
    if sys.version_info[0] >= 3:
        return to_py3_compatible(input)
    else:
        return to_py2_compatible(input)
