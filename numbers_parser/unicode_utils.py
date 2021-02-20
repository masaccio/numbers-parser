"""
Tools to handle yaml Unicode incompatibilities between Python 2 and 3.

tl;dr: when writing Yaml with emojis in it from Python 2 and reading from
Python 3 (or vice versa),  are incompatibilities with the unicode handling
that result in runtime exceptions. It's Bad.

This file aims to manually do some pre-parsing on the Yaml strings before
decoding to ensure that whichever version of Python is doing the decoding
can successfully parse the string. Since this library now only supports
Python 3, we only need to support reading Python 2-written strings in
Python 3, not vice versa.

Russell Cottrell's wonderful Surrogate Pair Calculator
(http://www.russellcottrell.com/greek/utilities/SurrogatePairCalculator.htm)
has been invaluable in understanding how to do this translation manually,
including the fact that Unicode surrogate pairs are only valid if the first
element is between 0xD800 and 0xDBFF, and the second element is between
0xDC00 and 0xDFFF.
"""

import re

PY2_SURROGATE_PAIR_RE = re.compile(
    r'\\u([Dd][89a-bA-B][0-9a-fA-F]{2})\\u([Dd][c-fC-F][0-9a-fA-F]{2})'
)


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
        input = input.replace("\\u%s\\u%s" % (high, low), "\\U%08x" % character)
        input = input.replace("\\U%s\\U%s" % (high, low), "\\U%08x" % character)
    return input


def fix_unicode(input):
    return to_py3_compatible(input)
