from keynote_parser.unicode_utils import fix_unicode, to_py3_compatible


def test_non_surrogate_pair():
    assert fix_unicode('\u2716\uFE0F') == '\u2716\uFE0F'
    assert to_py3_compatible('\u2716\uFE0F') == '\u2716\uFE0F'


def test_surrogate_pair():
    assert to_py3_compatible(r'\ud83c\udde8\ud83c\udde6') == r'\U0001f1e8\U0001f1e6'


def test_basic_multilingual_plane():
    srpska = r'\u0441\u0440\u043f\u0441\u043a\u0430'
    assert to_py3_compatible(srpska) == srpska


def test_german_example():
    deutsch_nicht_ersatzpaar = br'\uFFFC\u201C."'.decode('utf-8')
    assert fix_unicode(deutsch_nicht_ersatzpaar) == deutsch_nicht_ersatzpaar
