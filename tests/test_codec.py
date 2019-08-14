import yaml
from keynote_parser import codec
from keynote_parser.unicode_utils import fix_unicode


SIMPLE_FILENAME = './tests/data/simple-oneslide.iwa'
MULTILINE_FILENAME = './tests/data/multiline-oneslide.iwa'
MULTILINE_SURROGATE_FILENAME = './tests/data/multiline-surrogate.iwa'
MULTICHUNK_FILENAME = './tests/data/multi-chunk.iwa'
EMOJI_FILENAME = './tests/data/emoji-oneslide.iwa'
EMOJI_FILENAME_PY2_YAML = './tests/data/emoji-oneslide.py2.yaml'
EMOJI_FILENAME_PY3_YAML = './tests/data/emoji-oneslide.py3.yaml'


def roundtrip(filename):
    with open(filename, 'rb') as f:
        test_data = f.read()
    file = codec.IWAFile.from_buffer(test_data)
    assert file is not None
    assert file.to_buffer() == test_data


def test_iwa_simple_roundtrip():
    roundtrip(SIMPLE_FILENAME)


def test_iwa_multiline_roundtrip():
    roundtrip(MULTILINE_FILENAME)


def test_iwa_multiline_surrogate_roundtrip():
    roundtrip(MULTILINE_SURROGATE_FILENAME)


def test_iwa_emoji_roundtrip():
    roundtrip(EMOJI_FILENAME)


def test_yaml_parse_py2_emoji():
    with open(EMOJI_FILENAME_PY2_YAML, 'rb') as handle:
        file = codec.IWAFile.from_dict(yaml.load(
            fix_unicode(handle.read().decode('utf-8'))))
        assert file is not None


def test_yaml_parse_py3_emoji():
    with open(EMOJI_FILENAME_PY3_YAML, 'rb') as handle:
        file = codec.IWAFile.from_dict(yaml.load(
            fix_unicode(handle.read().decode('utf-8'))))
        assert file is not None


def test_iwa_multichunk_roundtrip():
    with open(MULTICHUNK_FILENAME, 'rb') as f:
        test_data = f.read()
    file = codec.IWAFile.from_buffer(test_data)
    assert file is not None
    rt_as_dict = codec.IWAFile.from_buffer(file.to_buffer()).to_dict()
    assert rt_as_dict == file.to_dict()
