import yaml
from keynote_parser import codec
from keynote_parser.unicode_utils import fix_unicode


SIMPLE_FILENAME = './tests/data/simple-oneslide.iwa'
MULTILINE_FILENAME = './tests/data/multiline-oneslide.iwa'
MULTICHUNK_FILENAME = './tests/data/multi-chunk.iwa'
EMOJI_FILENAME = './tests/data/emoji-oneslide.iwa'
MESSAGE_TYPE_ZERO_FILENAME = './tests/data/message-type-zero.iwa'
EMOJI_FILENAME_PY2_YAML = './tests/data/emoji-oneslide.py2.yaml'
EMOJI_FILENAME_PY3_YAML = './tests/data/emoji-oneslide.py3.yaml'


def roundtrip(filename):
    with open(filename, 'rb') as f:
        test_data = f.read()
    file = codec.IWAFile.from_buffer(test_data, filename)
    assert file is not None
    for chunk in file.chunks:
        for archive in chunk.archives:
            assert codec.IWAArchiveSegment.from_buffer(archive.to_buffer())[0] == archive
        assert codec.IWACompressedChunk.from_buffer(chunk.to_buffer())[0] == chunk
    assert file.to_dict()
    assert file.to_buffer() == test_data


def test_iwa_simple_roundtrip():
    roundtrip(SIMPLE_FILENAME)


def test_iwa_multiline_roundtrip():
    roundtrip(MULTILINE_FILENAME)


def test_iwa_emoji_roundtrip():
    roundtrip(EMOJI_FILENAME)


def test_iwa_message_type_zero_roundtrip():
    roundtrip(MESSAGE_TYPE_ZERO_FILENAME)


def test_yaml_parse_py2_emoji():
    with open(EMOJI_FILENAME_PY2_YAML, 'rb') as handle:
        file = codec.IWAFile.from_dict(yaml.load(fix_unicode(handle.read().decode('utf-8'))))
        assert file is not None


def test_yaml_parse_py3_emoji():
    with open(EMOJI_FILENAME_PY3_YAML, 'rb') as handle:
        file = codec.IWAFile.from_dict(yaml.load(fix_unicode(handle.read().decode('utf-8'))))
        assert file is not None


def test_iwa_multichunk_roundtrip():
    with open(MULTICHUNK_FILENAME, 'rb') as f:
        test_data = f.read()
    file = codec.IWAFile.from_buffer(test_data, MULTICHUNK_FILENAME)
    assert file is not None
    rt_as_dict = codec.IWAFile.from_buffer(file.to_buffer(), MULTICHUNK_FILENAME).to_dict()
    assert rt_as_dict == file.to_dict()
