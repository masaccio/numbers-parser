from keynote_parser import codec


SIMPLE_FILENAME = './tests/data/simple-oneslide.iwa'
MULTILINE_FILENAME = './tests/data/multiline-oneslide.iwa'


def roundtrip(filename):
    with open(filename, 'r') as f:
        test_data = f.read()
    file = codec.IWAFile.from_buffer(test_data)
    assert file is not None
    assert file.to_buffer() == test_data


def test_iwa_simple_roundtrip():
    roundtrip(SIMPLE_FILENAME)


def test_iwa_multiline_roundtrip():
    roundtrip(MULTILINE_FILENAME)
