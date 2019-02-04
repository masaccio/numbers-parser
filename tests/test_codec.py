from keynote_parser import codec


SIMPLE_FILENAME = './tests/data/simple-oneslide.iwa'
MULTILINE_FILENAME = './tests/data/multiline-oneslide.iwa'
MULTICHUNK_FILENAME = './tests/data/multi-chunk.iwa'


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


def test_iwa_multichunk_roundtrip():
    with open(MULTICHUNK_FILENAME, 'rb') as f:
        test_data = f.read()
    file = codec.IWAFile.from_buffer(test_data)
    assert file is not None
    rt_as_dict = codec.IWAFile.from_buffer(file.to_buffer()).to_dict()
    assert rt_as_dict == file.to_dict()
