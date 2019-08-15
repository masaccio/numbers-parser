from keynote_parser import codec
from keynote_parser.replacement import Replacement


MULTILINE_SURROGATE_FILENAME = './tests/data/multiline-surrogate.iwa'


def test_iwa_multiline_surrogate_replacement():
    with open(MULTILINE_SURROGATE_FILENAME, 'rb') as f:
        test_data = f.read()
    file = codec.IWAFile.from_buffer(test_data)
    data = file.to_dict()
    replacement = Replacement("\\$REPLACE_ME", "replaced!")
    replaced = replacement.perform_on(data)

    text_object = replaced['chunks'][0]['archives'][2]['objects'][0]
    replaced_character_indices = [
        entry['characterIndex']
        for entry in text_object['tableParaStyle']['entries']]

    assert replaced_character_indices == [0, 67, 72]
