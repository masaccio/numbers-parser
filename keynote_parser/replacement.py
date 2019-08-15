from builtins import zip
from builtins import object
import re
import json


def parse_json(json_path):
    replacements = []
    with open(json_path) as f:
        data = json.load(f)
        for replacement_dict in data['replacements']:
            try:
                replacements.append(Replacement(**{
                    x: replacement_dict[x]
                    for x in ('find', 'replace', 'key_path')
                    if x in replacement_dict
                }))
            except Exception as e:
                raise ValueError(
                    "Failed to parse %s (data: '%s'): %s" % (
                        json_path, replacement_dict, e))
    return replacements


def merge_two_dicts(x, y):
    z = x.copy()
    z.update(y)
    return z


class Replacement(object):
    DEFAULT_KEY_PATH = "chunks.[].archives.[].objects.[].text.[]"

    def __init__(self, find, replace, key_path=DEFAULT_KEY_PATH):
        self.find = find
        self.replace = replace
        if not isinstance(key_path, list):
            self.key_path = key_path.split('.')
        else:
            self.key_path = key_path

    def __repr__(self):
        return "s/%s/%s/g" % (self.find, self.replace)

    def correct_multiline_replacement(self, _dict):
        """
        If dealing with a multiline text block, Keynote does some bookkeeping
        with the tableParaStyle property to figure out which paragraph styles
        apply to which separate paragraphs. This method corrects for that and
        allows us to do multiline replacements.

        Fun fact: without this bookkeeping, one of these tableParaStyle indices
        might point beyond the end of the text, which causes Keynote to make
        the text box 2^16 points tall, eventually forcing it to crash.
        """
        text = _dict['text'][0]

        new_offsets = [0]

        surrogate_pair_correction = 0
        for i, c in enumerate(text):
            if c == '\n':
                new_offsets.append(i + 1 + surrogate_pair_correction)
            if ord(c) > 0xFFFF:
                surrogate_pair_correction += 1

        entries = _dict['tableParaStyle']['entries']
        if len(entries) != len(new_offsets):
            raise NotImplementedError(
                "New line count doesn't match old line count in data: %s",
                text)
        for para_entry, offset in zip(entries, new_offsets):
            para_entry['characterIndex'] = offset
        return _dict

    def correct_charstyle_replacement(self, data, key_path, depth, on_replace):
        """
        If dealing with text that contains changing styles, Keynote does
        even more bookkeeping with the tableCharStyle property to figure
        out which character styles to apply to which ranges. This method
        corrects for that and allows us to keep consistent styles on
        replaced text.

        TODO: Throw an error or print a warning if the text being
        replaced spans multiple style blocks.
        """
        new_start = 0
        text = data['text'][0]
        if 'tableCharStyle' not in data \
                or len(data['tableCharStyle']['entries']) == 1:
            old_value = data[key_path[0]]
            new_value = self.perform_on(old_value, depth + 1, on_replace)
            return merge_two_dicts(data, {key_path[0]: new_value})
        char_style_entries = data['tableCharStyle']['entries']
        parts = []
        new_indices = []
        for start, end in zip(char_style_entries, char_style_entries[1:]):
            start_index = start['characterIndex']
            end_index = end['characterIndex']
            chunk = text[start_index:end_index]
            chunk = re.sub(self.find, self.replace, chunk)
            parts.append(chunk)
            new_end = new_start + len(chunk)
            new_indices.append(new_start)
            new_start = new_end
        new_indices.append(new_indices[-1] + len(parts[-1]))
        parts.append(text[char_style_entries[-1]['characterIndex']:])
        data['text'][0] = ''.join(parts)
        for new_start, entry in zip(new_indices, char_style_entries):
            entry['characterIndex'] = new_start
        return data

    def perform_on(self, data, depth=0, on_replace=None):
        key_path = self.key_path[depth:]
        if not key_path:
            new_value = re.sub(self.find, self.replace, data)
            if new_value != data and on_replace:
                on_replace(self, data, new_value)
            return new_value
        if key_path[0] == "[]":
            return [
                self.perform_on(obj, depth + 1, on_replace)
                for obj in data
            ]
        if key_path[0] in data:
            if key_path[0] == 'text':
                output = self.correct_charstyle_replacement(
                    data, key_path, depth, on_replace)
                output = self.correct_multiline_replacement(output)
            else:
                old_value = data[key_path[0]]
                new_value = self.perform_on(old_value, depth + 1, on_replace)
                output = merge_two_dicts(data, {key_path[0]: new_value})
            return output
        else:
            return data

    def should_replace(self, data, depth=0):
        key_path = self.key_path[depth:]
        if not key_path:
            return re.search(self.find, data) is not None
        elif key_path[0] == "[]":
            return any([
                self.should_replace(obj, depth + 1) for obj in data
            ])
        elif hasattr(data, key_path[0]):
            return self.should_replace(
                getattr(data, key_path[0]), depth + 1)
