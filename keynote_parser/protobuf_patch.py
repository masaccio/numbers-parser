# Protocol Buffers - Google's data interchange format
# Copyright 2008 Google Inc.  All rights reserved.
# https://developers.google.com/protocol-buffers/
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are
# met:
#
#     * Redistributions of source code must retain the above copyright
# notice, this list of conditions and the following disclaimer.
#     * Redistributions in binary form must reproduce the above
# copyright notice, this list of conditions and the following disclaimer
# in the documentation and/or other materials provided with the
# distribution.
#     * Neither the name of Google Inc. nor the names of its
# contributors may be used to endorse or promote products derived from
# this software without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
# "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
# LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
# A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
# OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
# SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
# LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
# DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
# THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
# (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

# Until https://github.com/protocolbuffers/protobuf/pull/5609 gets merged,
# this file provides an alternative implementation of some Protobuf methods
# so that we can serialize/deserialize Protobuf files properly as dictionaries.

from google.protobuf import descriptor

from google.protobuf.descriptor import FieldDescriptor

import google.protobuf.json_format
from google.protobuf.json_format import _VALID_EXTENSION_NAME, \
    ParseError, \
    _ConvertScalarFieldValue, \
    _IsMapEntry

CPPTYPE_MESSAGE = FieldDescriptor.CPPTYPE_MESSAGE


class _Parser(google.protobuf.json_format._Parser):
    """JSON format parser for protocol message."""

    def __init__(self, ignore_unknown_fields, descriptor_pool):
        self.ignore_unknown_fields = ignore_unknown_fields
        self.descriptor_pool = descriptor_pool

    def _ConvertFieldValuePair(self, js, message):
        """Convert field value pairs into regular message.

        Args:
          js: A JSON object to convert the field value pairs.
          message: A regular protocol message to record the data.

        Raises:
          ParseError: In case of problems converting.
        """
        names = []
        message_descriptor = message.DESCRIPTOR
        fields_by_json_name = dict((f.json_name, f)
                                   for f in message_descriptor.fields)
        for name in js:
            try:
                field = fields_by_json_name.get(name, None)
                if not field:
                    field = message_descriptor.fields_by_name.get(name, None)
                if not field and _VALID_EXTENSION_NAME.match(name):
                    if not message_descriptor.is_extendable:
                        raise ParseError(
                            'Message type {0} does not have extensions'.format(
                                message_descriptor.full_name))
                    identifier = name[1:-1]  # strip [] brackets
                    identifier = '.'.join(identifier.split('.')[:-1])
                    # pylint: disable=protected-access
                    field = message.Extensions._FindExtensionByName(identifier)
                    # pylint: enable=protected-access
                if not field:
                    if self.ignore_unknown_fields:
                        continue
                    raise ParseError(
                        ('Message type "{0}" has no field named "{1}".\n'
                         ' Available Fields(except extensions): {2}').format(
                            message_descriptor.full_name, name,
                            [f.json_name for f in message_descriptor.fields]))
                if name in names:
                    raise ParseError(
                        'Message type "{0}" should not have multiple '
                        '"{1}" fields.'.format(
                            message.DESCRIPTOR.full_name, name))
                names.append(name)
                # Check no other oneof field is parsed.
                if field.containing_oneof is not None:
                    oneof_name = field.containing_oneof.name
                    if oneof_name in names:
                        raise ParseError(
                            'Message type "{0}" should not have multiple '
                            '"{1}" oneof fields.'.format(
                                message.DESCRIPTOR.full_name, oneof_name))
                    names.append(oneof_name)

                value = js[name]
                if value is None:
                    if (field.cpp_type == CPPTYPE_MESSAGE and
                        field.message_type.full_name ==
                            'google.protobuf.Value'):
                        sub_message = getattr(message, field.name)
                        sub_message.null_value = 0
                    else:
                        message.ClearField(field.name)
                    continue

                # Parse field value.
                if _IsMapEntry(field):
                    message.ClearField(field.name)
                    self._ConvertMapFieldValue(value, message, field)
                elif field.label == descriptor.FieldDescriptor.LABEL_REPEATED:
                    message.ClearField(field.name)
                    if not isinstance(value, list):
                        raise ParseError(
                            'repeated field {0} must be in [] which is '
                            '{1}.'.format(name, value))
                    if field.cpp_type == CPPTYPE_MESSAGE:
                        # Repeated message field.
                        for item in value:
                            sub_message = getattr(message, field.name).add()
                            # None is a null_value in Value.
                            if (item is None and
                                    sub_message.DESCRIPTOR.full_name !=
                                    'google.protobuf.Value'):
                                raise ParseError(
                                    'null is not allowed to be used as an '
                                    'element in a repeated field.')
                            self.ConvertMessage(item, sub_message)
                    else:
                        # Repeated scalar field.
                        for item in value:
                            if item is None:
                                raise ParseError(
                                    'null is not allowed to be used as an '
                                    'element in a repeated field.')
                            getattr(message, field.name).append(
                                _ConvertScalarFieldValue(item, field))
                elif field.cpp_type == CPPTYPE_MESSAGE:
                    if field.is_extension:
                        sub_message = message.Extensions[field]
                    else:
                        sub_message = getattr(message, field.name)
                    sub_message.SetInParent()
                    self.ConvertMessage(value, sub_message)
                else:
                    if field.is_extension:
                        message.Extensions[
                            field] = _ConvertScalarFieldValue(value, field)
                    else:
                        setattr(message, field.name,
                                _ConvertScalarFieldValue(value, field))
            except ParseError as e:
                if field and field.containing_oneof is None:
                    raise ParseError(
                        'Failed to parse {0} field: {1}'.format(name, e))
                else:
                    raise ParseError(str(e))
            except ValueError as e:
                raise ParseError(
                    'Failed to parse {0} field: {1}.'.format(name, e))
            except TypeError as e:
                raise ParseError(
                    'Failed to parse {0} field: {1}.'.format(name, e))


def ParseDict(js_dict,
              message,
              ignore_unknown_fields=False,
              descriptor_pool=None):
    """Parses a JSON dictionary representation into a message.

    Args:
      js_dict: Dict representation of a JSON message.
      message: A protocol buffer message to merge into.
      ignore_unknown_fields: If True, do not raise errors for unknown fields.
      descriptor_pool: A Descriptor Pool for resolving types. If None use the
        default.

    Returns:
      The same message passed as argument.
    """
    parser = _Parser(ignore_unknown_fields, descriptor_pool)
    parser.ConvertMessage(js_dict, message)
    return message
