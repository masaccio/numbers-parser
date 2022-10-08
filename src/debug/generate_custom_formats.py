import math

from enum import Enum
from uuid import uuid1

from numbers_parser import Document
from numbers_parser.constants import DOCUMENT_ID
from numbers_parser.generated import TSKArchives_pb2 as TSKArchives
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages


class FormatType(Enum):
    INT_DEFAULT = "I#"
    INT_ZEROS = "I0"
    INT_SPACES = "I-"
    DEC_DEFAULT = "D#"
    DEC_ZEROS = "D0"
    DEC_SPACES = "D-"


def add_format_to_table_data(table, format_uuid, format_id):
    base_data_store = table._model.objects[table._table_id].base_data_store
    format_table_id = base_data_store.format_table.identifier
    format_table = table._model.objects[format_table_id]
    format_table.entries.append(
        TSTArchives.TableDataList.ListEntry(
            key=format_id,
            refcount=1,
            format=TSKArchives.FormatStructArchive(
                format_type=270,
                custom_uid=TSPMessages.UUID(
                    lower=format_uuid.lower, upper=format_uuid.upper
                ),
            ),
        )
    )

    format_table_pre_bnc_id = base_data_store.format_table_pre_bnc.identifier
    format_table_pre_bnc = table._model.objects[format_table_pre_bnc_id]
    format_table_pre_bnc.entries.append(
        TSTArchives.TableDataList.ListEntry(
            key=format_id,
            refcount=1,
            format=TSKArchives.FormatStructArchive(
                format_type=270,
                custom_uid=TSPMessages.UUID(
                    lower=format_uuid.lower, upper=format_uuid.upper
                ),
            ),
        )
    )


def next_format_key(table):
    base_data_store = table._model.objects[table._table_id].base_data_store
    format_table_id = base_data_store.format_table.identifier
    format_table = table._model.objects[format_table_id]
    last_key = 0
    for entry in format_table.entries:
        if entry.key > last_key:
            last_key = entry.key
    format_table.nextListID = last_key + 1
    return format_table.nextListID


def custom_format(
    integer_format=FormatType.INT_DEFAULT,
    decimal_format=FormatType.DEC_DEFAULT,
    num_integers=0,
    num_decimals=0,
    show_thousands_separator=False,
) -> object:
    format_name = "Format_"
    format_name += integer_format.value + "_"
    format_name += "Int" + str(num_integers) + "_"
    format_name += decimal_format.value + "_"
    format_name += "Dec" + str(num_decimals) + "_"
    format_name += "Sep" if show_thousands_separator else "NoSep"

    if integer_format.value[1] == "-" or num_integers == 0:
        format_string = ""
    else:
        format_string = integer_format.value[1] * num_integers
    if num_decimals > 0:
        if decimal_format.value[1] == "-":
            format_string += "." + "0" * num_decimals
        else:
            format_string += "." + decimal_format.value[1] * num_decimals

    min_integer_width = 0

    num_nonspace_decimal_digits = 0

    num_nonspace_integer_digits = 0

    custom_format = TSKArchives.CustomFormatArchive(
        name=format_name,
        format_type_pre_bnc=270,
        format_type=270,
        default_format=TSKArchives.FormatStructArchive(
            contains_integer_token=False,
            custom_format_string=format_string,
            decimal_width=num_decimals,
            format_type=270,
            fraction_accuracy=0xFFFFFFFF,
            # is_complex=False,
            min_integer_width=min_integer_width,
            num_hash_decimal_digits=0,
            num_nonspace_decimal_digits=num_nonspace_decimal_digits,
            num_nonspace_integer_digits=num_nonspace_integer_digits,
            requires_fraction_replacement=False,
            scale_factor=1.0,
            show_thousands_separator=show_thousands_separator,
            total_num_decimal_digits=num_decimals,
            use_accounting_style=False,
        ),
    )
    hex_uuid = uuid1().hex
    uuid = TSPMessages.UUID(
        lower=int(hex_uuid[0:16], 16), upper=int(hex_uuid[16:32], 16)
    )
    return (uuid, custom_format)


def text_reference(
    value,
    integer_format=FormatType.INT_DEFAULT,
    decimal_format=FormatType.DEC_DEFAULT,
    num_integers=0,
    num_decimals=0,
    show_thousands_separator=False,
):
    integer = int(value) if num_decimals > 0 else round(value)
    decimal = math.modf(value)[0]
    if integer_format.value[1] == "0":
        if show_thousands_separator and integer != 0:
            num_integers += int(math.floor(math.log10(integer)) / 3)
            formatted_value = f"{integer:0{num_integers},}"
        else:
            formatted_value = f"{integer:0{num_integers}}"
    else:
        if show_thousands_separator and integer != 0:
            num_integers += int(math.floor(math.log10(integer)) / 3)
            formatted_value = f"{integer:{num_integers},}"
        else:
            formatted_value = str(integer)

    if num_decimals > 0:
        decimal = round(decimal, num_decimals)
        if decimal_format.value[1] == "0":
            formatted_value += "." + f"{decimal:.{num_decimals}f}"[2:]
        elif decimal_format.value[1] == "-":
            formatted_value += str(decimal).ljust(num_decimals)
        else:
            formatted_value += "." + str(decimal)[2:]
    return formatted_value


def write_formula_key(table, row_num, col_num, id):
    table.cell(row_num, col_num)._storage.formula_id = id


def write_format_key(table, row_num, col_num, id):
    table.cell(row_num, col_num)._storage.suggest_id = 1
    table.cell(row_num, col_num)._storage.num_format_id = id


doc = Document("tests/data/custom-format-stress-template.numbers")
table = doc.sheets[0].tables[0]
table.col_width(0, 200)
table.col_width(1, 30)
table.col_width(2, 30)
table.col_width(3, 35)
table.col_width(4, 35)
table.col_width(5, 40)
table.col_width(6, 50)
table.col_width(7, 100)
table.col_width(8, 130)

custom_format_list_id = doc._model.objects[
    DOCUMENT_ID
].super.custom_format_list.identifier
custom_format_list = doc._model.objects[custom_format_list_id]

test_formula_key = table.cell(1, 2)._storage.formula_id
test_format_id = table.cell(1, 2)._storage.num_format_id

row_num = 2
for integer_format in [
    FormatType.INT_DEFAULT,
    FormatType.INT_ZEROS,
    FormatType.INT_SPACES,
]:
    for decimal_format in [
        FormatType.DEC_DEFAULT,
        FormatType.DEC_ZEROS,
        FormatType.DEC_SPACES,
    ]:
        for num_integers in range(0, 10):
            for num_decimals in range(0, 10):
                if num_integers == 0 and num_decimals == 0:
                    continue
                for show_thousands_separator in [False, True]:
                    (uuid, format) = custom_format(
                        integer_format=integer_format,
                        decimal_format=decimal_format,
                        num_integers=num_integers,
                        num_decimals=num_decimals,
                        show_thousands_separator=show_thousands_separator,
                    )
                    custom_format_list.custom_formats.append(format)
                    custom_format_list.uuids.append(uuid)
                    format_id = next_format_key(table)
                    add_format_to_table_data(
                        table, custom_format_list.uuids[-1], format_id
                    )

                    for value in [
                        0.23,
                        2.34,
                        23.45,
                        234.56,
                        2345.67,
                    ]:
                        table.write(row_num, 0, format.name)
                        table.write(row_num, 1, integer_format.value)
                        table.write(row_num, 2, decimal_format.value)
                        table.write(row_num, 3, num_integers)
                        table.write(row_num, 4, num_decimals)
                        table.write(row_num, 5, show_thousands_separator)
                        table.write(row_num, 6, value)
                        table.write(row_num, 7, value)
                        write_format_key(table, row_num, 7, format_id)
                        write_formula_key(table, row_num, 7, test_formula_key)
                        table.write(
                            row_num,
                            8,
                            text_reference(
                                value,
                                integer_format=integer_format,
                                decimal_format=decimal_format,
                                num_integers=num_integers,
                                num_decimals=num_decimals,
                                show_thousands_separator=show_thousands_separator,
                            ),
                        )
                        row_num += 1

doc.save("tests/data/custom-format-stress.numbers")
