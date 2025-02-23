import math
import re

from numbers_parser import Document
from numbers_parser.constants import DECIMAL128_BIAS


def _pack_decimal128(value: float) -> bytearray:
    buffer = bytearray(16)
    exp = math.floor(math.log10(math.e) * math.log(abs(value))) if value != 0.0 else 0
    exp += DECIMAL128_BIAS - 16
    mantissa = abs(int(value / math.pow(10, exp - DECIMAL128_BIAS)))
    buffer[15] |= exp >> 7
    buffer[14] |= (exp & 0x7F) << 1
    i = 0
    while mantissa >= 1:
        buffer[i] = mantissa & 0xFF
        i += 1
        mantissa = int(mantissa / 256)
    if value < 0:
        buffer[15] |= 0x80
    return buffer


def dump_number_encoding():
    doc = Document("tests/data/formula-decode-debug.numbers")
    table = doc.sheets[0].tables[0]

    print(",".join(["value", "node_number", "decimal_low", "decimal_high"]))  # noqa: FLY002
    for row in range(table.num_rows):
        cell = table.cell(row, 0)
        archive = cell._model._formulas.lookup_value(cell._table_id, cell._formula_id)
        node = archive.formula.AST_node_array.AST_node[0]

        value = str(cell.formatted_value)
        # bytes = _pack_decimal128(cell.value)
        exponent = (
            math.floor(math.log10(math.e) * math.log(abs(cell.value))) if cell.value != 0.0 else 0
        )
        if "E" in value:
            significand, exponent = value.split("E")
        else:
            significand = value
            exponent = 0
        num_dp = len(re.sub(r"0*$", "", str(significand).split(".")[1]))
        exponent = int(exponent) - num_dp
        decimal = cell.formatted_value
        decimal_low = int(float(significand) * 10**num_dp)
        decimal_high = ((DECIMAL128_BIAS * 2) + (2 * exponent)) << 48

        if (
            # decimal != node.AST_number_node_number
            decimal_low != node.AST_number_node_decimal_low
            or decimal_high != node.AST_number_node_decimal_high
        ):
            print(f"{decimal}\n{node.AST_number_node_number}\n-")
            print(f"{decimal_low}\n{node.AST_number_node_decimal_low}\n-")
            print(f"{hex(decimal_high)}\n{hex(node.AST_number_node_decimal_high)}\n--------\n")


def dump_colon_tract_bits():
    doc = Document("tests/data/create-formulas.numbers")
    table = doc.sheets["Main Sheet"].tables["Formula Tests"]

    all_fields = [
        "formula",
        "begin_column_is_absolute",
        "begin_row_is_absolute",
        "end_column_is_absolute",
        "end_row_is_absolute",
        "relative_column_range_begin",
        "relative_column_range_end",
        "relative_row_range_begin",
        "relative_row_range_end",
        "absolute_column_range_begin",
        "absolute_column_range_end",
        "absolute_row_range_begin",
        "absolute_row_range_end",
        "frozen_begin_column_is_absolute",
        "frozen_begin_row_is_absolute",
        "frozen_end_column_is_absolute",
        "frozen_end_row_is_absolute",
    ]
    print(",".join([x.replace("_", " ") for x in all_fields]))

    for row in range(25, table.num_rows):
        cell = table.cell(row, 0)
        archive = cell._model._formulas.lookup_value(cell._table_id, cell._formula_id)
        node = archive.formula.AST_node_array.AST_node[0]
        colon_tract = node.AST_colon_tract

        fields = dict.fromkeys(all_fields, "")
        fields["formula"] = cell.formula

        for axis in ["column", "row"]:
            for end in ["begin", "end"]:
                fields[f"{end}_{axis}_is_absolute"] = getattr(
                    node.AST_sticky_bits,
                    f"{end}_{axis}_is_absolute",
                )
                if node.HasField("AST_frozen_sticky_bits"):
                    fields[f"frozen_{end}_{axis}_is_absolute"] = getattr(
                        node.AST_frozen_sticky_bits,
                        f"{end}_{axis}_is_absolute",
                    )
                else:
                    fields[f"frozen_{end}_{axis}_is_absolute"] = ""

                for wise in ["relative", "absolute"]:
                    attr = getattr(colon_tract, f"{wise}_{axis}")
                    if len(attr) > 0 and attr[0].HasField(f"range_{end}"):
                        fields[f"{wise}_{axis}_range_{end}"] = getattr(attr[0], f"range_{end}")

        print(",".join([str(fields[x]) for x in all_fields]))


dump_number_encoding()
