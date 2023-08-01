import argparse

from numbers_parser import Document, xl_range
from numbers_parser.model import rgb

MAX_DEPTH = 10


def dump_strokes_for_layer(objects, model, layer_ids, side):
    for layer_id in layer_ids:
        stroke_layer = objects[layer_id.identifier]
        for stroke_run in stroke_layer.stroke_runs:
            width = round(stroke_run.stroke.width, 2)
            color = rgb(stroke_run.stroke.color)
            style = model.stroke_type(stroke_run)
            length = stroke_run.length
            order = stroke_run.order
            border_value = (
                f"width={width}, color={color}, style={style}, length={length}, order={order}"
            )

            if side in ["top", "bottom"]:
                start_row = stroke_layer.row_column_index
                start_column = stroke_run.origin
                for col_num in range(start_column, start_column + stroke_run.length):
                    ref = xl_range(start_row, col_num, start_row, col_num)
                    print(f"{side}: {ref}[{start_row},{col_num}]: {border_value}")
            else:
                start_row = stroke_run.origin
                start_column = stroke_layer.row_column_index
                for row_num in range(start_row, start_row + stroke_run.length):
                    ref = xl_range(row_num, start_column, row_num, start_column)
                    print(f"{side}: {ref}[{row_num},{start_column}]: {border_value}")


def dump_strokesfor_doc(doc):
    objects = doc._model.objects
    for sheet in doc.sheets:
        for table in sheet.tables:
            print(f"==== {sheet.name} - {table.name}")
            table_obj = objects[table._table_id]
            sidecar_obj = objects[table_obj.stroke_sidecar.identifier]
            dump_strokes_for_layer(objects, doc._model, sidecar_obj.top_row_stroke_layers, "top")
            dump_strokes_for_layer(
                objects, doc._model, sidecar_obj.right_column_stroke_layers, "right"
            )
            dump_strokes_for_layer(
                objects, doc._model, sidecar_obj.bottom_row_stroke_layers, "bottom"
            )
            dump_strokes_for_layer(
                objects, doc._model, sidecar_obj.left_column_stroke_layers, "left"
            )


parser = argparse.ArgumentParser()
parser.add_argument(
    "numbers",
    nargs="*",
    metavar="numbers-filename",
    help="Numbers folders/files to dump",
)
args = parser.parse_args()

for filename in args.numbers:
    doc = Document(filename)
    dump_strokesfor_doc(doc)
