from datetime import datetime

from numbers_parser import BoolCell, DateCell, Document, NumberCell, RichTextCell, TextCell

TEXT_CATEGORIES = {
    "Transport": [
        ["Airplane", "Transport", 5.0],
        ["Bicycle", "Transport", 40.0],
        ["Bus", "Transport", 5.0],
        ["Car", "Transport", 30.0],
        ["Helicopter", "Transport", 5.0],
        ["Motorcycle", "Transport", 30.0],
        ["Scooter", "Transport", 10.0],
        ["Train", "Transport", 10.0],
        ["Truck", "Transport", 20.0],
        ["Van", "Transport", 5.0],
    ],
    "Fruit": [
        ["Apple", "Fruit", 40.0],
        ["Banana", "Fruit", 40.0],
        ["Grape", "Fruit", 120.0],
        ["Kiwi", "Fruit", 30.0],
        ["Mango", "Fruit", 30.0],
        ["Orange", "Fruit", 40.0],
        ["Peach", "Fruit", 40.0],
        ["Pineapple", "Fruit", 40.0],
        ["Strawberry", "Fruit", 120.0],
        ["Watermelon", "Fruit", 40.0],
    ],
    "Animal": [
        ["Bear", "Animal", 10.0],
        ["Cat", "Animal", 20.0],
        ["Deer", "Animal", 10.0],
        ["Dog", "Animal", 20.0],
        ["Elephant", "Animal", 10.0],
        ["Horse", "Animal", 30.0],
        ["Lion", "Animal", 10.0],
        ["Parrot", "Animal", 30.0],
        ["Rabbit", "Animal", 120.0],
        ["Zebra", "Animal", 30.0],
    ],
}

NESTED_BOOL_CATEGORIES = {
    False: {
        "Transport": [
            ["Airplane", False, "Transport", 5.0],
            ["Bicycle", False, "Transport", 40.0],
            ["Bus", False, "Transport", 5.0],
            ["Car", False, "Transport", 30.0],
            ["Helicopter", False, "Transport", 5.0],
            ["Motorcycle", False, "Transport", 30.0],
            ["Scooter", False, "Transport", 10.0],
            ["Train", False, "Transport", 10.0],
            ["Truck", False, "Transport", 20.0],
            ["Van", False, "Transport", 5.0],
        ],
        "Animal": [
            ["Bear", False, "Animal", 10.0],
            ["Cat", False, "Animal", 20.0],
            ["Dog", False, "Animal", 20.0],
            ["Elephant", False, "Animal", 10.0],
            ["Lion", False, "Animal", 10.0],
            ["Parrot", False, "Animal", 30.0],
        ],
    },
    True: {
        "Fruit": [
            ["Apple", True, "Fruit", 40.0],
            ["Banana", True, "Fruit", 40.0],
            ["Grape", True, "Fruit", 120.0],
            ["Kiwi", True, "Fruit", 30.0],
            ["Mango", True, "Fruit", 30.0],
            ["Orange", True, "Fruit", 40.0],
            ["Peach", True, "Fruit", 40.0],
            ["Pineapple", True, "Fruit", 40.0],
            ["Strawberry", True, "Fruit", 120.0],
            ["Watermelon", True, "Fruit", 40.0],
        ],
        "Animal": [
            ["Deer", True, "Animal", 10.0],
            ["Horse", True, "Animal", 30.0],
            ["Rabbit", True, "Animal", 120.0],
            ["Zebra", True, "Animal", 30.0],
        ],
    },
}

NUMBER_CATEGORIES = {
    5.0: [
        ["Airplane", "Transport", 5.0],
        ["Bus", "Transport", 5.0],
        ["Helicopter", "Transport", 5.0],
        ["Van", "Transport", 5.0],
    ],
    10.0: [
        ["Bear", "Animal", 10.0],
        ["Deer", "Animal", 10.0],
        ["Elephant", "Animal", 10.0],
        ["Lion", "Animal", 10.0],
        ["Scooter", "Transport", 10.0],
        ["Train", "Transport", 10.0],
    ],
    20.0: [
        ["Cat", "Animal", 20.0],
        ["Dog", "Animal", 20.0],
        ["Truck", "Transport", 20.0],
    ],
    30.0: [
        ["Car", "Transport", 30.0],
        ["Horse", "Animal", 30.0],
        ["Kiwi", "Fruit", 30.0],
        ["Mango", "Fruit", 30.0],
        ["Motorcycle", "Transport", 30.0],
        ["Parrot", "Animal", 30.0],
        ["Zebra", "Animal", 30.0],
    ],
    40.0: [
        ["Apple", "Fruit", 40.0],
        ["Banana", "Fruit", 40.0],
        ["Bicycle", "Transport", 40.0],
        ["Orange", "Fruit", 40.0],
        ["Peach", "Fruit", 40.0],
        ["Pineapple", "Fruit", 40.0],
        ["Watermelon", "Fruit", 40.0],
    ],
    120.0: [
        ["Grape", "Fruit", 120.0],
        ["Rabbit", "Animal", 120.0],
        ["Strawberry", "Fruit", 120.0],
    ],
}

NESTED_DATE_CATEGORIES = {
    "Fruit": {
        2010: [
            ["Banana", datetime(2010, 11, 5), "Fruit", 40.0],
            ["Grape", datetime(2010, 9, 22), "Fruit", 120.0],
            ["Watermelon", datetime(2010, 6, 27), "Fruit", 40.0],
        ],
        2011: [
            ["Pineapple", datetime(2011, 10, 23), "Fruit", 40.0],
        ],
        2012: [
            ["Apple", datetime(2012, 4, 15), "Fruit", 40.0],
            ["Kiwi", datetime(2012, 7, 13), "Fruit", 30.0],
            ["Orange", datetime(2012, 10, 5), "Fruit", 40.0],
        ],
        2013: [
            ["Mango", datetime(2013, 2, 1), "Fruit", 30.0],
        ],
        2014: [
            ["Strawberry", datetime(2014, 8, 16), "Fruit", 120.0],
        ],
        2015: [
            ["Peach", datetime(2015, 2, 28), "Fruit", 40.0],
        ],
    },
    "Animal": {
        2010: [
            ["Deer", datetime(2010, 5, 18), "Animal", 10.0],
            ["Rabbit", datetime(2010, 2, 14), "Animal", 120.0],
        ],
        2011: [
            ["Cat", datetime(2011, 8, 29), "Animal", 20.0],
        ],
        2013: [
            ["Elephant", datetime(2013, 11, 11), "Animal", 10.0],
            ["Zebra", datetime(2013, 10, 14), "Animal", 30.0],
        ],
        2014: [
            ["Bear", datetime(2014, 9, 19), "Animal", 10.0],
            ["Dog", datetime(2014, 3, 4), "Animal", 20.0],
            ["Lion", datetime(2014, 11, 30), "Animal", 10.0],
            ["Parrot", datetime(2014, 5, 9), "Animal", 30.0],
        ],
        2015: [
            ["Horse", datetime(2015, 4, 25), "Animal", 30.0],
        ],
    },
    "Transport": {
        2010: [
            ["Motorcycle", datetime(2010, 12, 17), "Transport", 30.0],
        ],
        2011: [
            ["Bicycle", datetime(2011, 3, 21), "Transport", 40.0],
            ["Helicopter", datetime(2011, 6, 8), "Transport", 5.0],
            ["Train", datetime(2011, 1, 30), "Transport", 10.0],
        ],
        2012: [
            ["Bus", datetime(2012, 12, 3), "Transport", 5.0],
            ["Truck", datetime(2012, 3, 9), "Transport", 20.0],
        ],
        2013: [
            ["Airplane", datetime(2013, 7, 26), "Transport", 5.0],
            ["Scooter", datetime(2013, 5, 6), "Transport", 10.0],
        ],
        2015: [
            ["Car", datetime(2015, 1, 10), "Transport", 30.0],
            ["Van", datetime(2015, 6, 5), "Transport", 5.0],
        ],
    },
}

MAXIMALLY_NESTED_CATEGORIES = (
    {
        "Transport": {
            False: {
                5: {
                    2013: [
                        ["Airplane", False, datetime(2013, 7, 26), "Transport", 5.0],
                    ],
                    2012: [
                        ["Bus", False, datetime(2012, 12, 3), "Transport", 5.0],
                    ],
                    2011: [["Helicopter", False, datetime(2011, 6, 8), "Transport", 5.0]],
                    2015: [["Van", False, datetime(2015, 6, 5), "Transport", 5.0]],
                },
                40: {2011: [["Bicycle", False, datetime(2011, 3, 21), "Transport", 40.0]]},
                30: {
                    2015: [["Car", False, datetime(2015, 1, 10), "Transport", 30.0]],
                    2010: [["Motorcycle", False, datetime(2010, 12, 17), "Transport", 30.0]],
                },
                10: {
                    2013: [["Scooter", False, datetime(2013, 5, 6), "Transport", 10.0]],
                    2011: [["Train", False, datetime(2011, 1, 30), "Transport", 10.0]],
                },
                20: {2012: [["Truck", False, datetime(2012, 3, 9), "Transport", 20.0]]},
            },
        },
        "Fruit": {
            True: {
                40: {
                    2012: [
                        ["Apple", True, datetime(2012, 4, 15), "Fruit", 40.0],
                        ["Orange", True, datetime(2012, 10, 5), "Fruit", 40.0],
                    ],
                    2010: [
                        ["Banana", True, datetime(2010, 11, 5), "Fruit", 40.0],
                        ["Watermelon", True, datetime(2010, 6, 27), "Fruit", 40.0],
                    ],
                    2015: [["Peach", True, datetime(2015, 2, 28), "Fruit", 40.0]],
                    2011: [["Pineapple", True, datetime(2011, 10, 23), "Fruit", 40.0]],
                },
                120: {
                    2010: [["Grape", True, datetime(2010, 9, 22), "Fruit", 120.0]],
                    2014: [
                        ["Strawberry", True, datetime(2014, 8, 16), "Fruit", 120.0],
                    ],
                },
                30: {
                    2012: [["Kiwi", True, datetime(2012, 7, 13), "Fruit", 30.0]],
                    2013: [["Mango", True, datetime(2013, 2, 1), "Fruit", 30.0]],
                },
            },
        },
        "Animal": {
            False: {
                10: {
                    2014: [
                        ["Bear", False, datetime(2014, 9, 19), "Animal", 10.0],
                        ["Lion", False, datetime(2014, 11, 30), "Animal", 10.0],
                    ],
                    2013: [["Elephant", False, datetime(2013, 11, 11), "Animal", 10.0]],
                },
                20: {
                    2011: [["Cat", False, datetime(2011, 8, 29), "Animal", 20.0]],
                    2014: [["Dog", False, datetime(2014, 3, 4), "Animal", 20.0]],
                },
                30: {2014: [["Parrot", False, datetime(2014, 5, 9), "Animal", 30.0]]},
            },
            True: {
                10: {2010: [["Deer", True, datetime(2010, 5, 18), "Animal", 10.0]]},
                30: {
                    2015: [["Horse", True, datetime(2015, 4, 25), "Animal", 30.0]],
                    2013: [["Zebra", True, datetime(2013, 10, 14), "Animal", 30.0]],
                },
                120: {
                    2010: [
                        ["Rabbit", True, datetime(2010, 2, 14), "Animal", 120.0],
                    ],
                },
            },
        },
    },
)


def test_group_lookups():
    doc = Document("tests/data/test-categories.numbers")

    ungrouped_table = doc.sheets[0].tables["Uncategorized"]
    grouped_table = doc.sheets[0].tables["Categories"]

    assert ungrouped_table.categorized_data() is None

    assert ungrouped_table.cell("A5").value == "Bear"
    assert ungrouped_table.cell("E5").value == 10
    assert grouped_table.cell("A5").value == "Car"
    assert grouped_table.cell("C5").value == 30

    data = [x[0] for x in grouped_table.iter_rows(min_row=3, max_row=5, values_only=True)]
    assert data == ["Bus", "Car", "Helicopter"]

    data = grouped_table.iter_cols(
        min_col=0,
        max_col=0,
        min_row=3,
        max_row=5,
        values_only=True,
    )
    assert list(data) == [("Bus", "Car", "Helicopter")]

    table = doc.sheets[0].tables["Maximal Nesting"]
    assert table.cell("D8").value == "Transport"
    assert table.cell("B9").value == False  # noqa: E712
    assert table.cell("C10").value == datetime(2011, 1, 30)


def ref_data_to_types(item):
    if isinstance(item, dict):
        return {k: ref_data_to_types(v) for k, v in item.items()}

    if isinstance(item, list):
        return [ref_data_to_types(v) for v in item]

    return type(item)


def data_to_types(item):  # noqa: PLR0911
    if isinstance(item, dict):
        return {k: data_to_types(v) for k, v in item.items()}

    if isinstance(item, list):
        return [data_to_types(x) for x in item]

    if isinstance(item, DateCell):
        return datetime
    if isinstance(item, BoolCell):
        return bool
    if isinstance(item, TextCell):
        return str
    if isinstance(item, NumberCell):
        return float
    if isinstance(item, (TextCell, RichTextCell)):
        return str
    return None


def data_to_values(item):
    if isinstance(item, dict):
        return {k: data_to_values(v) for k, v in item.items()}

    if isinstance(item, list):
        return [data_to_values(v) for v in item]

    return item.value


def test_category_trees():
    doc = Document("tests/data/test-categories.numbers")
    sheet = doc.sheets["Categories"]

    table = doc.sheets[0].tables["Date Categories"]
    assert table.cell("A2").value == "Banana"

    categories = sheet.tables["Uncategorized"].categorized_data()
    assert categories is None

    categories = sheet.tables["Categories"].categorized_data()
    assert data_to_values(categories) == TEXT_CATEGORIES
    assert data_to_types(categories) == ref_data_to_types(TEXT_CATEGORIES)

    categories = sheet.tables["Nested Categories"].categorized_data()
    assert data_to_values(categories) == NESTED_BOOL_CATEGORIES
    assert data_to_types(categories) == ref_data_to_types(NESTED_BOOL_CATEGORIES)

    categories = sheet.tables["Number Categories"].categorized_data()
    assert data_to_values(categories) == NUMBER_CATEGORIES
    assert data_to_types(categories) == ref_data_to_types(NUMBER_CATEGORIES)

    categories = sheet.tables["Date Categories"].categorized_data()
    assert data_to_values(categories) == NESTED_DATE_CATEGORIES
    assert data_to_types(categories) == ref_data_to_types(NESTED_DATE_CATEGORIES)

    categories = sheet.tables["Maximal Nesting"].categorized_data()
    assert data_to_values(categories) == MAXIMALLY_NESTED_CATEGORIES
    assert data_to_types(categories) == ref_data_to_types(MAXIMALLY_NESTED_CATEGORIES)
