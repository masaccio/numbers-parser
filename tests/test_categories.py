from datetime import datetime

import pytest

from numbers_parser import CellValueType, Document

TEXT_CATEGORIES = {
    "Transport": [
        {"Description": "Airplane", "Category": "Transport", "Count": 5},
        {"Description": "Bicycle", "Category": "Transport", "Count": 40},
        {"Description": "Bus", "Category": "Transport", "Count": 15},
        {"Description": "Car", "Category": "Transport", "Count": 30},
        {"Description": "Helicopter", "Category": "Transport", "Count": 3},
        {"Description": "Motorcycle", "Category": "Transport", "Count": 18},
        {"Description": "Scooter", "Category": "Transport", "Count": 25},
        {"Description": "Train", "Category": "Transport", "Count": 10},
        {"Description": "Truck", "Category": "Transport", "Count": 20},
    ],
    "Fruit": [
        {"Description": "Apple", "Category": "Fruit", "Count": 120},
        {"Description": "Banana", "Category": "Fruit", "Count": 95},
        {"Description": "Grape", "Category": "Fruit", "Count": 100},
        {"Description": "Kiwi", "Category": "Fruit", "Count": 45},
        {"Description": "Mango", "Category": "Fruit", "Count": 60},
        {"Description": "Orange", "Category": "Fruit", "Count": 70},
        {"Description": "Peach", "Category": "Fruit", "Count": 55},
        {"Description": "Pineapple", "Category": "Fruit", "Count": 50},
        {"Description": "Strawberry", "Category": "Fruit", "Count": 85},
        {"Description": "Watermelon", "Category": "Fruit", "Count": 35},
    ],
    "Animal": [
        {"Description": "Bear", "Category": "Animal", "Count": 14},
        {"Description": "Cat", "Category": "Animal", "Count": 68},
        {"Description": "Deer", "Category": "Animal", "Count": 20},
        {"Description": "Dog", "Category": "Animal", "Count": 75},
        {"Description": "Elephant", "Category": "Animal", "Count": 12},
        {"Description": "Horse", "Category": "Animal", "Count": 30},
        {"Description": "Lion", "Category": "Animal", "Count": 9},
        {"Description": "Parrot", "Category": "Animal", "Count": 22},
        {"Description": "Rabbit", "Category": "Animal", "Count": 40},
        {"Description": "Van", "Category": "Transport", "Count": 8},
        {"Description": "Zebra", "Category": "Animal", "Count": 16},
    ],
}

NESTED_TEXT_CATEGORIES = {
    False: {
        "Transport": [
            {"Description": "Airplane", "Edible": "No", "Category": "Transport", "Count": 5},
            {"Description": "Bicycle", "Edible": "No", "Category": "Transport", "Count": 40},
            {"Description": "Bus", "Edible": "No", "Category": "Transport", "Count": 15},
            {"Description": "Car", "Edible": "No", "Category": "Transport", "Count": 30},
            {"Description": "Helicopter", "Edible": "No", "Category": "Transport", "Count": 3},
            {"Description": "Motorcycle", "Edible": "No", "Category": "Transport", "Count": 18},
            {"Description": "Scooter", "Edible": "No", "Category": "Transport", "Count": 25},
            {"Description": "Train", "Edible": "No", "Category": "Transport", "Count": 10},
            {"Description": "Truck", "Edible": "No", "Category": "Transport", "Count": 20},
            {"Description": "Van", "Edible": "No", "Category": "Transport", "Count": 8},
        ],
        "Animal": [
            {"Description": "Bear", "Edible": "No", "Category": "Animal", "Count": 14},
            {"Description": "Cat", "Edible": "No", "Category": "Animal", "Count": 68},
            {"Description": "Dog", "Edible": "No", "Category": "Animal", "Count": 75},
            {"Description": "Elephant", "Edible": "No", "Category": "Animal", "Count": 12},
            {"Description": "Lion", "Edible": "No", "Category": "Animal", "Count": 9},
            {"Description": "Parrot", "Edible": "No", "Category": "Animal", "Count": 22},
        ],
    },
    True: {
        "Fruit": [
            {"Description": "Apple", "Edible": "Yes", "Category": "Fruit", "Count": 120},
            {"Description": "Banana", "Edible": "Yes", "Category": "Fruit", "Count": 95},
            {"Description": "Mango", "Edible": "Yes", "Category": "Fruit", "Count": 60},
            {"Description": "Strawberry", "Edible": "Yes", "Category": "Fruit", "Count": 85},
            {"Description": "Pineapple", "Edible": "Yes", "Category": "Fruit", "Count": 50},
            {"Description": "Kiwi", "Edible": "Yes", "Category": "Fruit", "Count": 45},
            {"Description": "Peach", "Edible": "Yes", "Category": "Fruit", "Count": 55},
            {"Description": "Grape", "Edible": "Yes", "Category": "Fruit", "Count": 100},
            {"Description": "Orange", "Edible": "Yes", "Category": "Fruit", "Count": 70},
            {"Description": "Watermelon", "Edible": "Yes", "Category": "Fruit", "Count": 35},
        ],
        "Animal": [
            {"Description": "Horse", "Edible": "Yes", "Category": "Animal", "Count": 30},
            {"Description": "Zebra", "Edible": "Yes", "Category": "Animal", "Count": 16},
            {"Description": "Rabbit", "Edible": "Yes", "Category": "Animal", "Count": 40},
            {"Description": "Deer", "Edible": "Yes", "Category": "Animal", "Count": 20},
        ],
    },
}

NUMBER_CATEGORIES = {
    5: [
        {"Description": "Airplane", "Category": "Transport", "Count": 5},
        {"Description": "Helicopter", "Category": "Transport", "Count": 5},
        {"Description": "Van", "Category": "Transport", "Count": 5},
        {"Description": "Bus", "Category": "Transport", "Count": 5},
    ],
    10: [
        {"Description": "Bear", "Category": "Animal", "Count": 10},
        {"Description": "Deer", "Category": "Animal", "Count": 10},
        {"Description": "Elephant", "Category": "Animal", "Count": 10},
        {"Description": "Train", "Category": "Transport", "Count": 10},
        {"Description": "Lion", "Category": "Animal", "Count": 10},
        {"Description": "Scooter", "Category": "Transport", "Count": 10},
    ],
    20: [
        {"Description": "Cat", "Category": "Animal", "Count": 20},
        {"Description": "Dog", "Category": "Animal", "Count": 20},
        {"Description": "Truck", "Category": "Transport", "Count": 20},
    ],
    30: [
        {"Description": "Car", "Category": "Transport", "Count": 30},
        {"Description": "Horse", "Category": "Animal", "Count": 30},
        {"Description": "Kiwi", "Category": "Fruit", "Count": 30},
        {"Description": "Mango", "Category": "Fruit", "Count": 30},
        {"Description": "Motorcycle", "Category": "Transport", "Count": 30},
        {"Description": "Parrot", "Category": "Animal", "Count": 30},
        {"Description": "Zebra", "Category": "Animal", "Count": 30},
    ],
    40: [
        {"Description": "Watermelon", "Category": "Fruit", "Count": 40},
        {"Description": "Apple", "Category": "Fruit", "Count": 40},
        {"Description": "Banana", "Category": "Fruit", "Count": 40},
        {"Description": "Bicycle", "Category": "Transport", "Count": 40},
        {"Description": "Orange", "Category": "Fruit", "Count": 40},
        {"Description": "Pineapple", "Category": "Fruit", "Count": 40},
        {"Description": "Peach", "Category": "Fruit", "Count": 40},
    ],
    120: [
        {"Description": "Rabbit", "Category": "Animal", "Count": 120},
        {"Description": "Grape", "Category": "Fruit", "Count": 120},
        {"Description": "Strawberry", "Category": "Fruit", "Count": 120},
    ],
}

DATE_CATEGORIES = {
    "Fruit": {
        2010: [
            {
                "Description": "Banana",
                "Date Seen": datetime(2010, 11, 5),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40,
            },
            {
                "Description": "Grape",
                "Date Seen": datetime(2010, 9, 22),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 120,
            },
            {
                "Description": "Watermelon",
                "Date Seen": datetime(2010, 6, 27),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40,
            },
        ],
        2011: [
            {
                "Description": "Pineapple",
                "Date Seen": datetime(2011, 10, 23),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40,
            },
        ],
        2012: [
            {
                "Description": "Apple",
                "Date Seen": datetime(2012, 4, 15),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40,
            },
            {
                "Description": "Kiwi",
                "Date Seen": datetime(2012, 7, 13),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 30,
            },
            {
                "Description": "Orange",
                "Date Seen": datetime(2012, 10, 5),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40,
            },
        ],
        2013: [
            {
                "Description": "Mango",
                "Date Seen": datetime(2013, 2, 1),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 30,
            },
        ],
        2014: [
            {
                "Description": "Strawberry",
                "Date Seen": datetime(2014, 8, 16),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 120,
            },
        ],
        2015: [
            {
                "Description": "Peach",
                "Date Seen": datetime(2015, 2, 28),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40,
            },
        ],
    },
    "Animal": {
        2010: [
            {
                "Description": "Deer",
                "Date Seen": datetime(2010, 5, 18),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 10,
            },
            {
                "Description": "Rabbit",
                "Date Seen": datetime(2010, 2, 14),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 120,
            },
        ],
        2011: [
            {
                "Description": "Cat",
                "Date Seen": datetime(2011, 8, 29),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 20,
            },
        ],
        2013: [
            {
                "Description": "Elephant",
                "Date Seen": datetime(2013, 11, 11),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 10,
            },
            {
                "Description": "Zebra",
                "Date Seen": datetime(2013, 10, 14),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 30,
            },
        ],
        2014: [
            {
                "Description": "Bear",
                "Date Seen": datetime(2014, 9, 19),  # noqa: DTZ001  # noqa: DTZ001
                "Category": "Animal",
                "Count": 10,
            },
            {
                "Description": "Dog",
                "Date Seen": datetime(2014, 3, 4),
                "Category": "Animal",
                "Count": 20,
            },
            {
                "Description": "Lion",
                "Date Seen": datetime(2014, 11, 30),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 10,
            },
            {
                "Description": "Parrot",
                "Date Seen": datetime(2014, 5, 9),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 30,
            },
        ],
        2015: [
            {
                "Description": "Horse",
                "Date Seen": datetime(2015, 4, 25),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 30,
            },
        ],
    },
    "Transport": {
        2010: [
            {
                "Description": "Motorcycle",
                "Date Seen": datetime(2010, 12, 17),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 30,
            },
        ],
        2011: [
            {
                "Description": "Bicycle",
                "Date Seen": datetime(2011, 3, 21),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 40,
            },
            {
                "Description": "Helicopter",
                "Date Seen": datetime(2011, 6, 8),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 5,
            },
            {
                "Description": "Train",
                "Date Seen": datetime(2011, 1, 30),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 10,
            },
        ],
        2012: [
            {
                "Description": "Bus",
                "Date Seen": datetime(2012, 12, 3),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 5,
            },
            {
                "Description": "Truck",
                "Date Seen": datetime(2012, 3, 9),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 20,
            },
        ],
        2013: [
            {
                "Description": "Airplane",
                "Date Seen": datetime(2013, 7, 26),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 5,
            },
            {
                "Description": "Scooter",
                "Date Seen": datetime(2013, 5, 6),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 10,
            },
        ],
        2015: [
            {
                "Description": "Car",
                "Date Seen": datetime(2015, 1, 10),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 30,
            },
            {
                "Description": "Van",
                "Date Seen": datetime(2015, 6, 5),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 5,
            },
        ],
    },
}


def test_categories():
    doc = Document("tests/data/test-categories.numbers")
    grouped_table = doc.sheets["Categories"].tables["Categorised"]

    objects = doc._model.objects
    category_owner_id = objects[grouped_table._table_id].category_owner.identifier
    category_archive_id = objects[category_owner_id].group_by[0].identifier
    category_archive = objects[category_archive_id]

    cell_value_types = [
        x.group_cell_value.cell_value_type for x in category_archive.group_node_root.child
    ]
    if not all(x == CellValueType.STRING_TYPE for x in cell_value_types):
        pytest.fail("Not all cell value types are strings")

    cell_value_map = {
        x.group_cell_value.string_value.value: x.row_lookup_uids.entries
        for x in category_archive.group_node_root.child
    }
    row = 1
    row_map = {}
    for category, row_offsets in cell_value_map.items():
        for obj in row_offsets:
            if obj.range_end > 0:
                range_begin = obj.range_begin
                range_end = obj.range_end
            else:
                range_begin = obj.range_begin
                range_end = range_begin
            for r in range(range_begin, range_end + 1):
                row_map[row] = r
                row += 1

        row += 1

    categories = doc.sheets["Text"].tables["Uncategorised"].categorized_data()
    assert categories is None

    categories = doc.sheets["Text"].tables["Categories"].categorized_data()
    assert categories == TEXT_CATEGORIES

    categories = doc.sheets["Text"].tables["Nested Categories"].categorized_data()
    assert categories == NESTED_TEXT_CATEGORIES

    categories = doc.sheets["Text"].tables["Number Categories"].categorized_data()
    assert categories == NUMBER_CATEGORIES

    categories = doc.sheets["Text"].tables["Date Categories"].categorized_data()
    assert categories == DATE_CATEGORIES
