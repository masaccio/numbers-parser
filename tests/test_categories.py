from datetime import datetime

from numbers_parser import Document
from numbers_parser.numbers_uuid import NumbersUUID

TEXT_CATEGORIES = {
    "Transport": [
        {"Description": "Airplane", "Category": "Transport", "Count": 5.0},
        {"Description": "Bicycle", "Category": "Transport", "Count": 40.0},
        {"Description": "Bus", "Category": "Transport", "Count": 5.0},
        {"Description": "Car", "Category": "Transport", "Count": 30.0},
        {"Description": "Helicopter", "Category": "Transport", "Count": 5.0},
        {"Description": "Motorcycle", "Category": "Transport", "Count": 30.0},
        {"Description": "Scooter", "Category": "Transport", "Count": 10.0},
        {"Description": "Train", "Category": "Transport", "Count": 10.0},
        {"Description": "Truck", "Category": "Transport", "Count": 20.0},
        {"Description": "Van", "Category": "Transport", "Count": 5.0},
    ],
    "Fruit": [
        {"Description": "Apple", "Category": "Fruit", "Count": 40.0},
        {"Description": "Banana", "Category": "Fruit", "Count": 40.0},
        {"Description": "Grape", "Category": "Fruit", "Count": 120.0},
        {"Description": "Kiwi", "Category": "Fruit", "Count": 30.0},
        {"Description": "Mango", "Category": "Fruit", "Count": 30.0},
        {"Description": "Orange", "Category": "Fruit", "Count": 40.0},
        {"Description": "Peach", "Category": "Fruit", "Count": 40.0},
        {"Description": "Pineapple", "Category": "Fruit", "Count": 40.0},
        {"Description": "Strawberry", "Category": "Fruit", "Count": 120.0},
        {"Description": "Watermelon", "Category": "Fruit", "Count": 40.0},
    ],
    "Animal": [
        {"Description": "Bear", "Category": "Animal", "Count": 10.0},
        {"Description": "Cat", "Category": "Animal", "Count": 20.0},
        {"Description": "Deer", "Category": "Animal", "Count": 10.0},
        {"Description": "Dog", "Category": "Animal", "Count": 20.0},
        {"Description": "Elephant", "Category": "Animal", "Count": 10.0},
        {"Description": "Horse", "Category": "Animal", "Count": 30.0},
        {"Description": "Lion", "Category": "Animal", "Count": 10.0},
        {"Description": "Parrot", "Category": "Animal", "Count": 30.0},
        {"Description": "Rabbit", "Category": "Animal", "Count": 120.0},
        {"Description": "Zebra", "Category": "Animal", "Count": 30.0},
    ],
}

NESTED_TEXT_CATEGORIES = {
    False: {
        "Transport": [
            {"Description": "Airplane", "Edible": "No", "Category": "Transport", "Count": 5.0},
            {"Description": "Bicycle", "Edible": "No", "Category": "Transport", "Count": 40.0},
            {"Description": "Bus", "Edible": "No", "Category": "Transport", "Count": 15.0},
            {"Description": "Car", "Edible": "No", "Category": "Transport", "Count": 30.0},
            {"Description": "Helicopter", "Edible": "No", "Category": "Transport", "Count": 3.0},
            {"Description": "Motorcycle", "Edible": "No", "Category": "Transport", "Count": 18.0},
            {"Description": "Scooter", "Edible": "No", "Category": "Transport", "Count": 25.0},
            {"Description": "Train", "Edible": "No", "Category": "Transport", "Count": 10.0},
            {"Description": "Truck", "Edible": "No", "Category": "Transport", "Count": 20.0},
            {"Description": "Van", "Edible": "No", "Category": "Transport", "Count": 8.0},
        ],
        "Animal": [
            {"Description": "Bear", "Edible": "No", "Category": "Animal", "Count": 14.0},
            {"Description": "Cat", "Edible": "No", "Category": "Animal", "Count": 68.0},
            {"Description": "Dog", "Edible": "No", "Category": "Animal", "Count": 75.0},
            {"Description": "Elephant", "Edible": "No", "Category": "Animal", "Count": 12.0},
            {"Description": "Lion", "Edible": "No", "Category": "Animal", "Count": 9.0},
            {"Description": "Parrot", "Edible": "No", "Category": "Animal", "Count": 22.0},
        ],
    },
    True: {
        "Fruit": [
            {"Description": "Apple", "Edible": "Yes", "Category": "Fruit", "Count": 120.0},
            {"Description": "Banana", "Edible": "Yes", "Category": "Fruit", "Count": 95.0},
            {"Description": "Mango", "Edible": "Yes", "Category": "Fruit", "Count": 60.0},
            {"Description": "Strawberry", "Edible": "Yes", "Category": "Fruit", "Count": 85.0},
            {"Description": "Pineapple", "Edible": "Yes", "Category": "Fruit", "Count": 50.0},
            {"Description": "Kiwi", "Edible": "Yes", "Category": "Fruit", "Count": 45.0},
            {"Description": "Peach", "Edible": "Yes", "Category": "Fruit", "Count": 55.0},
            {"Description": "Grape", "Edible": "Yes", "Category": "Fruit", "Count": 100.0},
            {"Description": "Orange", "Edible": "Yes", "Category": "Fruit", "Count": 70.0},
            {"Description": "Watermelon", "Edible": "Yes", "Category": "Fruit", "Count": 35.0},
        ],
        "Animal": [
            {"Description": "Horse", "Edible": "Yes", "Category": "Animal", "Count": 30.0},
            {"Description": "Zebra", "Edible": "Yes", "Category": "Animal", "Count": 16.0},
            {"Description": "Rabbit", "Edible": "Yes", "Category": "Animal", "Count": 40.0},
            {"Description": "Deer", "Edible": "Yes", "Category": "Animal", "Count": 20.0},
        ],
    },
}

NUMBER_CATEGORIES = {
    5: [
        {"Description": "Airplane", "Category": "Transport", "Count": 5.0},
        {"Description": "Helicopter", "Category": "Transport", "Count": 5.0},
        {"Description": "Van", "Category": "Transport", "Count": 5.0},
        {"Description": "Bus", "Category": "Transport", "Count": 5.0},
    ],
    10: [
        {"Description": "Bear", "Category": "Animal", "Count": 10.0},
        {"Description": "Deer", "Category": "Animal", "Count": 10.0},
        {"Description": "Elephant", "Category": "Animal", "Count": 10.0},
        {"Description": "Train", "Category": "Transport", "Count": 10.0},
        {"Description": "Lion", "Category": "Animal", "Count": 10.0},
        {"Description": "Scooter", "Category": "Transport", "Count": 10.0},
    ],
    20: [
        {"Description": "Cat", "Category": "Animal", "Count": 20.0},
        {"Description": "Dog", "Category": "Animal", "Count": 20.0},
        {"Description": "Truck", "Category": "Transport", "Count": 20.0},
    ],
    30: [
        {"Description": "Car", "Category": "Transport", "Count": 30.0},
        {"Description": "Horse", "Category": "Animal", "Count": 30.0},
        {"Description": "Kiwi", "Category": "Fruit", "Count": 30.0},
        {"Description": "Mango", "Category": "Fruit", "Count": 30.0},
        {"Description": "Motorcycle", "Category": "Transport", "Count": 30.0},
        {"Description": "Parrot", "Category": "Animal", "Count": 30.0},
        {"Description": "Zebra", "Category": "Animal", "Count": 30.0},
    ],
    40: [
        {"Description": "Watermelon", "Category": "Fruit", "Count": 40.0},
        {"Description": "Apple", "Category": "Fruit", "Count": 40.0},
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
                "Count": 40.0,
            },
            {
                "Description": "Grape",
                "Date Seen": datetime(2010, 9, 22),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 120.0,
            },
            {
                "Description": "Watermelon",
                "Date Seen": datetime(2010, 6, 27),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40.0,
            },
        ],
        2011: [
            {
                "Description": "Pineapple",
                "Date Seen": datetime(2011, 10, 23),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40.0,
            },
        ],
        2012: [
            {
                "Description": "Apple",
                "Date Seen": datetime(2012, 4, 15),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40.0,
            },
            {
                "Description": "Kiwi",
                "Date Seen": datetime(2012, 7, 13),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 30.0,
            },
            {
                "Description": "Orange",
                "Date Seen": datetime(2012, 10, 5),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40.0,
            },
        ],
        2013: [
            {
                "Description": "Mango",
                "Date Seen": datetime(2013, 2, 1),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 30.0,
            },
        ],
        2014: [
            {
                "Description": "Strawberry",
                "Date Seen": datetime(2014, 8, 16),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 120.0,
            },
        ],
        2015: [
            {
                "Description": "Peach",
                "Date Seen": datetime(2015, 2, 28),  # noqa: DTZ001
                "Category": "Fruit",
                "Count": 40.0,
            },
        ],
    },
    "Animal": {
        2010: [
            {
                "Description": "Deer",
                "Date Seen": datetime(2010, 5, 18),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 10.0,
            },
            {
                "Description": "Rabbit",
                "Date Seen": datetime(2010, 2, 14),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 120.0,
            },
        ],
        2011: [
            {
                "Description": "Cat",
                "Date Seen": datetime(2011, 8, 29),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 20.0,
            },
        ],
        2013: [
            {
                "Description": "Elephant",
                "Date Seen": datetime(2013, 11, 11),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 10.0,
            },
            {
                "Description": "Zebra",
                "Date Seen": datetime(2013, 10, 14),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 30.0,
            },
        ],
        2014: [
            {
                "Description": "Bear",
                "Date Seen": datetime(2014, 9, 19),  # noqa: DTZ001  # noqa: DTZ001
                "Category": "Animal",
                "Count": 10.0,
            },
            {
                "Description": "Dog",
                "Date Seen": datetime(2014, 3, 4),
                "Category": "Animal",
                "Count": 20.0,
            },
            {
                "Description": "Lion",
                "Date Seen": datetime(2014, 11, 30),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 10.0,
            },
            {
                "Description": "Parrot",
                "Date Seen": datetime(2014, 5, 9),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 30.0,
            },
        ],
        2015: [
            {
                "Description": "Horse",
                "Date Seen": datetime(2015, 4, 25),  # noqa: DTZ001
                "Category": "Animal",
                "Count": 30.0,
            },
        ],
    },
    "Transport": {
        2010: [
            {
                "Description": "Motorcycle",
                "Date Seen": datetime(2010, 12, 17),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 30.0,
            },
        ],
        2011: [
            {
                "Description": "Bicycle",
                "Date Seen": datetime(2011, 3, 21),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 40.0,
            },
            {
                "Description": "Helicopter",
                "Date Seen": datetime(2011, 6, 8),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 5.0,
            },
            {
                "Description": "Train",
                "Date Seen": datetime(2011, 1, 30),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 10.0,
            },
        ],
        2012: [
            {
                "Description": "Bus",
                "Date Seen": datetime(2012, 12, 3),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 5.0,
            },
            {
                "Description": "Truck",
                "Date Seen": datetime(2012, 3, 9),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 20.0,
            },
        ],
        2013: [
            {
                "Description": "Airplane",
                "Date Seen": datetime(2013, 7, 26),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 5.0,
            },
            {
                "Description": "Scooter",
                "Date Seen": datetime(2013, 5, 6),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 10.0,
            },
        ],
        2015: [
            {
                "Description": "Car",
                "Date Seen": datetime(2015, 1, 10),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 30.0,
            },
            {
                "Description": "Van",
                "Date Seen": datetime(2015, 6, 5),  # noqa: DTZ001
                "Category": "Transport",
                "Count": 5.0,
            },
        ],
    },
}


def test_categories():
    doc = Document("tests/data/test-categories.numbers")
    grouped_table = doc.sheets["Categories"].tables["Categories"]

    objects = doc._model.objects
    category_owner_id = objects[grouped_table._table_id].category_owner.identifier
    category_archive_id = objects[category_owner_id].group_by[0].identifier
    category_archive = objects[category_archive_id]

    table_info = objects[doc._model.table_info_id(grouped_table._table_id)]
    category_order = objects[table_info.category_order.identifier]
    row_uid_map = objects[category_order.uid_map.identifier]
    sorted_row_uuids = [
        NumbersUUID(row_uid_map.sorted_row_uids[i]).hex for i in row_uid_map.row_uid_for_index
    ]
    uuid_to_row_index = {
        NumbersUUID(uuid).hex: i for i, uuid in enumerate(category_archive.row_uid_lookup.uuids)
    }
    group_uuid_map = {
        NumbersUUID(x.group_uid).hex: x.group_cell_value
        for x in category_archive.group_node_root.child
    }

    categories = {}
    key = None
    data = grouped_table.rows()
    assert uuid_to_row_index[sorted_row_uuids[0]] == 0
    header = [cell.value for cell in data[0]]
    for uuid in sorted_row_uuids[1:]:
        if uuid in group_uuid_map:
            key = group_uuid_map[uuid].string_value.value
            categories[key] = []
        else:
            row = uuid_to_row_index[uuid]
            categories[key].append({header[col]: cell.value for col, cell in enumerate(data[row])})

    # categories = doc.sheets["Text"].tables["Uncategorised"].categorized_data()
    # assert categories is None

    # categories = doc.sheets["Text"].tables["Categories"].categorized_data()
    assert categories == TEXT_CATEGORIES

    # categories = doc.sheets["Text"].tables["Nested Categories"].categorized_data()
    # assert categories == NESTED_TEXT_CATEGORIES

    # categories = doc.sheets["Text"].tables["Number Categories"].categorized_data()
    # assert categories == NUMBER_CATEGORIES

    # categories = doc.sheets["Text"].tables["Date Categories"].categorized_data()
    # assert categories == DATE_CATEGORIES
