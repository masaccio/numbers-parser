from datetime import datetime

from numbers_parser import Document

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
    5.0: [
        {"Description": "Airplane", "Category": "Transport", "Count": 5.0},
        {"Description": "Bus", "Category": "Transport", "Count": 5.0},
        {"Description": "Helicopter", "Category": "Transport", "Count": 5.0},
        {"Description": "Van", "Category": "Transport", "Count": 5.0},
    ],
    10.0: [
        {"Description": "Bear", "Category": "Animal", "Count": 10.0},
        {"Description": "Deer", "Category": "Animal", "Count": 10.0},
        {"Description": "Elephant", "Category": "Animal", "Count": 10.0},
        {"Description": "Lion", "Category": "Animal", "Count": 10.0},
        {"Description": "Scooter", "Category": "Transport", "Count": 10.0},
        {"Description": "Train", "Category": "Transport", "Count": 10.0},
    ],
    20.0: [
        {"Description": "Cat", "Category": "Animal", "Count": 20.0},
        {"Description": "Dog", "Category": "Animal", "Count": 20.0},
        {"Description": "Truck", "Category": "Transport", "Count": 20.0},
    ],
    30.0: [
        {"Description": "Car", "Category": "Transport", "Count": 30.0},
        {"Description": "Horse", "Category": "Animal", "Count": 30.0},
        {"Description": "Kiwi", "Category": "Fruit", "Count": 30.0},
        {"Description": "Mango", "Category": "Fruit", "Count": 30.0},
        {"Description": "Motorcycle", "Category": "Transport", "Count": 30.0},
        {"Description": "Parrot", "Category": "Animal", "Count": 30.0},
        {"Description": "Zebra", "Category": "Animal", "Count": 30.0},
    ],
    40.0: [
        {"Description": "Apple", "Category": "Fruit", "Count": 40.0},
        {"Description": "Banana", "Category": "Fruit", "Count": 40.0},
        {"Description": "Bicycle", "Category": "Transport", "Count": 40.0},
        {"Description": "Orange", "Category": "Fruit", "Count": 40.0},
        {"Description": "Peach", "Category": "Fruit", "Count": 40.0},
        {"Description": "Pineapple", "Category": "Fruit", "Count": 40.0},
        {"Description": "Watermelon", "Category": "Fruit", "Count": 40.0},
    ],
    120.0: [
        {"Description": "Grape", "Category": "Fruit", "Count": 120.0},
        {"Description": "Rabbit", "Category": "Animal", "Count": 120.0},
        {"Description": "Strawberry", "Category": "Fruit", "Count": 120.0},
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
    sheet = doc.sheets["Categories"]

    categories = sheet.tables["Uncategorised"].categorized_data()
    assert categories is None

    categories = sheet.tables["Categories"].categorized_data()
    assert categories == TEXT_CATEGORIES

    # categories = sheet.tables["Nested Categories"].categorized_data()
    # assert categories == NESTED_TEXT_CATEGORIES

    categories = sheet.tables["Number Categories"].categorized_data()
    assert categories == NUMBER_CATEGORIES

    # categories = sheet.tables["Date Categories"].categorized_data()
    # assert categories == DATE_CATEGORIES
