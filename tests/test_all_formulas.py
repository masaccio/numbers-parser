import pytest

from numbers_parser import Document


@pytest.mark.experimental
def test_table_functions():
    doc = Document("tests/data/test-all-forumulas.numbers")
    for sheet in doc.sheets():
        if sheet.name.endswith("*") or sheet.name.endswith("Data"):
            continue
        print("====", sheet.name)
        tables = sheet.tables()
        for i, row in enumerate(tables["Tests"].rows()):
            if i == 0:
                continue
            formula_text = row[0].value
            formula_value = row[2].value
            if row[3].value:
                if not row[1].has_formula:
                    print(f"#{i}: missing formula")
                elif row[1].formula != formula_text:
                    print(
                        f"#{i}: incorrect formula ({row[1].formula} vs {formula_text}"
                    )
                elif row[1].value != formula_value:
                    print(f"#{i}: incorrect value ({row[1].value} vs {formula_value}")
