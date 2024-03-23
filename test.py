from numbers_parser import Document

doc = Document("/Users/jon/Downloads/test1.numbers")
for row in range(doc.default_table.num_rows):
    doc.default_table.write(row, 1, "TEST")
doc.save("/Users/jon/Downloads/test2.numbers")
