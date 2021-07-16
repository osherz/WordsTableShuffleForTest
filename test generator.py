from docx.api import Document
import random

FILENAME = 'תמצית הקורס.docx'
TEST_FILENAME = 'test.docx'

doc = Document(docx=FILENAME)
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                if random.randrange(0, 6) > 3:
                    paragraph.text = '0'
doc.save(TEST_FILENAME)