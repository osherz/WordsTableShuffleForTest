from docx.api import Document
import random

FILENAME = 'תמצית הקורס.docx'
TEST_FILENAME = 'test.docx'


def main():
    doc = Document(docx=FILENAME)
    for table in doc.tables:
        replace_random_paragraphs_from_table(table)
    doc.save(TEST_FILENAME)


def replace_random_paragraphs_from_table(table, paragraph_replace_precent=0.5, replace_with='0'):
    """
    Replace random paragraph from table with alternative text.

    :param table:
    :param replace_with:
    :return:
    """
    MAX_NUMBER_TO_RANDOM = 6
    # To not include the header
    for row in table.rows[1:]:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                if random.randrange(0, MAX_NUMBER_TO_RANDOM) > MAX_NUMBER_TO_RANDOM * paragraph_replace_precent:
                    paragraph.text = replace_with


if __name__ == '__main__':
    main()
