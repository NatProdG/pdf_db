from docx import Document
from faker import Faker
import time
import random
import os


def main():
    template_file_path = 'templates/test2.docx'
    output_file_path = 'output/'

    variables = {
        "${EMPLOEE_NAME}": get_name,
        "${EMPLOEE_PHONE}": get_phone_number,
        "${START_DATE}": get_date,
        "${ADDRESS}": get_address,
        "${PRIX}": get_prix,
        "${ITEM}": get_item,
        "${QUANTITY}": get_quantity,
        "${COMPANY}": get_company
    }

    for j in range(0, 20):
        template_document = Document(template_file_path)

        for variable_key, variable_value in variables.items():
            for paragraph in template_document.paragraphs:
                replace_text_in_paragraph(paragraph, variable_key, variable_value)

            for table in template_document.tables:
                for col in table.columns:
                    for cell in col.cells:
                        for paragraph in cell.paragraphs:
                            replace_text_in_paragraph(paragraph, variable_key, variable_value)
        template_document.save(output_file_path + template_file_path.split('/')[1].split('.')[0] + str(j + 1) + ".docx")


def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs

        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value())


def get_phone_number():
    nb = '+33'
    for i in range(1, 10):
        nb = nb + str(random.randint(0, 9))
    return nb


def get_prix():
    price = random.randrange(10, 1000, 10)
    return str(price)


def get_item():
    return random.choice(list(open('construction.txt')))


def get_quantity():
    return random.choice(['1', '2', '5', '10', '15', '25', '50'])


def get_company():
    fake = Faker()
    return fake.company()


def get_date():
    fake = Faker()
    return fake.date()


def get_address():
    fake = Faker()
    Faker.seed(random.randint(0, 9999))
    return fake.address()


def get_name():
    fake = Faker()
    Faker.seed(random.randint(0, 9999))
    return fake.name()


if __name__ == '__main__':
    main()
