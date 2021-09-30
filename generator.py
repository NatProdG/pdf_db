from docx import Document
from faker import Faker
import datetime
import time
import random
import os

prices = []
quantities = []
date = ''
total = 0

def main():
    template_file_path = 'templates/test2.docx'
    output_file_path = 'output/'

    variables = {
        "${EMPLOEE_NAME}": get_name,
        "${EMPLOEE_PHONE}": get_phone_number,
        "${START_DATE}": get_start_date,
        "${END_DATE}": get_end_date,
        "${ADDRESS}": get_address,
        "${PRIX}": get_prix,
        "${ITEM}": get_item,
        "${QUANTITY}": get_quantity,
        "${INTER}": get_inter,
        "${TOTAL}": get_total,
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
    price = random.randrange(10, 150, 10)
    global prices
    prices.append(price)
    return str(price)


def get_item():
    return random.choice(list(open('construction.txt')))


def get_quantity():
    q = random.choice([1,2,5,10,15,20,25])
    global quantities
    quantities.append(q)
    return str(q)


def get_company():
    fake = Faker()
    return fake.company()


def get_start_date():
    fake = Faker()
    global date
    d = fake.date_between('-5y', 'today')
    date = d
    return str(d)


def get_end_date():
    global date
    delta = datetime.timedelta(days=random.randint(5,90))
    end_date = date + delta
    return str(end_date)


def get_address():
    fake = Faker('fr_FR')
    return fake.address()


def get_name():
    fake = Faker('fr_FR')
    return fake.name()


def get_inter():
    global prices
    global quantities
    res = prices.pop(0) * quantities.pop(0)
    global total
    total = total + res
    return str(res)


def get_total():
    global total
    res = total
    total = 0
    return str(res)


if __name__ == '__main__':
    main()
