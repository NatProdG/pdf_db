from docx import Document
from docx2pdf import convert
from faker import Faker
import datetime
import time
import random
import os

address = ""
prices = []
quantities = []
date = ''
total = 0
taux = 0
heures = 0
heures_supp = 0
sal_brut = 0
charges = 0
school = ''


def main():
    template_file_path = 'templates/diplome.docx'
    output_file_path = 'output/'

    variables = {
        "${EMPLOEE_NAME}": get_name,
        "${EMPLOEE_PHONE}": get_phone_number,
        "${START_DATE}": get_start_date,
        "${END_DATE}": get_end_date,
        "${ADDRESS}": get_address,
        "${CITY_ADDRESS}": get_city_address,
        "${JOB}" : get_job,
        "${PRIX}": get_prix,
        "${ITEM}": get_item,
        "${QUANTITY}": get_quantity,
        "${INTER}": get_inter,
        "${TOTAL}": get_total,
        "${COMPANY}": get_company,
        "${HEURES}": get_heures,
        "${HEURES_SUPP}": get_heures_supp,
        "${TAUX}": get_taux_horaire,
        "${SAL_HEURES}": get_sal_heures,
        "${SAL_HEURES_SUPP}": get_sal_heures_supp,
        "${SALAIRE_BRUT}": get_brut_salary,
        "${CHARGES}": get_social_charge,
        "${SALAIRE_NET}": get_net_salary,
        "${SCHOOL}": get_school
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

def get_school():
    global school
    if school != '':
        temp = school
        school = ''
        return temp
    else:
        school = random.choice(list(open('ecole.txt')))
        return school


def get_quantity():
    q = random.choice([1, 2, 5, 10, 15, 20, 25])
    global quantities
    quantities.append(q)
    return str(q)


def get_company():
    fake = Faker()
    return fake.company()


def get_job():
    fake = Faker('fr_FR')
    return fake.job()


def get_start_date():
    fake = Faker()
    global date
    d = fake.date_between('-5y', 'today')
    date = d
    return str(d)


def get_end_date():
    global date
    delta = datetime.timedelta(days=random.randint(5, 90))
    end_date = date + delta
    return str(end_date)


def get_address():
    global address
    fake = Faker('fr_FR')
    add = fake.address()
    address = add
    return add


def get_city_address():
    global address
    c = address.split(' ')[-1]
    return c


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


def get_heures():
    global heures
    h = random.randint(120, 200)
    heures = h
    return str(h)


def get_heures_supp():
    global heures_supp
    h = random.choice([0, 0, 0, 0, 0, 5, 10])
    heures_supp = h
    return str(h)


def get_taux_horaire():
    global taux
    t = round(random.uniform(10.03, 15.0),2)
    taux = t
    return str(t)


def get_sal_heures():
    global heures
    global taux
    return str(heures*taux)


def get_sal_heures_supp():
    global heures_supp
    global taux
    return str(heures_supp*(taux*1.25))


def get_brut_salary():
    global sal_brut
    global heures
    global heures_supp
    global taux
    s = round((heures * taux + heures_supp * (taux * 1.25)),2)
    sal_brut = s
    return str(s)


def get_social_charge():
    global sal_brut
    global charges
    c = round(sal_brut * 0.22, 2)
    charges = c
    return str(c)


def get_net_salary():
    global sal_brut
    global charges
    s = round(sal_brut - charges,2)
    return str(s)


if __name__ == '__main__':
    main()
