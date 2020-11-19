import xlsxwriter
import random
from faker import Faker

fake = Faker()


def random_data():

    workbook = xlsxwriter.Workbook('filename.xlsx')
    worksheet = workbook.add_worksheet()

    # creating different formats for data

    bold = workbook.add_format({'bold': 1})
    italic = workbook.add_format({'italic': 1})
    underline = workbook.add_format({'underline': 1})
    superscript = workbook.add_format({'font_script': 1})
    subscript = workbook.add_format({'font_script': 2})

    name_list = []
    for _ in range(random.randint(10, 100)):
        name_list.append(fake.name())

    worksheet.write_column('A1', name_list)
    workbook.close()


if __name__ == "__main__":
    random_data()
