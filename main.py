import xlsxwriter
import random
from faker import Faker

fake = Faker()


def random_data():

    rand = random.randint(10, 100)

    workbook = xlsxwriter.Workbook('filename.xlsx')
    worksheet = workbook.add_worksheet()

    # creating different formats for data
    bold = workbook.add_format({'bold': 1})
    italic = workbook.add_format({'italic': 1})
    underline = workbook.add_format({'underline': 1})
    superscript = workbook.add_format({'font_script': 1})
    subscript = workbook.add_format({'font_script': 2})
    red = workbook.add_format({'font_color': 'red'})

    format_list = [bold, italic, underline, superscript, subscript, red]  # list of created formats

    # format for header
    header_format = workbook.add_format()
    header_format.set_font_size(18)
    header_format.set_bold()
    header_format.set_bg_color('green')
    header_format.set_align('center')
    header_format.set_align('vcenter')

    header_list = ['No.', 'Name', 'Address']

    for i in range(rand):
        form = random.choice(format_list)
        name = fake.name()
        worksheet.write(0 + i, 0, name, form)

    workbook.close()


if __name__ == "__main__":
    random_data()
