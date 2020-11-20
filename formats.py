import xlsxwriter
import random
from faker import Faker

fake = Faker()


def random_data():  # this function will create the xlsx file with random data in the table

    rand = random.randint(10, 100)

    # setting up the file
    workbook = xlsxwriter.Workbook('random_data.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A1:B1', 30)
    worksheet.set_column('C1:C1', 60)

    # creating different formats for data
    bold = workbook.add_format()
    bold.set_align('center')
    bold.set_align('vcenter')
    bold.set_bold()
    bold.set_bg_color('blue')
    bold.set_border(1)

    italic = workbook.add_format()
    italic.set_align('center')
    italic.set_align('vcenter')
    italic.set_italic()
    italic.set_bg_color('magenta')
    italic.set_border(1)

    underline = workbook.add_format()
    underline.set_align('center')
    underline.set_align('vcenter')
    underline.set_underline()
    underline.set_bg_color('purple')
    underline.set_border(1)

    superscript = workbook.add_format()
    superscript.set_align('center')
    superscript.set_align('vcenter')
    superscript.set_font_script(1)
    superscript.set_bg_color('orange')
    superscript.set_border(1)

    subscript = workbook.add_format()
    subscript.set_align('center')
    subscript.set_align('vcenter')
    subscript.set_font_script(2)
    subscript.set_bg_color('yellow')
    subscript.set_border(1)

    red = workbook.add_format()
    red.set_align('center')
    red.set_align('vcenter')
    red.set_font_color('red')
    red.set_bg_color('navy')
    red.set_border(1)

    # list of created formats
    format_list = [bold, italic, underline, superscript, subscript, red]

    # format for header
    header_format = workbook.add_format()
    header_format.set_font_size(18)
    header_format.set_bold()
    header_format.set_bg_color('green')
    header_format.set_align('center')
    header_format.set_align('vcenter')
    header_format.set_border(1)

    # format for no.
    number_format = workbook.add_format()
    number_format.set_font_size(18)
    number_format.set_bold()
    number_format.set_bg_color('cyan')
    number_format.set_align('center')
    number_format.set_align('vcenter')
    number_format.set_border(1)

    # writing the header of the table
    header_list = ['No.', 'Name', 'Address']
    worksheet.write_row('A1', header_list, header_format)

    # loops to write no., name, address
    for i in range(rand):
        worksheet.write(1 + i, 0, i+1, number_format)

    for i in range(rand):
        form = random.choice(format_list)
        name = fake.name()
        worksheet.write(1 + i, 1, name, form)
    for i in range(rand):
        form = random.choice(format_list)
        address = fake.address()
        worksheet.write(1 + i, 2, address, form)

    worksheet.write_url('E5', 'https://www.youtube.com/watch?v=dQw4w9WgXcQ', string='Click Me!')

    workbook.close()
