import xlsxwriter


def random_data():

    workbook = xlsxwriter.Workbook('filename.xlsx')
    worksheet = workbook.add_worksheet()

    # creating different formats for data

    bold = workbook.add_format({'bold': 1})
    italic = workbook.add_format({'italic': 1})
    underline = workbook.add_format({'underline': 1})
    superscript = workbook.add_format({'font_script': 1})
    subscript = workbook.add_format({'font_script': 2})

    workbook.close()


if __name__ == "__main__":
    random_data()
