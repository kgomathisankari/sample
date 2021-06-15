import openpyxl as xl


def appending_data_func(data, subject, chapter, duration, filename="Aravind_Studies.xlsx"):
    wb = xl.load_workbook(filename)

    sheet = wb['Sheet1']
    max_row = 2

    while True:
        if sheet[f"A{max_row}"].value:
            max_row = max_row + 1
        else:
            break

    sheet[f"A{max_row}"] = data
    sheet[f"B{max_row}"] = subject
    sheet[f"C{max_row}"] = chapter
    sheet[f"D{max_row}"] = duration

    wb.save(filename)
