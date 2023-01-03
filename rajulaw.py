import datetime
import os
import openpyxl
from utils import excelhandling as ex

excel = os.path.abspath('rajulaw.xlsx')
max_row = ex.getRowCount(excel, 'All Client Data')
max_col = ex.getColCount(excel, 'All Client Data')
# print(max_row)
# print(max_col)
wb = openpyxl.load_workbook(excel)
one = 1
email_list = []
date = str(datetime.date.today())
split_date = date.split('-')
todays_format = f'{split_date[0][2:4]}{split_date[1]}{split_date[2]}'
for row in range(2, max_row + 1):
    row_all_value = str(ex.open_and_read_excel_file_by_row(excel, 'All Client Data', row)[0])
    id = row_all_value[0:6]
    # ex.writeData(excel, 'All Client Data', row, 1, f"{split_date[0]}_{'Dec'}_{row - 1}")
    if id != todays_format:
        ex.writeData(excel, 'All Client Data', row, 1, f'{todays_format}{one}')
        one = one + 1
        # ex.writeData(excel, 'All Client Data', row, 1, f"{split_date[0][2:4]}_{'Dec'}_{row - 1}")
    else:
        print('user already exist')

for row in range(2, max_row + 1):
    service_request2 = ex.open_and_read_excel_file_by_row(excel, 'All Client Data', row)
    service_visa = service_request2[-1]
    email_list.append(service_request2[2])

    if service_visa == "Student Visa":
        currentSheets = wb[service_visa]
        currentSheets_max_row = ex.getRowCount(excel, service_visa)
        currentSheets_max_col = ex.getColCount(excel, service_visa)
        email = ex.open_and_read_excel_file(excel, service_visa, 3)

        if email_list[row - 2] not in email:
            for col in range(1, currentSheets_max_col + 1):
                ex.writeData(excel, service_visa, currentSheets_max_row + 1, col, service_request2[col - 1])
        else:
            print(f"{email_list[row - 2]} is Already Exist")



    elif service_visa == "Work Visa":
        currentSheets = wb[service_visa]
        currentSheets_max_row = ex.getRowCount(excel, service_visa)
        currentSheets_max_col = ex.getColCount(excel, service_visa)
        email = ex.open_and_read_excel_file(excel, service_visa, 3)

        if email_list[row - 2] not in email:
            for col in range(1, currentSheets_max_col + 1):
                ex.writeData(excel, service_visa, currentSheets_max_row + 1, col, service_request2[col - 1])
        else:
            print(f"{email_list[row - 2]} is Already Exist")

    elif service_visa == "Family Immigration":
        currentSheets = wb[service_visa]
        currentSheets_max_row = ex.getRowCount(excel, service_visa)
        currentSheets_max_col = ex.getColCount(excel, service_visa)
        email = ex.open_and_read_excel_file(excel, service_visa, 3)

        if email_list[row - 2] not in email:
            for col in range(1, currentSheets_max_col + 1):
                ex.writeData(excel, service_visa, currentSheets_max_row + 1, col, service_request2[col - 1])
        else:
            print(f"{email_list[row - 2]} is Already Exist")

    elif service_visa == "Marriage Immigration":
        currentSheets = wb[service_visa]
        currentSheets_max_row = ex.getRowCount(excel, service_visa)
        currentSheets_max_col = ex.getColCount(excel, service_visa)
        email = ex.open_and_read_excel_file(excel, service_visa, 3)

        if email_list[row - 2] not in email:
            for col in range(1, currentSheets_max_col + 1):
                ex.writeData(excel, service_visa, currentSheets_max_row + 1, col, service_request2[col - 1])
        else:
            print(f"{email_list[row - 2]} is Already Exist")

    elif service_visa == "PR":
        currentSheets = wb[service_visa]
        currentSheets_max_row = ex.getRowCount(excel, service_visa)
        currentSheets_max_col = ex.getColCount(excel, service_visa)
        email = ex.open_and_read_excel_file(excel, service_visa, 3)

        if email_list[row - 2] not in email:
            for col in range(1, currentSheets_max_col + 1):
                ex.writeData(excel, service_visa, currentSheets_max_row + 1, col, service_request2[col - 1])
        else:
            print(f"{email_list[row - 2]} is Already Exist")

    elif service_visa == "Citizenship":
        currentSheets = wb[service_visa]
        currentSheets_max_row = ex.getRowCount(excel, service_visa)
        currentSheets_max_col = ex.getColCount(excel, service_visa)
        email = ex.open_and_read_excel_file(excel, service_visa, 3)

        if email_list[row - 2] not in email:
            for col in range(1, currentSheets_max_col + 1):
                ex.writeData(excel, service_visa, currentSheets_max_row + 1, col, service_request2[col - 1])
        else:
            print(f"{email_list[row - 2]} is Already Exist")