import os
import openpyxl
from utils import excelhandling as ex

excel = os.path.abspath('rajulaw.xlsx')
max_row = ex.getRowCount(excel, 'All Client Data')
# max_col = ex.getColCount(excel, 'All Client Data')
# print(max_row)
# print(max_col)

wb = openpyxl.load_workbook(excel)
email_list = []
for row in range(2, max_row + 1):
    service_request2 = ex.open_and_read_excel_file_by_row(excel, 'All Client Data', row)
    service_visa = service_request2[-1]
    # email_list = service_request2[2]
    # print(email_list)
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
