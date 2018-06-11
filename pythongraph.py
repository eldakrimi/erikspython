import openpyxl

wb = openpyxl.Workbook()
sheet = wb.create_sheet('My Sheet')
import random
for i in range(1, 11):
    sheet['A' + str(i)].value = random.randint(1, 100)

wb.save('exem3.xlsx')

