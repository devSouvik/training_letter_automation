import time
import openpyxl
from docxtpl import DocxTemplate

start_time = time.perf_counter ()

wb = openpyxl.load_workbook("employee_details.xlsx")
thisDoc = DocxTemplate("Letter_of_Training.docx")

sheet = wb["Sheet1"]
header = []
data = []
ws = wb.active
rows_count = 0

for row in ws:
    if not all([cell.value is None for cell in row]):
        rows_count += 1

columns = sheet.max_column

for emp in range(2, rows_count + 1):
    for cols in range(1, columns + 1):
        header.append(sheet.cell(1, cols).value)

    for cols in range(1, columns + 1):
        data.append(sheet.cell(emp, cols).value)

    context = dict(zip(header, map(str, data)))
    # print(context)
    # print("====")

    thisDoc.render(context)
    thisDoc.save(f'{context["employee_name"]}_training_letter.docx')

end_time = time.perf_counter()
print(end_time - start_time, "seconds")
