import time
import openpyxl
import os.path
from docxtpl import DocxTemplate
from docx2pdf import convert
import format


start_time = time.perf_counter()


wb = openpyxl.load_workbook("employee_details.xlsx")
thisDoc = DocxTemplate("Letter_of_Training.docx")

# format the excel file
format.format_xl("employee_details.xlsx")

sheet = wb["Sheet1"]
header = []
data = []
ws = wb.active
# print(ws)
rows_count = 0

for row in ws:
    if not all([cell.value is None for cell in row]):
        rows_count += 1

# check for empty rows in between

columns = sheet.max_column
print(columns)
print(rows_count)
# {emp_name:"souvik", "emp_salary":2000,}

for emp in range(2, rows_count + 1):
    for cols in range(1, columns + 1):
        header.append(sheet.cell(1, cols).value)
        data.append(sheet.cell(emp, cols).value)

# print(header)
# print("")
# print(data)
# print("")

    context = dict(zip(header, map(str, data)))
    # print(context)
    # print("")

    try:
        if os.path.isfile(f'./letters/{context["employee_name"]}_training_letter.docx'):
            print("File already exists")
        else:
            thisDoc.render(context)
            thisDoc.save(f'./letters/{context["employee_name"]}_training_letter.docx')
            print(f"new file created for => {context['employee_name']}")

            # convert to pdf

            # input_file = f'./letters/{context["employee_name"]}_training_letter.docx'
            # output_file = f'./letters/{context["employee_name"]}_training_letter.pdf'
            # convert(input_file, output_file)
            # convert("./letters")
    except:
        print("something went wrong")

end_time = time.perf_counter()
print(end_time - start_time, "seconds")
