from utils import put_name_to_pdf
import openpyxl
import os


new_direcotry="Доска почета"
os.makedirs(new_direcotry, exist_ok=True)


file_xlsx = "Сотрудники.xlsx"

read_excel = openpyxl.load_workbook(file_xlsx)
worksheet = read_excel.active
data=[]


for row in worksheet.iter_rows(min_row=2, values_only=True):
    first_name, last_name, middle_name = row
    data.append((first_name, last_name, middle_name))


for first_name, last_name, middle_name in data:
    FIO = f"{first_name} {last_name} {middle_name}"
    name_pdf_file = f"{first_name} {last_name}"
    put_name_to_pdf("Шаблон.docx", name_pdf_file, FIO, new_direcotry)

read_excel.close()