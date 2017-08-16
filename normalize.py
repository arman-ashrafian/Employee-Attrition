import openpyxl
import csv

#Open an xlsx for reading
wb = openpyxl.load_workbook('Results.xlsx')
#Get the current Active Sheet
sheet = wb.get_sheet_by_name('Training')

education_feild_col = sheet.columns[7]

for cell in education_feild_col:
    if cell.value == 'Life Sciences':
        sheet[cell.coordinate].value = 1
    elif cell.value == 'Medical':
        sheet[cell.coordinate].value  = 2
    elif cell.value == 'Technical Degree':
        sheet[cell.coordinate].value  = 3
    elif cell.value == 'Human Resources':
        sheet[cell.coordinate].value  = 4
    elif cell.value == 'Marketing':
        sheet[cell.coordinate].value  = 5
    elif cell.value == 'Other':
        sheet[cell.coordinate].value  = 6

wb.save('Normalized.xlsx')
