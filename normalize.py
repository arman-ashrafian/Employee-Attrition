import openpyxl

#Open an xlsx for reading
wb = openpyxl.load_workbook('Results.xlsx')
#Get the current Active Sheet
sheet = wb.get_sheet_by_name('Training')

def norm_education_feild():
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

def norm_attrition():
    attrition_col = sheet.columns[1]

    for cell in attrition_col:
        if cell.value == 'Yes':
            sheet[cell.coordinate].value = 1
        elif cell.value == 'No':
            sheet[cell.coordinate].value = 0

def norm_bussiness_travel():
    bussiness_travel_col = sheet.columns[2]

    for cell in bussiness_travel_col:
        if cell.value == 'Non-Travel':
            sheet[cell.coordinate].value = 0
        elif cell.value == 'Travel_Rarely':
            sheet[cell.coordinate].value = 1
        elif cell.value == 'Travel_Frequently':
            sheet[cell.coordinate].value = 2

def norm_department():
    department_col = sheet.columns[4]

    for cell in department_col:
        if cell.value == 'Research & Development':
            sheet[cell.coordinate].value = 0
        elif cell.value == 'Human Resources':
            sheet[cell.coordinate].value = 1
        elif cell.value == 'Sales':
            sheet[cell.coordinate].value = 2

def norm_gender():
    gender_col = sheet.columns[11]

    for cell in gender_col:
        if cell.value == 'Female':
            sheet[cell.coordinate].value = 0
        elif cell.value == 'Male':
            sheet[cell.coordinate].value = 1

def norm_job_role():
    job_role_col = sheet.columns[15]

    for cell in job_role_col:
        if cell.value == 'Research Scientist':
            sheet[cell.coordinate].value = 0
        elif cell.value == 'Sales Representative':
            sheet[cell.coordinate].value = 1
        elif cell.value == 'Sales Executive':
            sheet[cell.coordinate].value = 2
        elif cell.value == 'Laboratory Technician':
            sheet[cell.coordinate].value = 3
        elif cell.value == 'Healthcare Representative':
            sheet[cell.coordinate].value = 4
        elif cell.value == 'Research Director':
            sheet[cell.coordinate].value = 5
        elif cell.value == 'Manufacturing Director':
            sheet[cell.coordinate].value = 6
        elif cell.value == 'Manager':
            sheet[cell.coordinate].value = 7
        elif cell.value == 'Human Resources':
            sheet[cell.coordinate].value = 8

def norm_marital_status():
    marital_status_col = sheet.columns[17]

    for cell in marital_status_col:
        if cell.value == 'Single':
            sheet[cell.coordinate].value = 0
        elif cell.value == 'Married':
            sheet[cell.coordinate].value = 1
        elif cell.value == 'Divorced':
            sheet[cell.coordinate].value = 2

def norm_overtime():
    overtime_col = sheet.columns[22]

    for cell in overtime_col:
        if cell.value == 'No':
            sheet[cell.coordinate].value = 0
        elif cell.value == 'Yes':
            sheet[cell.coordinate].value = 1

# norm_gender()
# norm_job_role()
# norm_overtime()
norm_attrition()
# norm_department()
# norm_marital_status()
# norm_education_feild()
# norm_bussiness_travel()
wb.save('Normalized.xlsx')
