import os
from openpyxl import load_workbook

# Local Excel file connected to Jenkins through Docker
file_path = "/employee-data/employees.xlsx"

# Read employee details from Jenkins parameters
s_no = os.environ["S_NO"].strip()
name = os.environ["EMPLOYEE_NAME"].strip()
employee_id = os.environ["EMPLOYEE_ID"].strip()

# Validate entered details
if not s_no or not name or not employee_id:
    raise ValueError("Please enter all employee details")

# Open existing Excel file
workbook = load_workbook(file_path)
sheet = workbook.active

# Verify column headings
headers = [
    str(sheet.cell(1, col).value).strip()
    for col in (1, 2, 3)
]

if headers != ["S_NO", "NAME", "EMPLOYEE_ID"]:
    raise ValueError(f"Unexpected Excel headings: {headers}")

# Check for duplicate serial numbers
for row in sheet.iter_rows(min_row=2, max_col=3, values_only=True):
    if row[0] is not None and str(row[0]).strip() == s_no:
        raise ValueError("This serial number already exists")

# Find next available row
next_row = sheet.max_row + 1

# Add new employee
for column, value in enumerate(
    [s_no, name, employee_id], start=1
):
    cell = sheet.cell(row=next_row, column=column)
    cell.value = value
    cell.number_format = "@"

# Save updated Excel file
workbook.save(file_path)

print("Employee added successfully!")
print("Serial Number:", s_no)
print("Employee Name:", name)
print("Employee ID:", employee_id)
print("Total Employees:", sheet.max_row - 1)