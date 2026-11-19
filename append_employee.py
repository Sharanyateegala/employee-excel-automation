import os
from pathlib import Path
from openpyxl import load_workbook

# Excel file downloaded from GitHub into Jenkins workspace
file_path = Path(__file__).resolve().parent / "employees.xlsx"

# Read employee details from Jenkins parameters
s_no = os.environ["S_NO"].strip()
name = os.environ["EMPLOYEE_NAME"].strip()
employee_id = os.environ["EMPLOYEE_ID"].strip()

# Validate entered details
if not s_no or not name or not employee_id:
    raise ValueError("Please enter all employee details")

# Check if Excel file exists
if not file_path.exists():
    raise FileNotFoundError(f"Excel file not found: {file_path}")

# Open existing Excel file
workbook = load_workbook(file_path)
sheet = workbook.active

# Verify column headings
headers = [
    str(sheet.cell(1, col).value).strip()
    for col in (1, 2, 3)
]

if headers != ["S_NO", "EMPLOYEE_NAME", "EMPLOYEE_ID"]:
    raise ValueError(f"Unexpected Excel headings: {headers}")

# Check for duplicate serial numbers
for row in sheet.iter_rows(
    min_row=2,
    max_col=3,
    values_only=True
):
    if row[0] is not None and str(row[0]).strip() == s_no:
        raise ValueError(
            f"Serial number {s_no} already exists"
        )

# Find the next available row
next_row = sheet.max_row + 1

# Add new employee details
employee_details = [s_no, name, employee_id]

for column, value in enumerate(employee_details, start=1):
    cell = sheet.cell(row=next_row, column=column)
    cell.value = value

    # Store values as text to preserve leading zeros
    cell.number_format = "@"

# Save updated Excel file
workbook.save(file_path)

# Display output in Jenkins Console
print("==================================")
print("Employee added successfully!")
print("==================================")
print("Serial Number:", s_no)
print("Employee Name:", name)
print("Employee ID:", employee_id)
print("Total Employees:", sheet.max_row - 1)
print("Excel File:", file_path)
print("==================================")