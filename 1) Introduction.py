# openpyxl --> read, create, update and format .xlsx files

from openpyxl import Workbook

wb=Workbook() # create a new excel workbook

ws=wb.active # get the active worksheet

ws.title="Employees" # rename the worksheet

ws.append(["Employee ID", "Name", "Salary"]) # add column headings

# add employee records
ws.append([101, "Arun", 30000])
ws.append([102, "Sneha", 40000])
ws.append([103, "Ravi", 35000])

wb.save("employees.xlsx") # save the workbook

print("Excel file created successfully")