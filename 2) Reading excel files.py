# load_workbook() --> open an existing excel file
# sheetnames --> get all worksheet names
# wb.active --> access the active worksheet
# wb["SheetName"] --> access a specific worksheet
# ws["A1"].value --> read an individual cell
# ws.cell() --> read cells using row/column numbers
# iter_rows() --> read multiple rows
# iter_cols() --> read multiple columns
# max_row, max_column --> find worksheet boundaries

# ------------------------------------- 
from openpyxl import load_workbook

wb=load_workbook("employees.xlsx") 

print("Workbook loaded successfully")

print(wb.sheetnames)

ws=wb.active
print(ws.title)

ws=wb["Employees"]
print(ws.title)

print(ws["A1"].value)
print(ws["B2"].value)
print(ws["C3"].value)

print(ws.cell(row=1, column=2).value)
print(ws.cell(row=3, column=3).value)

print(ws.max_row)
print(ws.max_column)

# read all rows using a loop
for i in range(1, ws.max_row+1):
    for j in range(1, ws.max_column+1):
        print(ws.cell(row=i, column=j).value, end=" ")
    print()

# read rows using iter_rows()
for i in ws.iter_rows(values_only=True):
    print(i)

# skipping the heading row
for i in ws.iter_rows(min_row=2,values_only=True):
    print(i)

# read columns using iter_cols()
for col in ws.iter_cols(values_only=True):
    print(col)

for i in ws.iter_cols(min_col=3, max_col=3, min_row=2, values_only=True):
    print(i)

# ------------------------------------- 
from openpyxl import load_workbook

wb=load_workbook("employees.xlsx")
ws=wb["Employees"]

for i in ws.iter_rows(min_row=2, values_only=True):
    employee_id, name, salary=i

    if salary>30000:
        print(name, salary)

wb.close()