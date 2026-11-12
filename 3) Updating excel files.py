# ws.cell() --> update using row and column numbers
# ws.append() --> add new employee records
# insert_rows() --> insert a row at a particular position
# delete_rows() --> delete existing rows
# delete_cols() --> delete columns
# ws.iter_rows() --> update multiple records using loops
# wb.save() --> save changes to excel

""" from openpyxl import load_workbook

# open the excel file
wb=load_workbook("employees.xlsx")
ws=wb["Employees"]

# update Arun's salary
ws["C2"]=50000
wb.save("employees.xlsx")

# update cells using ws.cell()
ws.cell(row=2, column=3).value=35000
wb.save("employees.xlsx")

# add new employee records
ws.append([104, "Priya", 45000])
ws.append([105, "Kiran", 50000])
wb.save("employees.xlsx")

# insert a row at specific position
ws.insert_rows(3)

ws.cell(row=3, column=1).value=106
ws.cell(row=3, column=2).value="Meena"
ws.cell(row=3, column=3).value=38000
wb.save("employees.xlsx")

# delete a row
ws.delete_rows(3)
ws.delete_cols(3)
 """

# ------------------------------------
# increase everyone's salary by 10%

from openpyxl import load_workbook

wb=load_workbook("employees.xlsx")
ws=wb["Employees"]

# loop through employee rows
for i in range(2, ws.max_row+1):
    new_salary=ws.cell(row=i, column=3).value * 1.10 # calculate new salary
    ws.cell(row=i, column=3).value=new_salary

wb.save("employees.xlsx")

print("Salary increment completed successfully")
