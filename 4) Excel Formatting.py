# Font --> change font size, boldness and color
# PatternFill --> apply background colors
# Alignment --> align text within cells
# Border --> appy cell borders
# column_dimensions --> adjust column widths
# number_format --> format salary, dates, percentages
# freeze_panes --> keep headers visible while scrolling
# auto_filter --> enable excel filtering
# ConditionalFormatting --> highlight values based on conditions 

# ---------------------------------------

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
from openpyxl.formatting.rule import CellIsRule

wb=load_workbook("employees.xlsx")
ws=wb["Employees"]

# make all headers bold
for i in ws[1]: # selects all cells in row 1
    i.font=Font(
        bold=True,
        size=12,
        color="000000"
    )

# add background colors
for i in ws[1]:
    i.fill=PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

# text alignment
for i in ws[1]:
    i.alignment=Alignment(
        horizontal="center",
        vertical="center"
    )

# borders
for i in ws.iter_rows():
    for j in i:
        j.border=Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

# adjust column width
ws.column_dimensions["A"].width=18
ws.column_dimensions["B"].width=22
ws.column_dimensions["C"].width=18

# set row height
ws.row_dimensions[1].height=28

# format salry values using number_format
for i in range(2, ws.max_row+1):
    ws.cell(row=i, column=3).number_format="₹#,##0"

# freeze the header row
ws.freeze_panes="A2"

# enable excel filters- adds filter dropdowns to your excel headings
ws.auto_filter.ref=ws.dimensions

# conditional formatting- highlight salaries greater than 40000
ws.conditional_formatting.add(
    f"C2:C{ws.max_row}",
    CellIsRule(
        operator="greaterThan",
        formula=["40000"],
        fill=PatternFill(
            fill_type="solid",
            fgColor="FFC6EFCE"
        )
    )
)

wb.save("employees.xlsx")