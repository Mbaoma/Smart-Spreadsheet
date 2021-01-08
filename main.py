# import the openpyxl library
from openpyxl import Workbook

#add stying to the sheet by importing the following libraries
from openpyxl.styles import NamedStyle, Font, Color, Border, colors, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule

#create the spreadsheet
our_workbook = Workbook()
sheet_one = our_workbook.active

#name the sheet
sheet_one.title = "Employee details"

#effect the styling
bold_font = Font(bold=True, color=colors.BLUE, size=10)
font_border = Border(bottom=Side(border_style='thin'))
sheet_one.freeze_panes = "A1"

#add titles to the spreadsheet
sheet_one["A1"] = "NAME"
sheet_one["A1"].font = bold_font

sheet_one["B1"] = "GROSS PAY (N)"
sheet_one["B1"].font = bold_font

sheet_one["C1"] = "TOTAL WITH-HOLDINGS (N)"
sheet_one["C1"].font = bold_font

sheet_one["D1"] = "NET AMOUNT PAYABLE (N)"
sheet_one["D1"].font = bold_font

#save the file with the 'xlsx' extension
our_workbook.save(filename= "sheet_one.xlsx")

