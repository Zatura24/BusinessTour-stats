from openpyxl import Workbook

#creating workbook
wb = Workbook()

#setting filename
dest_filename = 'Test.xlsx'

#using first worksheet and setting name
ws = wb.active
ws.title = "Sheet1"

#defining first cell and changing value
c = ws['A1']
c.value = "Memes"

#saving workbook
wb.save(filename = dest_filename)