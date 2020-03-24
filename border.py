from openpyxl import Workbook
from openpyxl.styles import Border, Side

wb = Workbook()
ws = wb.active
rng = ws['B2':'D3']
s = Side(style = 'dashed', color = '0000ff')
for row in rng:
    for c in row:
        c.value = 123
        c.border = Border(left = s, right = s, top = s, bottom = s) 
wb.save('border.xlsx')

