from openpyxl import Workbook

wb = Workbook()
ws = wb.active
rng = ws['A1':'C2']
for row in rng:
    for c in row:
        c.value = 123
wb.save('range.xlsx')

