#coding: utf-8
from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
ws = wb.active
c = ws['A1']
c.value = 123
c.font = Font(name = u'メイリオ', size = 12, color = 'FF0000', italic = True, bold = True)
wb.save('font.xlsx')
