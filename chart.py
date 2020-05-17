#coding: utf-8
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

wb = Workbook()
ws = wb.active
rows = [
    [u'年', u'東京', u'大阪'],
    [2017, 1000, 700],
    [2018, 1200, 900],
    [2019, 1100, 800]
]
for row in rows:
    ws.append(row)

values = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=4)
cats = Reference(ws, min_col=1, min_row=2, max_row=4)
chart = BarChart()
chart.add_data(values, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "E1")
wb.save("chart.xlsx")
