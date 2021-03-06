#coding: utf-8
from openpyxl import Workbook
from openpyxl.chart import AreaChart, Reference

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
chart = AreaChart()
chart.add_data(values, titles_from_data=True)
chart.set_categories(cats)
chart.title = u'売上推移'
chart.x_axis.title = u'年'
chart.y_axis.title = u'売上'
chart.y_axis.scaling.min = 0
chart.y_axis.scaling.max = 1500
chart.y_axis.majorUnit = 500
chart.y_axis.minorUnit = 100
ws.add_chart(chart, "E1")
wb.save("axis.xlsx")
