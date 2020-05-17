#coding: utf-8
from openpyxl import Workbook
from openpyxl.chart import BarChart3D, Reference
from openpyxl.chart.label import DataLabelList

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
chart = BarChart3D()
chart.add_data(values, titles_from_data=True)
chart.set_categories(cats)
chart.title = u'売上推移'
chart.x_axis.title = u'年'
chart.y_axis.title = u'売上'
chart.dataLabels = DataLabelList()
chart.dataLabels.showVal = True
ws.add_chart(chart, "E1")
wb.save("datalabel.xlsx")
