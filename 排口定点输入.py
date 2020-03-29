# coding: UTF-8

import xlrd, dxfwrite
from dxfwrite import DXFEngine as dxf

# draw a cad picture
drawing = dxf.drawing('排口坐标.dxf')

# read excel datas
data = xlrd.open_workbook('1.xls')
num_sheets = len(data.sheets())
for n in range(num_sheets):
    table = data.sheets()[n]
    nrows = table.nrows
    for i in range(2, nrows):
        x_cood = float(table.cell(i, 2).value)
        y_cood = float(table.cell(i, 3).value)
        pai_num = table.cell(i, 1).value
        circle = dxf.circle(2.0)
        circle['layer'] = 'paikou'
        circle['color'] = 2
        text = dxf.text(pai_num, (y_cood, x_cood), height=1.207)
        text['layer']='paikou'
        text['color'] = 2
        block = dxf.block(name='paikou')
        block.add(circle)
        drawing.blocks.add(block)
        blockref = dxf.insert(blockname='paikou', insert=(y_cood, x_cood))
        drawing.add(blockref)
        drawing.add(text)

drawing.save()

