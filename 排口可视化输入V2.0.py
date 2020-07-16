# coding: utf-8

import os, re
from openpyxl import *
from dxfwrite import DXFEngine as dxf

def create_cad():
    '''
    :return: a cad file(.dxf)
    '''
    global cad_name
    project_list = []
    excel_name = input("\033[1;32m请输入Excel文件名称(粘贴复制无后辍哦~): \033[0m")
    num_projects = int(input("\033[1;32m请输入写入项目数(同Excel但除XY): \033[0m"))
    for project_num in range(1, num_projects + 1):
        pk_num= int(input("\033[1;32m请输入项目%s所在列号：\033[0m " % project_num))
        project_list.append(pk_num)
    pk_x_cood = int(input("\033[1;32m请输入排口横坐标所在列号：\033[0m"))
    pk_y_cood = int(input("\033[1;32m请输入排口纵坐标所在列号：\033[0m"))
    cad_name = input("请输入要保存的CAD文件名称：")
    # draw a cad picture
    drawing = dxf.drawing(os.getcwd() + "/%s.dxf" % cad_name)
    # read excel datas
    wb = load_workbook(os.getcwd() + "/" + excel_name + ".xlsx")
    ws = wb.active
    for row in ws.rows:
        for project_column in project_list:
            if str(row[pk_x_cood-1].value).replace(".", "").isdigit() == True and str(row[pk_y_cood-1].value).replace(".", "").isdigit() == True:
                print(row[project_column - 1].value, row[pk_x_cood - 1].value, row[pk_y_cood - 1].value)
                x_cood = float(row[pk_x_cood-1].value)
                y_cood = float(row[pk_y_cood-1].value)
                pai_num = row[project_column-1].value
                circle = dxf.circle(2.0)
                circle['layer'] = 'paikou'
                circle['color'] = 2
                text = dxf.text(pai_num, (x_cood, y_cood - 1.35 * project_list.index(project_column)), height=1.207)
                text['layer']='paikou'
                text['color'] = 2
                block = dxf.block(name='paikou')
                block.add(circle)
                drawing.blocks.add(block)
                blockref = dxf.insert(blockname='paikou', insert=(x_cood, y_cood))
                drawing.add(blockref)
                drawing.add(text)

    drawing.save()

while True:
    print('''
\033[1;34m【*】------------------------------------------------------------------------------------\n
【*】欢迎使用刘勇牌排口分布cad图生成器，请按照提示完全操作。\n
【*】需事先准备一份Excel文件（仅支持格式为.xlsx)，将其与该脚本放到同一个目录下。\n
【*】需按照提示分别输入Excel表格中排口编号、标高、排口横坐标、排口纵坐标等项目所在列号。\n
【*】处理过程中会看到有哪些坐标及文件信息被写入到CAD文件中。Hong\n
【*】处理完成后在脚本所在当前目录下生成名为\"*.dxf\"的文件，可打开查看或作外部参照。\n
【*】------------------------------------------------------------------------------------\033[0m
    ''')
    try:
        global cad_name, project_col_content
        create_cad()
        print('''
\033[36;5m【✔】处理完成，请在当前目录下查找\"%s.dxf\"文件!
【✔】-----------------------------------------\n\033[0m''' % cad_name)
    except:
        print("\033[1;31m列数必须为整数，且输入不能为空，请重新输入！\033[0m")