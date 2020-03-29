#!/usr/bin/env python3
# -*-coding: UTF-8-*-

import ezdxf, xlwt, os
from xlutils.copy import copy
from xlrd import open_workbook


def write_data_into_excel(excel_name, sheet_num, *args):
    '''write data extract from cad files into excel files(.xls)'''
    print("Begin to read excel files*****************")
    rb = open_workbook(excel_name)
    wb = copy(rb)    
    wn = wb.get_sheet(sheet_num) # get the sheet of wb
    row_start = len(wn.get_rows()) # get the total rows of wn
    for i in range(1, len(args)):
        wn.write(row_start, i, args[i])
    wb.save(os.getcwd() + "/" + excel_name)

def get_and_write_data(file_name, excel_name, sheet_num):
    '''extract data from cad files(.dxf)'''
    print("Begin to extract data from dxf files****************")
    doc = ezdxf.readfile(os.getcwd() + "/%s" %file_name)
    msp = doc.modelspace()
    for text in msp.query("TEXT"):
        if text.dxf.layer == "ADD_PK":
            text_content = text.dxf.text
            text_x_cood = text.dxf.insert[1]
            text_y_cood = text.dxf.insert[0]
            text_name = text_content.split("/")[0]
            pipe_diameter = text_content.split("/")[1]
            elevation = text_content.split("/")[2]
            text_info = text_name, text_x_cood, text_y_cood, pipe_diameter, elevation
            print(text_info)
            print("Begin to write data into excel files******************")
            write_data_into_excel(excel_name, sheet_num, *text_info)


for root, dirs, files in os.walk(os.getcwd()):
    for filename in files:
        if ".xls" in filename:
            global excel_name
            excel_name = filename
            print("[+]Get the excel file %s********************" % filename)
        if ".dxf" in filename:
            print("[+]Begin to operate %s********************" % filename)
            get_and_write_data(filename, excel_name, int(list(filter(str.isdigit, filename))[0])-1)


 


    
        
