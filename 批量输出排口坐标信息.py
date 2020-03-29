# coding: UTF-8

import ezdxf, xlwt, os
from xlutils.copy import copy
from xlrd import open_workbook


# get text data from dxf files
filename = input("filename")
txt_num = input("txt_num")
doc = ezdxf.readfile(os.getcwd() + "/%s" %filename)
msp = doc.modelspace()
for text in msp.query("TEXT"):
    if text.dxf.layer == "ADD_PK":
        text_content = text.dxf.text
        text_x_cood = text.dxf.insert[1]
        text_y_cood = text.dxf.insert[0]
        text_name = text_content.split("/")[0]
        pipe_diameter = text_content.split("/")[1]
        elevation = text_content.split("/")[2]
        print(text_name, x_cood, y_cood, pipe_diameter, elevation)
        