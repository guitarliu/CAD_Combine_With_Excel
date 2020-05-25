import ezdxf,os
from openpyxl import Workbook

while True:
    print("【+】**************************************")
    print("\t在读取dxf文件时，尽可能得将文件内无关内容清除！")
    print("【+】**************************************\n")
    filename = input("请在当前路径下输入dxf文件名（不需后缀,需解密）：")
    block_format = input("请输入块的格式（如块名‘20200525.115912’格式为20200525）,块名'预留口13'格式为预留口：")
    excel_name = input("请输入要保存的Excel文件名：（不需后缀）")
    doc = ezdxf.readfile(os.getcwd() + "/%s.dxf" % filename)
    msp = doc.modelspace()
    wb = Workbook()
    ws = wb.active
    for b in doc.blocks:
        if block_format in b.name:
            for flag_ref in msp.query('INSERT[name=="%s"]' % b.name):
                content = [entity.dxf.text for entity in flag_ref.virtual_entities() if entity.dxftype() == 'TEXT']
                print(content)
                ws.append(content)
                wb.save(os.getcwd() + "/" + "%s.xlsx" % excel_name)