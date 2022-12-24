'''
python 3.10.0
'''

import openpyxl
import os


class Excel:
    def __init__(self):
        self.filename = ''  # 文件名
        self.type = ''  # 出/入库类型
        self.rds = []
        pass

    pass


def readfile():
    e = Excel()
    with open(os.getcwd() + r'\text.dat', 'r', encoding='utf-8') as f:
        # 逐行读取
        e.type = f.readline()[0:-1]
        e.filename = f.readline()[0:-1]
        line = f.readline()[0:-1]
        while line:
            rd = line.split(' ')
            e.rds.append(rd)
            line = f.readline()
            pass
        pass
    return e
    pass


def create_excel(e):
    workbook = openpyxl.Workbook()  # 创建一个Workbook对象，相当于创建了一个Excel文件
    worksheet = workbook.active  # 获取当前活跃的worksheet,默认就是第一个worksheet
    worksheet.title = 'Sheet1'  # 给当前活跃的worksheet命名
    if e.type == 'import':
        worksheet.append(['日期', '物料编码', '图号名称', '入库', '单位', '剩余库存'])  # 写入表头即第一行
        pass
    else:
        worksheet.append(['日期', '物料编码', '图号名称', '出库', '单位', '剩余库存'])
        pass
    width = [12, 12, 12, 12, 12, 12]
    for i in e.rds:
        if width[1] < len(i[1]) + 1:
            width[1] = len(i[1]) + 1
            pass
        if width[2] < (len(i[2]) + 1) * 2:
            width[2] = (len(i[2]) + 1) * 2
            pass
        if width[3] < len(i[3]) + 1:
            width[3] = len(i[3]) + 1
            pass
        if width[4] < (len(i[4]) + 1) * 2:
            width[4] = (len(i[4]) + 1) * 2
            pass
        if width[5] < len(i[5]) + 1:
            width[5] = len(i[5]) + 1
            pass

        worksheet.append(i)
        pass
    # 设置单元格宽度
    for i in range(len(width)):
        worksheet.column_dimensions[chr(ord('A') + i)].width = width[i]
        pass

    workbook.save(e.filename)
    pass


if __name__ == '__main__':
    e = readfile()
    create_excel(e)
    pass
