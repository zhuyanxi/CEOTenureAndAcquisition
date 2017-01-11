import open_excel as oe
import numpy as np
from collections import OrderedDict


# 根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file, colnameindex=0, by_index=0):
    data = oe.open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    colnames = table.row_values(colnameindex)  # 某一行数据
    list = []
    for rownum in range(1, nrows):

        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
    return list


# 根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_name：Sheet1名称
def excel_table_to_NormalDict(file, by_name, colnameindex=0):
    data = oe.open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows  # 行数
    colnames = table.row_values(colnameindex)  # 某一行数据
    resultlist = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            resultlist.append(app)
            # print(app)
    return resultlist


def excel_table_to_OrderedDict_byIndex(file, colnameindex=0, by_index=0,dataRow=1):
    data = oe.open_excel(file)
    table = data.sheets()[by_index]
    nRows = table.nrows  # 行数
    colnames = table.row_values(colnameindex)  # 某一行数据
    resultList = []
    for rowNum in range(dataRow, nRows):
        row = table.row_values(rowNum)
        if row:
            app = OrderedDict()
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            resultList.append(app)
            # print(app)
    return resultList


def excel_table_to_OrderedDict_bySheetName(file, by_name, colnameindex=0):
    data = oe.open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows  # 行数
    colnames = table.row_values(colnameindex)  # 某一行数据
    resultlist = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = OrderedDict()
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            resultlist.append(app)
            # print(app)
    return resultlist


def excel_table_to_numpy_array(file, by_name, colnameindex=0):
    data = oe.open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows  # 行数
    colnames = table.row_values(colnameindex)  # 某一行数据
    result = np.array()


def subsample_filter(file, by_name, colnameindex=0):
    data = oe.open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows  # 行数
    colnames = table.row_values(colnameindex)  # 某一行数据
    resultlist = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            if app['D0702b'] != '':
                resultlist.append(app)
                print(app)
    return resultlist
