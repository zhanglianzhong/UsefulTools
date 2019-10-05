# -*- coding: utf-8 -*-
import xlrd

def readRecieveInfo(filePath):
    '''
    Args
     filePath:包含接收者信息的excel文件路径
    return:
    a list,[(username,email)]
    '''
    # 打开execl
    workbook = xlrd.open_workbook(filePath)

    # 根据sheet索引或者名称获取sheet内容
    Data_sheet = workbook.sheets()[0]  # 通过索引获取
    # Data_sheet = workbook.sheet_by_index(0)  # 通过索引获取
    # Data_sheet = workbook.sheet_by_name(u'名称')  # 通过名称获取

    #print(Data_sheet.name)  # 获取sheet名称
    rowNum = Data_sheet.nrows  # sheet行数
    colNum = Data_sheet.ncols  # sheet列数

    # 获取所有单元格的内容
    list = []
    for i in range(1,rowNum):
        #get username email
        list.append((Data_sheet.cell_value(i, 1), Data_sheet.cell_value(i, 2)))
    return  list
