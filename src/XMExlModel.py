#! /usr/bin/env python
# _*_ coding:utf-8 _*_

'ExlModel'

__author__ = 'DarrenW'

import xlrd

class BaseExlModel(object):

    def __init__(self,fileName):
        self.fileName = fileName
        self.xlsData = self.__loadXls(fileName)
        self.loadTable()

    #读取excel表 并返回excel表的python对象
    def __loadXls(self,fileName):
        xlsData = xlrd.open_workbook(fileName)
        return xlsData

    #读取某一个sheet表
    def loadTable(self,sheetIndex = 0):
        self.table = self.xlsData.sheet_by_index(sheetIndex)

    #获取sheet表的某一行,并返回
    def rowValues(self,rowNumber = 0):
        return self.table.row_values(rowNumber)

    #获取某一列的值,并返回
    def colValues(self,colNumber = 0):
        return self.table.col_values(colNumber)

    #获取具体某个单元格的值 col:列  row:行
    def cellValue(self, col, row):
        return self.table.cell_value(row,col)


# iOS周报数据model
class iOSExlModel(BaseExlModel):



    def __init__(self, fileName):
        super(iOSExlModel, self).__init__(fileName)
        self.hahaha()

    def hahaha(self):
        aa = self.colValues(1)
        for index, ss in enumerate(aa):
            print index,ss



























# android周报数据model
class AndroidExlModel(BaseExlModel):
    def __init__(self, fileName):
        super(AndroidExlModel, self).__init__(fileName)

# Java周报数据model
class JavaExlModel(BaseExlModel):
    def __init__(self, fileName):
        super(JavaExlModel, self).__init__(fileName)


# 测试周报数据model
class TestExlModel(BaseExlModel):
    def __init__(self, fileName):
        super(TestExlModel, self).__init__(fileName)


# 运维周报数据model
class YWExlModel(BaseExlModel):
    def __init__(self, fileName):
        super(YWExlModel, self).__init__(fileName)