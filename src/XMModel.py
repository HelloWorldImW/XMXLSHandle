#! /usr/bin/env python
# _*_ coding:utf-8 _*_

'ExlModel'

__author__ = 'DarrenW'

import xlrd

class baseExlModel(object):

    def __init__(self,fileName):
        self.fileName = fileName
        self.xlsData = self.loadXls(fileName)

    #读取excel表 并返回excel表的python对象
    def loadXls(self,fileName):
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


class iOSExlModel(baseExlModel):
    pass







class JavaExlModel(baseExlModel):

    pass








class TestExlModel(baseExlModel):
    pass









class YWExlModel(baseExlModel):
    pass