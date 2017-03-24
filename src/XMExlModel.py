#! /usr/bin/env python
# _*_ coding:utf-8 _*_

'ExlModel'

__author__ = 'DarrenW'

import xlrd

#exl表内容model
class XMExlContents(object):
    def __init__(self):
        # 工作项
        self.workItems = {}
        # 是否遇到问题或风险
        self.hasWarning = 'warning'
        # 下周工作计划
        self.nextWorkPlan = 'nextWorkPlan'
        # 本周完成情况
        self.completeStatus = 'completeStatus'
        # 负责人
        self.charge = 'charge'

    # 添加工作项
    def addWorkItem(self,item):
        self.workItems[item] = {}

    # 获取所有的工作项
    def getWorkItems(self):
        if len(self.workItems):
            return self.workItems.keys()

    # 给某一工作项添加完成情况
    def addCompleteStatus(self,item,status):
        if self.workItems[item]:
            self.workItems[item][self.completeStatus] = status

    # 获取某一工作项的完成情况[]
    def getWorkItemCompleteStatus(self, item):
        if self.workItems[item]:
            return self.workItems[item][self.completeStatus]

    # 给某一工作项添加负责人
    def addCharge(self,item,charge):
        if self.workItems[item]:
            self.workItems[item][self.charge] = charge

    # 获取某一工作项的负责人
    def getWorkItemCharge(self, item):
        if self.workItems[item]:
            return self.workItems[item][self.completeStatus]

    # 给某一工作项添加风险
    def addWarning(self,item,warning):
        if self.workItems[item]:
            self.workItems[item][self.hasWarning] = warning

    # 获取某一工作项的风险
    def getWorkItemWarning(self, item):
        if self.workItems[item]:
            return self.workItems[item][self.completeStatus]

    # 给某一工作项添加下周计划
    def addNextWorkPlan(self,item,plan):
        if self.workItems[item]:
            self.workItems[item][self.nextWorkPlan] = plan

    #获取某一工作项的下周计划
    def getWorkItemWorkPlan(self,item):
        if self.workItems[item]:
            return self.workItems[item][self.completeStatus]


class XMExlModel(object):

    def __init__(self,fileName):
        self.fileName = fileName
        self.contentModel = XMExlContents()
        self.xlsData = self.__loadXls(fileName)
        self.__loadTable()
        self.__loadWorkItem()

    #读取excel表 并返回excel表的python对象
    def __loadXls(self,fileName):
        xlsData = xlrd.open_workbook(fileName)
        return xlsData

    #读取某一个sheet表
    def __loadTable(self,sheetIndex = 0):
        self.table = self.xlsData.sheet_by_index(sheetIndex)

    #获取sheet表的某一行,并返回
    def __rowValues(self,rowNumber = 0):
        return self.table.row_values(rowNumber)

    #获取某一列的值,并返回
    def __colValues(self,colNumber = 0):
        return self.table.col_values(colNumber)

    #获取具体某个单元格的值 col:列  row:行
    def __cellValue(self, col, row):
        return self.table.cell_value(row,col)

    # 读取exl表中的工作项
    def __loadWorkItem(self):
        workItems = self.__colValues(0)
        itemT = None
        for index, item in enumerate(workItems):
            if index != 0 :
                if item != '':
                    print item
                    itemT = item
                    self.contentModel.addWorkItem(item)

    # 读取exl表中的完成情况
    def __loadCompleteStatus(self):
        completeStatus = self.__colValues(0)
        pass

    # 读取exl表中的风险
    def __loadWarning(self):
        pass

    # 读取exl表中的负责人
    def __loadCharge(self):
        pass

    # 读取exl表中的下周工作计划
    def loadNextWorkPlan(self):
        pass
