#! /usr/bin/env python
# _*_ coding:utf-8 _*_

'ExlModel数据处理'

__author__ = 'DarrenW'

import XMExlModel,os,xlwt

class ExlReadHandle(object):

    noteDate = ''

    def __init__(self):
        self.__hasiOS = False
        self.__hasAndroid = False
        self.__hasTest = False
        self.__hasJava = False
        self.__hasYunW = False

    #获取团队名称
    def teamTitle(self,fileName):
        strArray = fileName.split('_')
        if len(strArray) > 2:
            if len(ExlReadHandle.noteDate) < 1:
                date = strArray[2]
                dateArray = date.split('.')
                if len(dateArray) > 1:
                    self.__handleDate(dateArray[0])

            for index, str in enumerate(strArray):
                if index == 1:
                    return str
        else:
            print '<%s> \t error->文件名格式错误,无法解析'%fileName

    #查找某一目录下的xls文件
    def findExlFile(self,filePath = './'):
        exlList = None
        try:
            exlList = os.listdir(filePath)
        except:
            print '<error> -> 路径无效'
            return
        hasExl = False
        exlDic = {}
        for index,exl in enumerate(exlList):
            if exl.find('xls') >= 0:
                teamTitle = self.teamTitle(exl)
                if teamTitle:
                    if teamTitle.find('iOS') >= 0:
                        self.__hasiOS = True
                        exlDic['ios'] = exl
                    elif teamTitle.find('Android') >= 0:
                        self.__hasAndroid = True
                        exlDic['android'] = exl
                    elif teamTitle.find('运维') >= 0:
                        self.__hasYunW = True
                        exlDic['yunwei'] = exl
                    elif teamTitle.find('java') >= 0:
                        self.__hasJava = True
                        exlDic['java'] = exl
                    elif teamTitle.find('测试') >= 0:
                        self.__hasTest = True
                        exlDic['test'] = exl
                hasExl = True
        if hasExl:
            return self.handleExl(exlDic)

    # 处理exl表
    def handleExl(self,exls = None):
        hasError = False
        ios = None
        android = None
        java = None
        yunwei = None
        test = None
        if exls:
            error = None
            try:
                error = 'ios'
                ios = exls['ios']
                error = 'android'
                android = exls['android']
                # error = 'java'
                # java = exls['java']
                # error = 'yunwei'
                # yunwei = exls['yunwei']
                # error = 'test'
                # test = exls['test']
            except:
                hasError = True
                print 'warning -> 找不到"%s"的周报文件'%error
        else:
            print '<error>method:handleExl(self,exls)\t参数"exls"为None'

        if hasError:
            return

        models = []

        if ios:
            iosModel = XMExlModel.XMExlModel(ios)
            iosModel.teamTitle = u'iOS组'
            models.append(iosModel)
        if android:
            androidModel = XMExlModel.XMExlModel(android)
            androidModel.teamTitle = u'Android组'
            models.append(androidModel)
        if java:
            javaModel = XMExlModel.XMExlModel(java)
            javaModel.teamTitle = u'Java组'
            models.append(javaModel)
        if yunwei:
            yunweiModel = XMExlModel.XMExlModel(yunwei)
            yunweiModel.teamTitle = u'运维组'
            models.append(yunweiModel)
        if test:
            testModel = XMExlModel.XMExlModel(test)
            testModel.teamTitle = u'测试组'
            models.append(testModel)
        return models

    def __handleDate(self,noteDate):
        date = noteDate
        dateArray = date.split('-')
        year = None
        month = None
        day = None
        if len(dateArray) < 3:
            return
        for index,s in enumerate(dateArray):
            if index == 0:
                year = int(s)
            elif index == 1:
                month = int(s)
            elif index == 2:
                day = int(s)
        m1 = str(month)
        m2 = str(month)
        d1 = str(day - 4)
        d2 = str(day)
        difD = day-4
        if difD <= 0:
            numDay = 31
            m = month - 1
            if m == 4 or m == 6 or m == 7 or m == 9 or m == 11:
                numDay = 30
            elif m == 2:
                if (year % 4 == 0) and (year % 100 != 0):
                    numDay = 29
                else:
                    numDay = 28
            d1 = str(numDay + difD)
            if m == 0:
                m1 = '12'
            else:
                m1 = str(month-1)

        if int(m1) < 10:
            m1 = '0'+m1
        if int(m2) < 10:
            m2 = '0'+m2
        if int(d1) < 10:
            d1 = '0'+d1
        if int(d2) < 10:
            d2 = '0'+d2
        ExlReadHandle.noteDate = m1+d1+'-'+m2+d2

# 将model写入exl表
class ExlWriteHandle(object):

    def __init__(self, models):
        self.workbook = xlwt.Workbook()
        self.sheet = self.workbook.add_sheet(u"工作内容")
        self.xlsName = '项目周报_招乎团队_戴子奇_%s.xls'%ExlReadHandle.noteDate
        titles = [u'工作类型',u'分类',u'工作项',u'本周完成情况',u'是否遇到问题或风险',u'下周工作计划',u'负责人']
        lastIndex = 0
        # 工作类型
        self.__writeCell(1, 0, u'杭州招乎团队')
        # 写入各栏标题
        for index, title in enumerate(titles):
            self.__writeCell(0, index, title)
        # 填充内容
        for i, model in enumerate(models):
            row = lastIndex + 1
            self.__writeCell(row, 1, model.teamTitle)
            for index, title in enumerate(titles):
                self.__handleModel(row,index,model)
            lastIndex += model.rows

        self.__writeToExl()

    #处理model
    def __handleModel(self, row, col, model):
        items = model.contentModel.getWorkItems()
        if col == 0:
            return
        # 分类
        if col == 1:
            return
        # 工作项
        elif col == 2:
            length = 0
            for index, item in enumerate(items):
                newRow = index + row + length
                self.__writeCell(newRow, col, item)
                length = model.contentModel.getItemLength(item)
        # 本周完成情况
        elif col == 3:
            row1 = row
            for index, item in enumerate(items):
                status = model.contentModel.getWorkItemCompleteStatus(item)
                for i, s in enumerate(status):
                    newRow = row1
                    self.__writeCell(newRow, col, s)
                    row1 += 1

        # 是否遇到问题或风险
        elif col == 4:
            length = 0
            for index, item in enumerate(items):
                warning = model.contentModel.getWorkItemWarning(item)
                newRow = index + row + length
                self.__writeCell(newRow, col, warning)
                length = model.contentModel.getItemLength(item)
        # 下周工作计划
        elif col == 5:
            length = 0
            for index, item in enumerate(items):
                nextWorkPlan = model.contentModel.getWorkItemWorkPlan(item)
                newRow = index + row + length
                self.__writeCell(newRow, col, nextWorkPlan)
                length = model.contentModel.getItemLength(item)
        # 负责人
        elif col == 6:
            length = 0
            for index, item in enumerate(items):
                charge = model.contentModel.getWorkItemCharge(item)
                newRow = index + row + length
                self.__writeCell(newRow, col, charge)
                length = model.contentModel.getItemLength(item)

    # 写入某一单元格数据
    def __writeCell(self,row, col, value):
        self.sheet.write(row,col,value)

    # 生成一个exl表
    def __writeToExl(self):
        self.workbook.save(self.xlsName)