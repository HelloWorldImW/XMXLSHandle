#! /usr/bin/env python
# _*_ coding:utf-8 _*_

'ExlModel数据处理'

__author__ = 'DarrenW'

import XMExlModel,os,xlwt

class ExlReadHandle(object):

    def __init__(self):
        self.__hasiOS = False
        self.__hasAndroid = False
        self.__hasTest = False
        self.__hasJava = False
        self.__hasYunW = False
        self.__noteDate = ''

    #获取团队名称
    def teamTitle(self,fileName):
        strArray = fileName.split('_')
        if len(strArray) > 2:
            if len(self.__noteDate) < 1:
                date = strArray[2]
                dateArray = date.split('.')
                if len(dateArray) > 1:
                    self.__noteDate = dateArray[0]

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
            models.append(iosModel)
        if android:
            androidModel = XMExlModel.XMExlModel(android)
            models.append(androidModel)
        if java:
            javaModel = XMExlModel.XMExlModel(java)
            models.append(javaModel)
        if yunwei:
            yunweiModel = XMExlModel.XMExlModel(yunwei)
            models.append(yunweiModel)
        if test:
            testModel = XMExlModel.XMExlModel(test)
            models.append(testModel)
        return models

    # def handleModel(self,model):
    #     aa = model.contentModel.getWorkItems()
    #     for index, ss in enumerate(aa):
    #         if index == 0:
    #             print ss
    #             print '\n'
    #             status = model.contentModel.getWorkItemCompleteStatus(ss)
    #             for i, s in enumerate(status):
    #                 print i, s
    #             print '\n'
    #             warning = model.contentModel.getWorkItemWarning(ss)
    #             print warning
    #             print '\n'
    #             nextWorkPlan = model.contentModel.getWorkItemWorkPlan(ss)
    #             print nextWorkPlan
    #             print '\n'
    #             charge = model.contentModel.getWorkItemCharge(ss)
    #             print charge


class ExlWriteHandle(object):

    def __init__(self, models):
        self.workbook = xlwt.Workbook()
        self.sheet = self.workbook.add_sheet(u"工作内容")
        self.xlsName = '项目周报_招呼团队_戴子奇_0116-0120.xls'
        self.writeModel(models)

    def writeModel(self,models):
        for index, model in enumerate(models):
            print index,models
        self.writeToExl()

    def writeToExl(self):
        self.workbook.save(self.xlsName)