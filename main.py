from Algorithm.FileManager import FileManager
from Algorithm.Doc2Docx import Doc2Docx
from Algorithm.DocxManager import DocxManager

from Algorithm.Manager import Manager

from AlgorithmExtension.ManagerExtJieXi import  ManagerExtJieXi
from AlgorithmExtension.ManagerExtJieXiExtBsInfoExtXLSInfo import  ManagerExtJieXiExtBsInfoExtXLSInfo

import pandas as pd

import os

# print('initialize')
p = r'D:\李晨星文件夹\项目文件'
manager = ManagerExtJieXiExtBsInfoExtXLSInfo(p)
# print('over')



#############
tmppath = r'D:\李晨星文件夹\项目文件\塔里木程小桂\data\abcd.xlsx'
data = pd.read_excel(tmppath)

allname = data['well'].unique()

allname
####################


def check(idx):
    path_idx = idx
    manager.checkInfo( path_idx)
    #
    path = manager.returnPath()

    return path


def process(name, path):
    print('************************deal with info')
    try:
        print('==============================try getInfoFromText')
        if not manager.getInfoFromText(path):
            raise Exception('')
        try:
            manager.saveInfo(name)
        except:
            print('==============================storeInfoByText fail, check if you want to store anyway')
    except:
        print('==============================getInfoFromText fail, try getInfoByBsInfoTable')
        try:
            if not manager.getInfoFromBsInfoTable(path):
                raise Exception()
            try:
                manager.saveInfoFromBsInfoTable(name)
            except:
                print('==============================storeInfoByBsInfoTable fail, check if you want to store anyway')
        except:
            print('==============================getInfoByBsInfoTable fail, try getInfoByTable')
            try:
                if not manager.getInfo(path):
                    raise Exception()
                try:
                    manager.saveInfo(name)
                except:
                    print('==============================storeInfoByTable fail, check if you want to store anyway')
            except:
                print('==============================getInfoByTable fail, check word file')

    print('\n\n')
    print('********************deal with layer')
    try:
        print('==============================try getLayerFromTableDiJie')
        if not manager.getLayer(path):
            raise Exception()
        try:
            manager.saveLayer(name)
        except  Exception as e:
            print(e)
            print('==============================storeLayerByDiJie fail, check if you want to store anyway')
    except:
        print('==============================getLayerFromTableDiJie fail,  try getLayerFromTableJieXi')
        try:
            if not manager.getLayerByJieXi(path):
                raise Exception()
            try:
                manager.saveLayerByJieXi(name)
            except:
                print('==============================storeLayerByJieXi fail, check if you want to store anyway')
        except:
            print('==============================getLayerFromTableJieXi fail, check word file')

def processXLS(name, path):
    try:
        print('===============================try getInfoFromXLS')
        if not manager.getInfoFromXLS(path):
            raise Exception()
        try:
            manager.saveInfoFromXLS(name)
        except:
            print('==============================storeInfoFromXLS fail, check if you want to store anyway')
    except:
        print('getInfoFromXLS fail, check XLS file')
def finalProcess(idx, name, path):
    if path == '':
        path = manager.getPathByIdx(idx)

    print(path)
    ed = path.split('.')[-1]
    print(ed)
    if ed== 'xls' or ed == 'xlsx':
        processXLS(name, path)
    else:
        path = check(idx)
        process(name, path)


idx = 49
name = allname[idx]
print(name)

# name = '跃满25'
# path = r'C:\Users\HP\Desktop\tmp\顺北5井.docx'
manager.findFile( name )

#
allkey = ['%s.*录井总结报告','%s.*完井', '%s.*基本数据表', '%s.*地层分层数据表']
allidx = []
for i in allkey:
    # print( i% name)
    allidx.append( manager.findKeyWrodInFile(i% name))
print(allidx)
#
for idx in allidx:
    if idx == -1:
        continue
    finalProcess(idx, name, '')
    print('****************************************')

info = r'D:\李晨星文件夹\项目文件\塔里木程小桂\info\%s.csv'
layer = r'D:\李晨星文件夹\项目文件\塔里木程小桂\layer\%s.csv'
print('\n\n\n')
if os.path.exists(info%name):
    print(info%name, '      存在')
    df = pd.read_csv(info%name)
    print(df)
print('**************************')
if os.path.exists(layer%name):
    print(layer%name, '      存在')
    df = pd.read_csv(layer%name)
    print(df)

# name = '中古111'
# path = r'D:\李晨星文件夹\项目文件\塔里木程小桂\data\温压水数据\连续测温数据\中古111静梯报告(华油20100106).doc'
# # finalProcess(-1, name, path)
# process(name, path)