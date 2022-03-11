from Algorithm.FileManager import FileManager
from Algorithm.Doc2Docx import Doc2Docx
from Algorithm.DocxManager import DocxManager

from Algorithm.Manager import Manager

from AlgorithmExtension.ManagerExtJieXi import  ManagerExtJieXi
from AlgorithmExtension.ManagerExtJieXiExtBsInfoExtXLSInfo import  ManagerExtJieXiExtBsInfoExtXLSInfo

import pandas as pd

# print('initialize')
p = r'D:\李晨星文件夹\项目文件\塔里木程小桂'
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
        if manager.getInfoFromXLS(path):
            raise Exception()
        try:
            manager.saveInfoFromXLS(name)
        except:
            print('==============================storeInfoFromXLS fail, check if you want to store anyway')
    except:
        print('getInfoFromXLS fail, check XLS file')
def finalProcess():
    ed = path.split('.')[-1]
    if ed== 'xls' or ed == 'xlsx':
        processXLS()
    else:
        process(name, path)


idx = 31
name = allname[idx]
print(name)

# name = '顺北5'
# path = r'C:\Users\HP\Desktop\tmp\顺北5井.docx'
manager.findFile( name )

idx = 0
path = check(idx)
finalProcess()
