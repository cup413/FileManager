from Algorithm.FileManager import FileManager
from Algorithm.Doc2Docx import Doc2Docx
from Algorithm.DocxManager import DocxManager

from Algorithm.Manager import Manager

import pandas as pd

# print('initialize')
p = r'D:\李晨星文件夹\项目文件\塔里木程小桂'
manager = Manager(p)
# print('over')



#############
tmppath = r'D:\李晨星文件夹\项目文件\塔里木程小桂\data\abcd.xlsx'
data = pd.read_excel(tmppath)

allname = data['well'].unique()

allname


####################

idx = 26
name = allname[idx]
print(name)

# name = '顺'

manager.findFile( name )

path_idx = 0
manager.checkInfo( path_idx)
#
path = manager.returnPath()
manager.getInfoFromText(path)
# manager.saveInfo(name)

# manager.getInfo(path)
# # manager.saveInfo( name)

# # path=r'C:\Users\HP\Desktop\tmp\富贵.docx'
manager.getLayer(path)
# manager.saveLayer( name)

# path = r'D:\李晨星文件夹\项目文件\塔里木程小桂\data\塔里木单井资料\星火1井\星火1井地层简表.doc'
# manager.getLayer(path)
# manager.saveLayer('顺北1-2')