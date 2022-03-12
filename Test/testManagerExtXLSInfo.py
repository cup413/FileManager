from AlgorithmExtension.ManagerExtJieXiExtBsInfoExtXLSInfo import ManagerExtJieXiExtBsInfoExtXLSInfo

p = r'D:\李晨星文件夹\项目文件\塔里木程小桂'
manager = ManagerExtJieXiExtBsInfoExtXLSInfo(p)



path = r'C:\Users\HP\Desktop\abc\顺北5井基本数据表.xls'
manager.getInfoFromXLS(path)

manager.saveInfoFromXLS('顺北5')