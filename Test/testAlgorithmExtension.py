# from AlgorithmExtension.DocxManagerExtJieXi import DocxManagerLayerExtJieXi
#
# # docxManager = DocxManagerLayerExtJieXi()
#
#
# # path = r'D:\李晨星文件夹\项目文件\塔里木程小桂\data\完井报告\柯中107\KZ107_地质录井总结报告.docx'
# # # path = r'C:\Users\HP\Desktop\tmp\ab.docx'
# # docxManager.makeDoc(path)
# #
# # docxManager.getLayerByTableJieXi()


from AlgorithmExtension.DocxManagerExtJieXiExtBsInfo import  DocxManagerExJieXiExtBsInfo

from AlgorithmExtension.ManagerExtJieXiExtBsInfo import  ManagerExtJieXiExtBsInfo

# p = r'D:\李晨星文件夹\项目文件\塔里木程小桂'
# manager = ManagerExtJieXiExtBsInfo(p)
#
# path = r'C:\Users\HP\Desktop\tmp\顺北5井.docx'
#+
# name = '顺北5'
# manager.getInfoFromBsInfoTable(path)
#
# manager.saveInfoFromBsInfoTable(name)


from AlgorithmExtension.DocxManagerExtJieXiExtBsInfoExtXLSInfo import DocxManagerExJieXiExtBsInfoExtXLSInfo

# p = r'D:\李晨星文件夹\项目文件\塔里木程小桂'
docxManager = DocxManagerExJieXiExtBsInfoExtXLSInfo()

path = r'C:\Users\HP\Desktop\abc\顺北5井基本数据表.xls'
docxManager.getInfoFromXLS(path)