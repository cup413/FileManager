
from AlgorithmExtension.ManagerExtJieXiExtBsInfo import ManagerExtJieXiExtBsInfo

from AlgorithmExtension.DocxManagerExtJieXiExtBsInfoExtXLSInfo import  DocxManagerExJieXiExtBsInfoExtXLSInfo

class ManagerExtJieXiExtBsInfoExtXLSInfo(ManagerExtJieXiExtBsInfo):
    docxManager = DocxManagerExJieXiExtBsInfoExtXLSInfo()

    def __init__(self, p):
        ManagerExtJieXiExtBsInfo.__init__(self, p)

    def getInfoFromXLS(self, path):

        return self.docxManager.getInfoFromXLS(path)

    def saveInfoFromXLS(self, name):
        return self.docxManager.saveInfoFromBsInfoTable(name)