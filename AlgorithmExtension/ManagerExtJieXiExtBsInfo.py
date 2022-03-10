
from AlgorithmExtension.ManagerExtJieXi import ManagerExtJieXi

from AlgorithmExtension.DocxManagerExtJieXiExtBsInfo import  DocxManagerExJieXiExtBsInfo

class ManagerExtJieXiExtBsInfo(ManagerExtJieXi):
    docxManager = DocxManagerExJieXiExtBsInfo()

    def __init__(self, p):
        ManagerExtJieXi.__init__(self, p)

    def getInfoFromBsInfoTable(self, path):
        if path.split('.')[-1] == 'doc':
            try:
                self.doc2docx.doc2docx( path)
            except:
                pass
            path = path +'x'
        self.docxManager.makeDoc(path)
        return self.docxManager.getInfoFromBsInfoTable()

    def saveInfoFromBsInfoTable(self, name):
        self.docxManager.saveInfoFromBsInfoTable(name)