from Algorithm.Manager import Manager

from AlgorithmExtension.DocxManagerExtJieXi import DocxManagerLayerExtJieXi

class ManagerExtJieXi(Manager):

    docxManagerByJieXi = DocxManagerLayerExtJieXi()

    def __init__(self, p):
        Manager.__init__(self, p)

    def getLayerByJieXi(self, path):
        if path.split('.')[-1] == 'doc':
            try:
                self.doc2docx.doc2docx( path)
            except:
                pass
            path = path +'x'
        self.docxManagerByJieXi.makeDoc(path)
        return self.docxManagerByJieXi.getLayerByTableJieXi()
    def saveLayerByJieXi(self, name):
        return self.docxManagerByJieXi.saveLayerByTableJieXi(name)