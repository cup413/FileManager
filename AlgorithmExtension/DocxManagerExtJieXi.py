from Algorithm.DocxManager import DocxManager

import pandas as pd

class DocxManagerLayerExtJieXi(DocxManager):
    layer = ''

    def makeDoc(self, path):
        super(DocxManagerLayerExtJieXi, self).makeDoc(path)
        self.layer = ''

    def checkLayerInTablesByJieXi(self, tables):
        '''
        根据界-系检查文档中是否存在地层分层
        :param tables: 文档中的表格
        :return:
        '''
        for table in tables:
            try:
                s1 = table.cell(0,0).text
                # s2 = table.cell(1,0).text
                s3 = table.cell(0,1).text
            except:
                continue

            if '界' in s1 and '系' in s3 :
                return True
        return False

    def _getLayerLstFromTable(self, tables):
        # print('==========find layer================')
        flst = []
        for table in tables:
            lst = []
            col = len(table.columns)

            # print(table.cell(0, 0).text)
            s = ''
            for i in table.cell(0, 0).text:
                if i != ' ':
                    s = s + i
            try:
                table.cell(0,1).text
            except:
                continue

            # for i in range(col):
            #     print( table.cell(0,i).text )

            # print('333333333333333333333333')
            s2 = table.cell(0,1).text
            if ('界' in s) and ( '系' in s2 ):

                for i in range(0, len(table.rows)):  # 从表格第二行开始循环读取表格数据
                    alst = []
                    for j in range(col):
                        alst.append(table.cell(i, j).text)
                        # print(table.cell(i, j).text, end="  ")
                    lst.append(alst)
                    # print()
                # print('===========lst over')

                if len(lst)>len(flst):
                    flst = lst

        return flst
    def _getLstXiDiShenFromLst(self, lst ):
        def getDeepth(s):
            if ('底' in s and '深' in s) or ('井' in s and '深' in s):
                return True
            return False

        def getXi(s):
            if '系' in s:
                return True
            return False

        def getidx(lst, isTrue):
            # print(lst[0])
            idx = 0
            for i in range(len(lst[0])):

                if isTrue(lst[0][i]):
                    idx = i
                    break
            return idx
        # lst = self.getLayerLstFromTable()

        # print('========lst')
        # print(lst)

        idx = getidx(lst, getDeepth)
        idxXi = getidx(lst, getXi)
        # print('idx, idxXi', idx, idxXi)

        ret = []
        for i in lst:
            #             print(i[0],i[1], i[idx])
            ret.append([i[idxXi], i[idx]])

        return ret

    def _dropDuplicatesInLst(self, ret):

        # ret = self.getLayer()

        myret = []
        for i in ret:
            if i[1] == '':
                pass
            else:
                myret.append(i)
        ret = myret

        alreadyIn = set()
        fret = []
        for i in range(len(ret) - 1, -1, -1):
            if ret[i][0] in alreadyIn:
                continue

            alreadyIn.add(ret[i][0])
            fret.append([ret[i][0], ret[i][1]])
            print([ret[i][0], ret[i][1]])
        # print('==============')
        fret.reverse()

        return fret

    def getLayerByTableJieXi(self):
        if not self.checkLayerInTablesByJieXi(self.tables):
            print('地层分层界系表不存在')
            return False
        tables = self._getLayerLstFromTable(self.tables)
        # print('=======================tables')
        # print(tables)
        lst = self._getLstXiDiShenFromLst(tables)
        self.layer = self._dropDuplicatesInLst(lst)

        if len(self.layer) == 0:
            return False
        return True

    def saveLayerByTableJieXi(self, name, folder = r'D:\李晨星文件夹\项目文件\塔里木程小桂\layer\%s.csv'):
        if self.layer == '':
            print('地层分层界系表没有赋值')
            return False

        pth = folder % name

        f = pd.DataFrame(self.layer[1:], columns=['层', '底界深度'])
        f.to_csv(pth, index=False)

        print(pth, '   保存成功')
        return True