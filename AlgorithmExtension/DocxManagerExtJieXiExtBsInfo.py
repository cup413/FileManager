from AlgorithmExtension.DocxManagerExtJieXi import DocxManagerLayerExtJieXi

import pandas as pd


class DocxManagerExJieXiExtBsInfo(DocxManagerLayerExtJieXi):

    def makeDoc(self, path):
        super(DocxManagerLayerExtJieXi,self).makeDoc(path)

        self.info = ['T' for i in range(7)]

    def checkInfoInBsInfoTable(self, tables):
        for table in tables:
            col = len(table.columns)

            s = table.cell(0, 0).text
            s = self.toolDropSpace(s)

            if '基本数据表' not in s:
                continue
            return True

        return False

    def _getInfoLstFromBsInfoTable(self, tables):
        mytable = []
        for table in tables:
            col = len(table.columns)

            s = table.cell(0, 0).text
            s = self.toolDropSpace(s)

            if '基本数据表' not in s:
                continue

            for i in range(len(table.rows)):

                atable = []
                for j in range(col):
                    atable.append(table.cell(i, j).text)
                mytable.append(atable)
        # print('=============')
        return mytable

    def _getInfoFromTable(self, table):

        def addData(p, i, j, atable):
            # print(p,atable[i][j])
            if info[p] == 'T':
                info[p] = atable[i][j]

        info = ['T' for i in range(7)]
        for i in range(len(table)):
            for j in range(len(table[i])):
                if ('X：' in table[i][j]) :
                    addData(0, i, j, table)

                if ('Y：' in table[i][j]) :
                    addData(1, i, j, table)

                if 'H：' in table[i][j] :
                    addData(4, i, j, table)

                if '补心高' in table[i][j]:
                    addData(6, i, j, table)
        #                     self.info[6] = table[i][j+1]
        lst = ['X', 'Y', 'Lat', 'Lon', 'Hd', 'Hb', 'Hbb']
        for i in range(len(lst)):
            print(lst[i], info[i])

        return info

    def getInfoFromBsInfoTable(self):

        if not self.checkInfoInBsInfoTable(self.tables):
            print('不存在基本数据表')
            return

        # print('asdf')
        table = self._getInfoLstFromBsInfoTable(self.tables)
        self.info = self._getInfoFromTable(table)

        if self.info[0] == 'T':
            return False
        return True

    def saveInfoFromBsInfoTable(self, name, folder = r'D:\李晨星文件夹\项目文件\塔里木程小桂\info\%s.csv'):
        pth = folder % name

        if self.info[0] == 'T':
            print('info为空，不能保存')
            return False

        col = ['X', 'Y', 'Lat', 'Lon', 'Hd', 'Hb', 'Hbb']
        f = pd.DataFrame([self.info], columns=col)
        f.to_csv(pth, index=False)

        print(pth, '    保存成功')
        return True