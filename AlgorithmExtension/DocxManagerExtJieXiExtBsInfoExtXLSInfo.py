from AlgorithmExtension.DocxManagerExtJieXiExtBsInfo import DocxManagerExJieXiExtBsInfo

from Algorithm.XLS2XLSX import  XLS2XLSX

# -*- coding:UTF-8 -*-
from random import choice
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from openpyxl import load_workbook

import os

class DocxManagerExJieXiExtBsInfoExtXLSInfo(DocxManagerExJieXiExtBsInfo):

    xls2xlsx = XLS2XLSX()

    def getInfoFromXLS(self, path):
        print('==============')
        self.info = ['T' for i in range(7)]

        ed = path.split('.')[-1]
        # print(ed)
        if ed == 'xls':
            if not os.path.exists(path+'x'):
                print('==============')
                self.xls2xlsx.tras(path)
                print('==============')
            path = path +'x'
        elif ed=='xlsx':
            pass
        else:
            print('not xls or xlsx file, wrong')
            return False

        print('==============')
        #打开excel文件，如有公式，读取公式计算结果
        wb=load_workbook(path,data_only=True)

        #遍历EXCEL 文件中所有的worksheet
        for ws in wb.worksheets:
            rows=list(ws.rows)
            #根据worKsheet的行数和列数，在word 文件中创建核实大小的表格

            #从worksheet读取数据，写入word文件中的表格
            for  irow,row in enumerate(rows):
                for icol,col in enumerate(row):
                    if col.value == None:
                        continue
                    s = str( col.value )

                    if 'X：' in s and self.info[0] == 'T':
                        self.info[0] = s
                    if 'Y：' in s and self.info[1] == 'T':
                        self.info[1] = s
                    if 'H：' in s and self.info[4] == 'T':
                        self.info[4] = s
                    if '补心高' in s and self.info[6] == 'T':
                        self.info[6] = s
        col = ['X', 'Y', 'Lat', 'Lon', 'Hd', 'Hb', 'Hbb']
        for i in range(7):
            print(col[i], self.info[i])

        if self.info[0] == 'T':
            return False
        return True

    def saveinfoFromXLS(self, name, folder = r'D:\李晨星文件夹\项目文件\塔里木程小桂\info\%s.csv'):
        # pth = folder % name
        #
        # if self.info[0] == 'T':
        #     print('info为空，不能保存')
        #     return False
        #
        # col = ['X', 'Y', 'Lat', 'Lon', 'Hd', 'Hb', 'Hbb']
        # f = pd.DataFrame([self.info], columns=col)
        # f.to_csv(pth, index=False)
        #
        # print(pth, '    保存成功')
        # return True
        return super(DocxManagerExJieXiExtBsInfoExtXLSInfo, self).saveInfoFromBsInfoTable(name, folder)
