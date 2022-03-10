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
        self.info = ['T' for i in range(7)]

        ed = path.split('.')[-1]
        # print(ed)
        if ed == 'xls':
            if not os.path.exists(path+'x'):

                self.xls2xlsx.tras(path)
            path = path +'x'
        elif ed=='xlsx':
            pass
        else:
            print('not xls or xlsx file, wrong')
            return False


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
                        self.info[0] = s
                    if 'H：' in s and self.info[4] == 'T':
                        self.info[0] = s
                    if '补心高' in s and self.info[6] == 'T':
                        self.info[0] = s
        if self.info[0] == 'T':
            return False
        return True

    def saveinfoFromXLS(self, name, foulder = ):
