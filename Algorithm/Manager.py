
import os
import os.path as pt

from win32com import client as wc
import time

import docx

import pandas as pd

import threading

import os

import re


from Algorithm.FileManager import FileManager
from Algorithm.Doc2Docx import Doc2Docx
from Algorithm.DocxManager import DocxManager



class Manager:
    # p = r'D:\李晨星文件夹\项目文件\塔里木程小桂\data'

    doc2docx = Doc2Docx()

    docManager = DocxManager()

    def __init__(self, p):
        self.fileManager = FileManager(p)

    #     et = wc.Dispatch("et.Application")
    #     et.Visible = True # 确定ET是否可见

    et = wc.gencache.EnsureDispatch('kwps.application')

    def findFile(self, path):
        endby = ['doc', 'docx', 'xls', 'xlsx']
        file = self.fileManager.findAllPath( path )

        self.file =[]
        for i in file:
            if '$' in i:
                continue
            en = i.split('.')[-1]
            if en not in endby:
                continue

            reg = path + '[0-9]*'
            tmp = re.search(reg, i)
            if tmp.group() != path :
                continue


            # print(re.search(reg, i))

            self.file.append( i )



        for i in range(len(self.file)):
            print(i ,': ', self.file[i])

    def findKeyWrodInFile(self, key):
        # print(self.file)
        for i in range(len(self.file)):
            # if key in self.file[i]:
            #     print(key, ' in ', i)
            #     return i
            # print(key)
            # print(type(key))
            # key = str( key )
            # print(key, self.file[i])
            tmp = re.search(key, self.file[i])
            if tmp != None:
                # print(key, '  in ', i)
                return i
        return -1

    def checkInfo(self, path_idx):
        def action(path):
            os.system('wps %s ' %path)

        path = self.file[path_idx]
        print(path)

        if path.split('.')[-1] == 'doc':
            try:
                self.doc2docx.doc2docx( path)
            except:
                pass
            path = path +'x'

        self.et.Visible =True
        doc = self.et.Documents.Add(path)


        self.path = path
        self.docManager.makeDoc( path)
        #         print('asdfasd')
        self.docManager.checkInfo()
    #         print('asdfasd')

    def returnPath(self):
        return self.path
    def getPathByIdx(self, idx):
        return self.file[idx]

    def disp(self, path_idx):
        def action(path):
            os.system('wps %s ' %path)

        path = self.file[path_idx]
        #         print(path)

        if path.split('.')[-1] == 'doc':
            try:
                self.doc2docx.doc2docx( path)
            except:
                pass
            path = path +'x'

        self.et.Visible =True
        doc = self.et.Documents.Add(path)

    def getInfo(self, path):
        if path.split('.')[-1] == 'doc':
            try:
                self.doc2docx.doc2docx( path)
            except:
                pass
            path = path +'x'

        self.docManager.makeDoc( path)
        return self.docManager.getInfoFromTable()

    def getInfoFromText(self, path):
        if path.split('.')[-1] == 'doc':
            try:
                self.doc2docx.doc2docx( path)
            except:
                pass
            path = path +'x'

        self.docManager.makeDoc(path)
        return self.docManager.getInfoFromText()

    def saveInfo(self, path):
        self.docManager.saveInfo(path)

    def getLayer(self, path):
        if path.split('.')[-1] == 'doc':
            try:
                self.doc2docx.doc2docx( path)
            except:
                pass
            path = path +'x'

        self.docManager.makeDoc(path)
        return self.docManager.getLayer1()
    def saveLayer( self, name):
        return self.docManager.saveLayer(name)