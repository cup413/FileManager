import os
import os.path as pt

from win32com import client as wc
import time

import docx

import pandas as pd

import threading

import os

import re


class Doc2Docx:
    def __init__(self):
        self.w = wc.gencache.EnsureDispatch('kwps.application')

    #     w = wc.gencache.EnsureDispatch('kwps.application')

    def __del___(self):
        self.w.Quit()

    def doc2docx(self, path):
        #     path=r'D:\李晨星文件夹\项目文件\塔里木程小桂\data\完井报告\博孜9\博孜9井钻井井史.doc'
        #     path2=r'D:\李晨星文件夹\项目文件\塔里木程小桂\data\完井报告\博孜9\博孜9井钻a井5.docx'

        #         a = time.time()
        #         # word = client.Dispatch("Word.Application")
        #         w = wc.gencache.EnsureDispatch('kwps.application')
        #         print('wc:', time.time()-a )
        doc = self.w.Documents.Open(path)
        #         print('doc:', time.time()-a )
        doc.SaveAs2(path + 'x', 12)
        #         print('save:', time.time()-a )
        #         print(time.time()-a)
        doc.Close()