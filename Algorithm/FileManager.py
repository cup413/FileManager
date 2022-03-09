import os
import os.path as pt

from win32com import client as wc
import time

import docx

import pandas as pd

import threading

import os

import re


class FileManager:
    def __init__(self, path):
        self.allpath = self.scanFile(path)

    def scanFile(self, path):
        result = []
        for root, dirs, files in os.walk(path):
            for f in files:
                file_path = pt.abspath(pt.join(root, f))

                result.append(file_path)  # 保存路径与盘符

        return result

    def findAllPath(self, s):

        lst = []
        for i in self.allpath:
            if s in i:
                lst.append(i)
        #         for i in range(len(lst)):
        #             print(i, lst[i])

        return lst