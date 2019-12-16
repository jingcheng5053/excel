# -*- coding:utf-8 -*-
# !/usr/bin/env python

import os, io
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.pdfgen import canvas
from tkinter import *
import time
from PIL import Image, ImageTk
import numpy as np

# 图片文件名称列表
IMAGEFILES = []


class pdfTk(object):

    def __init__(self):
        '''用于生成主界面用于填写'''
        self.top = Tk()
        self.sw = self.top.winfo_screenwidth()
        self.sh = self.top.winfo_screenheight()
        self.topw = 500
        self.toph = 200
        self.top.title('图片转pdf生成器')
        self.top.geometry("%dx%d+%d+%d" % (self.topw, self.toph, (self.sw - self.topw) / 2, (self.sh - self.toph) / 2))

        self._DIRPATH = StringVar(self.top)

        self.emptfmone = Frame(self.top, height=50)
        self.emptfmone.pack()

        self.dirfm = Frame(self.top)
        self.descriptLabel = Label(self.dirfm, width=4, text='路径：')
        self.descriptLabel.pack(side=LEFT)
        self.dirn = Entry(self.dirfm, width=50, textvariable=self._DIRPATH)
        # self.dirn.bind('<Return>', self.setPath)
        self.dirn.pack(side=LEFT)
        self.dirfm.pack()

        self.emptfmtwo = Frame(self.top, height=30)
        self.emptfmtwo.pack()

        self.btnfm = Frame(self.top)
        self.converBtn = Button(self.btnfm, width=10, text='生成PDF', command=self.doneAnyThing,
                                activeforeground='white', activebackground='blue')
        self.quitBtn = Button(self.btnfm, width=10, text='退出', command=self.top.quit, activeforeground='white',
                              activebackground='blue')
        self.converBtn.pack(side=LEFT, padx=10)
        self.quitBtn.pack(side=LEFT, padx=10)
        self.btnfm.pack()

    def doneAnyThing(self):
        self.getListImages(self._DIRPATH.get())
        pdfFile = self.converPath(self._DIRPATH.get()) + self.dateStr() + ".pdf"
        self.convertpdf(pdfFile)

    def convertpdf(self, pdfFile):
        '''多个图片合成一个pdf文件'''
        # (w, h) = landscape(A4)  # 横向A4
        (w, h) = portrait(A4)  # 纵向A4
        cv = canvas.Canvas(pdfFile, pagesize=portrait(A4))
        for imagePath in IMAGEFILES:
            # img_open = Image.open(imagePath)
            # img_open.show()
            # print(type(img_open))
            # img = np.array(img_open, dtype=np.uint8)
            # img = ImageTk.PhotoImage(img_open)
            # mf = io.BytesIO()
            # mf.write(img)
            # mf.seek(0)
            # PhotoImage.write("222", format="jpeg")
            # print(type(img), img)
            cv.drawInlineImage(imagePath, 0, 0, w, h)
            # cv.drawImage(img_open, 0, 0, w, h)
            cv.showPage()
        cv.save()

    def getListImages(self, dirPath):
        '''读取指定文件夹下所有的JPEG图片，存入列表'''
        if dirPath is None or len(dirPath) == 0:
            raise ValueError('dirPath不能为空，该值为存放图片的具体路径文件夹！')
        if os.path.isfile(dirPath):
            raise ValueError('dirPath不能为具体文件，该值为存放图片的具体路径文件夹！')
        if os.path.isdir(dirPath):
            for imageName in os.listdir(dirPath):
                if imageName.endswith('.jpg') or imageName.endswith('.jpeg'):
                    absPath = self.converPath(dirPath) + imageName
                    print(absPath)
                    IMAGEFILES.append(absPath)

    def converPath(self, dirPath):
        '''用于转换路径，判断路径后是否为\\，如果有则直接输出，如果没有则添加'''
        if dirPath is None or len(dirPath) == 0:
            raise ValueError('dirPath不能为空！')
        if os.path.isfile(dirPath):
            raise ValueError('dirPath不能为具体文件，该值为文件夹路径！')
        if not str(dirPath).endswith("\\"):
            return dirPath + "\\"
        return dirPath

    def dateStr(self):
        '''用于生成指定格式的日期，目的是为了拼接字符串'''
        return time.strftime("%Y-%m-%d", time.localtime())


def main():
    '''该函数主要用于生成PDF文件'''
    pdfTk()
    mainloop()


if __name__ == '__main__':
    '''主函数，进行启动'''
    main()
