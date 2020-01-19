import pandas as pd
import numpy as numpy
from utils.logger import LoggerObj
import sys
import os
import tkinter
from excelFileOperate import DataInsertAndDelete
from seleniumTestSite1 import TestSiteOrder
from tkinter import messagebox
from tkinter import filedialog
from tkinter import Button,ttk,StringVar
from selenium import webdriver



class PythonGui():
    def init(self):
        pass

    def fileSelect(self,event):
        fTyp = [("","*.xlsx")]
        iDir = os.path.abspath(os.path.dirname(__file__))
        filename = filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
        excelFile=pd.ExcelFile(filename)
        reserveSheetTemp=excelFile.parse(sheet_name='予約シート',dtype='str',header=1)
        print(reserveSheetTemp.head())
        log=LoggerObj()
        driver=webdriver.Chrome('C:/webdrivers/chromedriver.exe')
        driver.get('http://example.selenium.jp/reserveApp/')

        reserveSheet=reserveSheetTemp.query('無効フラグ != "1"')
        testSideOrder=TestSiteOrder(driver,log,'test')
        # 勤務時間入力
        testSideOrder.inputOrder(reserveSheet)

        testSideOrder.createOkDialog('処理完了','登録処理完了')
    
    def fileInsert(self,event):
        dataInsertAndDelete=DataInsertAndDelete()
        dataInsertAndDelete.insertCell()

    def fileDelete(self,event):
        dataInsertAndDelete=DataInsertAndDelete()
        dataInsertAndDelete.insertCell()

    def main(self):
        root= tkinter.Tk()
        root.title(u"Python GUI")
        root.geometry("400x300")


        #ラベル
        Static1 = tkinter.Label(text=u'test')
        Static1.pack()
        button1 = Button(text=u'ファイル追加処理', width=20)

        button1.bind("<Button-1>", self.fileInsert)
        button1.pack()

        button2 = Button(text=u'ファイル削除処理', width=20)

        button2.bind("<Button-1>", self.fileDelete)
        button2.pack()

        button3 = Button(text=u'自動処理', width=20)
        button3.bind("<Button-1>", self.fileSelect)
        button3.pack()
        

        # 入力欄
        entry1=tkinter.Entry()
        entry1.pack()

        


        root.mainloop()



if  __name__ == "__main__":
    pythonGui=PythonGui()

    pythonGui.main()
