import pandas as pd
import numpy as numpy
from utils.logger import LoggerObj
import sys
import os
import tkinter
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from tkinter.ttk import *
import threading
from excelFileOperate import DataInsertAndDelete
from seleniumTestSite1 import TestSiteOrder
from tkinter import messagebox
from tkinter import filedialog
from tkinter import Button,ttk,StringVar
from selenium import webdriver


root= tkinter.Tk()
BUTTON_LABEL_REFERENCE='参照'
class PythonGui():

    inputFileName=StringVar()
    inputval=StringVar()

    inputFolder=StringVar()
    outputFolder=StringVar()

    progressMsg=StringVar()
    progressBar=None
    progressMsgBox=None

    def init(self):
        pass
    

    def execute(self):

        excelFile=pd.ExcelFile(self.inputFileName.get())
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
    
    def filenameDisp(self):
        fTyp = [("","*.xlsx")]
        iDir = os.path.abspath(os.path.dirname(__file__))
        filename = filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
        self.inputFileName.set(filename)

    def openFile(self):
        fTyp = [('','*.xlsx')]
        iDir = os.path.abspath(os.path.dirname(__file__))
        filename = filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
        return filename

    def fileButton(self):
       filename= self.openFile()
       self.inputFileName.set(filename)

    def doExecute(self):
        # threading.Lock
        thread=threading.Thread(target=self.execute)
        thread.start()


    def fileInsert(self):
        dataInsertAndDelete=DataInsertAndDelete()
        dataInsertAndDelete.insertCell()

    def progressMsgSet(self,msg):
        self.progressMsg.set(msg)

    def progressStart(self):
        self.progressBar.start(100)

    def inputResultFolderButton(self):
        dirname = filedialog.askdirectory()
        self.outputFolder.set(dirname)
        
    def fileDelete(self):
        dataInsertAndDelete=DataInsertAndDelete()
        dataInsertAndDelete.insertCell()

    def popUpMsg(self,event):
        tkinter.messagebox.showinfo('inputValue',self.inputval.get())

    def main(self):
        root.title("Python GUI")

        content = ttk.Frame(root)
        frame = ttk.Frame(content,  relief="sunken", width=400, height=500)
        title = ttk.Label(content, text="Python GUI")

        # inputFileName=StringVar()
        # inputval=StringVar()
        content.grid(column=0, row=0)

        title.grid(column=0, row=0, columnspan=4)

        fileLabel=ttk.Label(content,text="予約情報")
        resultFolderLabel=ttk.Label(content,text="フォルダ指定")

        fileInput=ttk.Entry(content,textvariable=self.inputFileName,width=70)
        resultFolderInput=ttk.Entry(content,textvariable=self.outputFolder,width=70)

        self.progressMsgBox=ttk.Label(content,textvariable=self.progressMsg,width=70)

        self.progressBar=ttk.Progressbar(content,orient=HORIZONTAL,length=140,mode='indeterminate')
        self.progressBar.configure(maximum=10,value=0)

        fileInputButton=ttk.Button(content, text=BUTTON_LABEL_REFERENCE,command=self.fileButton)
        resultDirectoryInputButton=ttk.Button(content, text=BUTTON_LABEL_REFERENCE,command=self.inputResultFolderButton)
         
        executeButton=ttk.Button(content,text='実行',command=self.doExecute)
        fileExecuteButton1=ttk.Button(content,text='ファイル操作 挿入実行',command=self.fileInsert)
        fileExecuteButton2=ttk.Button(content,text='ファイル操作 デリート実行',command=self.fileDelete)

        fileLabel.grid(column=1, row=1,sticky='w')
        resultFolderLabel.grid(column=1, row=5,sticky='w')

        executeButton.grid(column=1, row=6,columnspan=3,sticky='we')
        fileExecuteButton1.grid(column=1, row=7,columnspan=3,sticky='we')
        fileExecuteButton2.grid(column=1, row=8,columnspan=3,sticky='we')
        self.progressMsgBox.grid(column=1, row=9,columnspan=3,sticky='we')
        self.progressBar.grid(column=1, row=10,columnspan=3,sticky='we')

        fileInput.grid(column=2, row=1)
        resultFolderInput.grid(column=2, row=5)

        fileInputButton.grid(column=3, row=1)
        resultDirectoryInputButton.grid(column=3, row=5)




        root.mainloop()



if  __name__ == "__main__":
    pythonGui=PythonGui()

    pythonGui.main()
