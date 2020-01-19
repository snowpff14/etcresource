import os
import time
from datetime import datetime
from tkinter import Tk, messagebox

import numpy as np
import openpyxl as excel
import pandas as pd
import xlrd
import configparser

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select, WebDriverWait
from seleniumOperationBase import SeleniumOperationBase
from utils.logger import LoggerObj

# driver=webdriver.Chrome('C:/webdrivers/chromedriver.exe')
#driver=webdriver.Firefox('C:/webdrivers/geckodriver.exe')
#driver=webdriver.Firefox()
driver=None
# ログイン情報などを取得
iniFile=configparser.ConfigParser()
iniFile.read("resources/appConfig.ini","UTF-8")


TARGET_URL=iniFile.get('info','url')

RESERVE_YEAR='//div/div/form/input[1]'
RESERVE_MONTH='//div/div/form/input[2]'
RESERVE_DAY='//div/div/form/input[3]'
REST_DAY='//div/div/form/input[4]'

NUMBER_PEOPLE='//div/div/form/input[5]'

NEED_MORNING='//div/div/form/input[{0}]'

PLAN_SELECT='//div/div/form/input[{0}]'

NAME='//div/div/form/input[10]'

MORNING_TYPE={'あり':6,'なし':7}
PLAN_TYPE={'昼からチェックインプラン':8,'お得な観光プラン':9}

NEXT_BUTTON='//div/div/form/button'

PAY_DETAIL='/html/body/div[1]/div/form/div/div/a'
BACK_BUTTON='//div/div/button'
CONFIRM_BUTTON='//div/div/form/button'  
RESULT_BACK_BUTTON='//div/div/button'


class TestSiteOrder(SeleniumOperationBase):

    def __init__(self,driver,log,screenShotBaseName='screenShotName'):
        super().__init__(driver,log,screenShotBaseName)
    
    def pullDownSelect(self,webElement,inputText):
        target=self.driver.find_element_by_xpath(webElement)
        selecttargetForm=Select(target)
        # 画面に表示されるプルダウンのテキストで選択を行う
        selecttargetForm.select_by_visible_text(inputText.zfill(2))

    def inputOrder(self,reserveSheet):

        reserveSheetDict=reserveSheet.to_dict('index')

        for index,data in reserveSheetDict.items():

            number=data['項番']
            if number !=number:
                # 項番がない状態であれば登録処理を終了
                break
            day=data['宿泊日']
            year=day.split('/')[0]
            month=day.split('/')[1]
            restDay=day.split('/')[2]
            visitDay=data['宿泊数']
            numberOfPeople=data['人数']
            morningType=data['朝食バイキング']
            plan=data['プラン']
            name=data['名前']
            remark=data['備考']

            
            super().sendTextWaitDisplay(RESERVE_YEAR,year)
            super().sendTextWaitDisplay(RESERVE_MONTH,month)
            super().sendTextWaitDisplay(RESERVE_DAY,restDay)

            super().sendTextWaitDisplay(REST_DAY,visitDay)

            super().sendTextWaitDisplay(NUMBER_PEOPLE,numberOfPeople)

            super().webElementClick(NEED_MORNING.format(MORNING_TYPE[morningType]))
            super().webElementClick(PLAN_SELECT.format(PLAN_TYPE[plan]))
            super().sendText(NAME,name)
            super().getScreenShot(screenShotName=name+'_入力画面')
            super().webElementClick(NEXT_BUTTON)

            # 次の画面に遷移後
            super().waitWebElementVisibility(PAY_DETAIL)
            super().getScreenShot(screenShotName=name+'_確認画面')
            super().webElementClick(PAY_DETAIL)
            # アラートはスクリーンショットがとれないのでコメントアウト
            # super().getScreenShot(screenShotName=name+'_確認画面_料金詳細')
            self.driver.switch_to.alert.accept()
            super().webElementClick(CONFIRM_BUTTON)

            # 完了画面
            super().getScreenShot(screenShotName=name+'_完了画面',sleepTime=5)

            # 初期画面に戻る
            super().webElementClickWaitDisplay(RESULT_BACK_BUTTON)
            super().webElementClickWaitDisplay(BACK_BUTTON)



        

# メイン処理
if __name__=="__main__":
    log=LoggerObj()
    driver.get(TARGET_URL)
    excelFile=pd.ExcelFile('data/予約データ.xlsx')
    reserveSheetTemp=excelFile.parse(sheet_name='予約シート',dtype='str',header=1)
    print(reserveSheetTemp.head())

    reserveSheet=reserveSheetTemp.query('無効フラグ != "1"')

    testSideOrder=TestSiteOrder(driver,log,'test')
    # 勤務時間入力
    testSideOrder.inputOrder(reserveSheet)
    time.sleep(2)

    testSideOrder.createOkDialog('処理完了','登録処理完了')


    #driver.close()



