import configparser
import os
import time
from datetime import datetime
from tkinter import Tk, messagebox

import numpy as np
import openpyxl as excel
import pandas as pd
import xlrd

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import MoveTargetOutOfBoundsException
from selenium.common.exceptions import TimeoutException,StaleElementReferenceException

import traceback

SCREEN_SHOT_NAME='screenShot_{0}.png'

# 各画面の処理の共通クラス各画面のクラスはこのクラスを継承する
class SeleniumOperationBase:

    driver=None
    wait=None
    log=None
    # スクリーンショットを格納するディレクトリ
    screenShotBaseName=None

    # 初期化処理
    def __init__(self,driver,log,screenShotBaseName='screenShot'):
        self.driver=driver
        # 画面描画の待ち時間
        self.wait=WebDriverWait(self.driver,20)
        driver.implicitly_wait(30)
        self.log=log
        self.screenShotBaseName=screenShotBaseName+'/'+datetime.now().strftime("%Y%m%d%H%M%S")+'/'
        os.makedirs(self.screenShotBaseName,exist_ok=True)
    
    # スクリーンショットの格納フォルダ名称を修正する
    def setScreenShotBaseName(self,screenShotBaseName):
        self.screenShotBaseName=screenShotBaseName

    # スクリーンショットの名称を修正する
    def setTagetName(self,targetName):
        self.targetName=targetName

    # 画面要素をクリックする 待ち時間が必要な場合は指定を行う
    # 描画などを待たずに処理を行うときに使う
    def webElementClick(self,webElement,waitTime=0):
        try:
            time.sleep(waitTime)
            self.driver.find_element_by_xpath(webElement).click()
        except StaleElementReferenceException:
            # 画面描画前に押下してしまったときの対応
            time.sleep(1)
            self.driver.find_element_by_xpath(webElement).click()
        except SystemError as err:
            self.log.error('画面要素押下失敗:'+webElement)
            self.log.error('例外発生 {}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise
        
    
    # 画面要素をクリックする 要素が表示されるまで20秒待つ さらに待ち時間が必要な場合は指定を行う
    def webElementClickWaitDisplay(self,webElement,waitTime=0):
        try:
            time.sleep(waitTime)
            self.wait.until(expected_conditions.element_to_be_clickable((By.XPATH,webElement)))
            self.driver.find_element_by_xpath(webElement).click()
        except StaleElementReferenceException:
            # 画面描画前に押下してしまったときの対応
            time.sleep(1)
            self.driver.find_element_by_xpath(webElement).click()
        except SystemError as err:
            self.log.error('画面要素押下失敗:'+webElement)
            self.log.error('例外発生 {}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise

    # 画面要素をクリックしてマウスを移動させる クリックした後のメニューの操作を行いたいときに使用する
    # 待ち時間が必要な場合は指定を行う
    # 画面の要素がクリックできるまでまつ
    def webElementClickAndMoveWaitDisplay(self,webElement,waitTime=0):
        try:
            time.sleep(waitTime)
            self.wait.until(expected_conditions.element_to_be_clickable((By.XPATH,webElement)))
            target = self.driver.find_element_by_xpath(webElement)
            actions = ActionChains(self.driver)
            actions.click(target).move_by_offset(10,10).perform()
        except MoveTargetOutOfBoundsException :
                # 要素が範囲外のときは一度スクロール処理を入れる
                self.moveScroll(webElement)
                time.sleep(1)
                actions.click(target).move_by_offset(10,10).perform()
        except SystemError as err:
            self.log.error('画面要素押下失敗:'+webElement)
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise

    # 画面要素をクリックしてマウスを移動させる クリックした後のメニューの操作を行いたいときに使用する
    # 待ち時間が必要な場合は指定を行う
    def webElementClickAndMove(self,webElement,waitTime=0):
        try:
            time.sleep(waitTime)
            target = self.driver.find_element_by_xpath(webElement)
            actions = ActionChains(self.driver)
            actions.click(target).move_by_offset(10,10).perform()
        except MoveTargetOutOfBoundsException :
                # 要素が範囲外のときは一度スクロール処理を入れる
                self.moveScroll(webElement)
                time.sleep(1)
                actions.click(target).move_by_offset(10,10).perform()
        except SystemError as err:
            self.log.error('画面要素押下失敗:'+webElement)
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise
    
    # テキストを送る 描画などを待たずに処理を行う
    def sendText(self,webElement,sendTexts,waitTime=0):
        try:
            time.sleep(waitTime)
            self.driver.find_element_by_xpath(webElement).clear()
            self.driver.find_element_by_xpath(webElement).send_keys(sendTexts)
        except SystemError as err:
            self.log.error('テキスト送信失敗:'+webElement)
            self.log.error('例外発生 {}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise

    # テキストを送る 要素が表示されるまで20秒待つ さらに待ち時間が必要な場合は指定を行う
    def sendTextWaitDisplay(self,webElement,sendTexts,waitTime=0):
        try:
            time.sleep(waitTime)
            self.wait.until(expected_conditions.element_to_be_clickable((By.XPATH,webElement)))
            self.driver.find_element_by_xpath(webElement).clear()
            self.driver.find_element_by_xpath(webElement).send_keys(sendTexts)
        except SystemError as err:
            self.log.error('テキスト送信失敗:'+webElement)
            self.log.error('例外発生 {}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise

    # テキストを入力後エンターキーを押す 要素が表示されるまで20秒待つ さらに待ち時間が必要な場合は指定を行う
    def sendTextAndEnterWaitDisplay(self,webElement,sendTexts,waitTime=0):
        try:
            time.sleep(waitTime)
            self.wait.until(expected_conditions.element_to_be_clickable((By.XPATH,webElement)))
            self.driver.find_element_by_xpath(webElement).clear()
            self.driver.find_element_by_xpath(webElement).send_keys(sendTexts,Keys.ENTER)
        except SystemError as err:
            self.log.error('テキスト送信失敗:'+webElement)
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise

    # プルダウンを選択する 要素が表示されるまで20秒待つ さらに待ち時間が必要な場合は指定を行う
    def selectPullDownWaitDisplay(self,webElement,inputValue,waitTime=0):
        try:
            time.sleep(waitTime)
            self.wait.until(expected_conditions.element_to_be_clickable((By.XPATH,webElement)))
            selectList=self.driver.find_element_by_xpath(webElement)
            selectForm=Select(selectList)
            # 画面に表示されるプルダウンのテキストで選択を行う
            selectForm.select_by_visible_text(inputValue)
        except SystemError as err:
            self.log.error('テキスト送信失敗:'+webElement)
            self.log.error('例外発生 {}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise
    
    # GoogleFormの場合プルダウンが通常の方法では選択できないので
    def selectPullDownGoogleForm(self,buttonInfo,target,pullDownPosition):
        self.webElementClickOverlay(buttonInfo)
        time.sleep(3)
        options=self.driver.find_elements_by_class_name("exportSelectPopup")
        contents = options[pullDownPosition].find_elements_by_tag_name('content')
        [i.click() for i in contents if i.text == target]

    # 指定した要素が表示されるまで待機する
    def waitWebElementVisibility(self,webElement,waitTime=0):
        try:
            time.sleep(waitTime)
            self.wait.until(expected_conditions.visibility_of_element_located((By.XPATH,webElement)))
        except SystemError as err:
            self.log.error('要素待機失敗:'+webElement)
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise
    
    # 画面要素が重なっている箇所に対してクリック処理を行う
    def webElementClickOverlay(self,webElement):
        try:
            target = self.driver.find_element_by_xpath(webElement)
            self.driver.execute_script("arguments[0].click();", target)
        except StaleElementReferenceException:
            # 画面描画前に押下してしまったときの対応
            time.sleep(1)
            self.driver.execute_script("arguments[0].click();", webElement)
        except SystemError as err:
            self.log.error('画面要素押下失敗:'+webElement)
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise

    # 画面のすべての要素が読み込まれるまで待機をする
    def waitWebElementsRead(self,waitTime=0):
        try:
            time.sleep(waitTime)
            self.wait.until(expected_conditions.presence_of_all_elements_located)
        except SystemError as err:
            self.log.error('要素待機失敗')
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException('画面読み込み中の異常')
            raise

    # 指定した要素が存在するかを確認する要素が存在するとTrue
    def existenceWebElements(self,webElement,waitTime=0):
        time.sleep(waitTime)
        elements=self.driver.find_elements_by_xpath(webElement)
        if len(elements) ==0:
            return False
        else:
            return True

    # 指定した要素を返す elementsの状態で返す
    def getWebElements(self,webElement,waitTime=0):
        time.sleep(waitTime)
        elements=self.driver.find_elements_by_xpath(webElement)
        return elements


    # スクリーンショットを取得する 任意の名前を付けられるようにする
    def getScreenShot(self,screenShotName='',sleepTime=3):
        time.sleep(sleepTime)
        if screenShotName=='':
            self.driver.save_screenshot(self.screenShotBaseName+'_'+SCREEN_SHOT_NAME.format(datetime.now().strftime("%Y%m%d_%H%M%S")))
        else :
            self.driver.save_screenshot(self.screenShotBaseName+'_'+screenShotName+'_'+SCREEN_SHOT_NAME.format(datetime.now().strftime("%Y%m%d_%H%M%S")))


    # ポップアップを表示する
    def createOkDialog(self,titlemessage,dispWord):
        root=Tk()
        # メッセージボックス表示時に表示されるウィンドウを最小表示にする
        root.withdraw()
        root.attributes('-topmost', True)
        root.withdraw()
        root.lift()
        root.focus_force()
        messagebox.showinfo(titlemessage,dispWord)
        # メッセージボックス表示時に表示されるウィンドウを消す
        root.quit()

    # エラー発生時の注意喚起用のダイアログを作成する必要に応じて呼び出し元でメッセージなどを設定する
    def errorAlertDialog(self,titemessage='エラー発生',dispWord='エラーが発生しています。'):
        root=Tk()
        # メッセージボックス表示時に表示されるウィンドウを最小表示にする
        root.withdraw()
        root.attributes('-topmost', True)
        root.withdraw()
        root.lift()
        root.focus_force()
        messagebox.showwarning(titemessage,dispWord)
        # メッセージボックス表示時に表示されるウィンドウを消す
        root.quit()
    
    # 指定した位置までスクロールを行う
    def moveScroll(self,webElement):
        # スクロール確認の処理 指定した位置までスクロールを行う
        try:
            self.wait.until(expected_conditions.visibility_of_element_located((By.XPATH,webElement)))
            inputLabel=self.driver.find_element_by_xpath(webElement)
            self.driver.execute_script("arguments[0].scrollIntoView();", inputLabel)

        except SystemError as err:
            self.log.error('画面スクロール失敗:'+webElement)
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException(webElement)
            raise

    # 対象の要素までスクロールしてスクリーンショットをとる
    # find_element_by_xpathで指定した後の要素に対して処理を行う
    def moveScrollAndGetScreenShot(self,webElement,screenShotName=''):
        self.driver.execute_script("arguments[0].scrollIntoView();", webElement)
        self.adjustScroll(-10)
        self.getScreenShot(screenShotName)
    # 指定した要素の文字列を返す ラベルを確認するときに使用する
    def getWebElementTextWaitDisplay(self,webElement,waitTime=0):
        try:
            time.sleep(waitTime)
            self.wait.until(expected_conditions.visibility_of_element_located((By.XPATH,webElement)))
            return self.driver.find_element_by_xpath(webElement).text
        except TimeoutException:
            # 要素が取れないときは空文字を返す
            return ''
        except SystemError as err:
            self.log.error('要素取得失敗:'+webElement)
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.outputException('画面読み込み中の異常')
            raise

    # 画面をスクロールさせる
    def adjustScroll(self,offset):
        try:
            self.driver.execute_script("window.scrollTo(0, window.pageYOffset +"+str(offset)+")" )

        except SystemError as err:
            self.log.error('画面スクロール失敗:'+str(offset))
            self.log.error('例外発生 {0}'.format(err))
            self.getScreenShot()
            raise
        except:
            self.log.error(traceback.format_exc())
            self.getScreenShot()
            raise

    # エラー発生時にエラーログを出力する
    def outputException(self,webElement):
            self.log.error('想定外の例外:{0}'.format(webElement))
            self.log.error(traceback.format_exc())
            self.getScreenShot()


# メイン処理
if __name__=="__main__":
    print('')