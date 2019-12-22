import pandas as pd
import numpy as np
import re
import configparser
from enum import Enum
import os
import glob
iniFile=configparser.ConfigParser()
iniFile.read("resources/appConfig.ini","UTF-8")

class SORT_ENUM(Enum):
    SORT_OFF='0'
    SORT_LENGTH_DESC='1'
    SORT_LENGTH_ASC='2'

class REG_ENUM(Enum):
    REG_TYPE_CONTAIN='0'
    REG_TYPE_FRONT='1'
    REG_TYPE_BACKWARD='2'
    REG_TYPE_EXACT_MATCH='3'



class ExtractReg():

    def createReg(self):
        searchItems=pd.read_excel('resources/検索データ.xlsx')
        sortTypeCode=iniFile.get('info','sortType')

        searchItemArray=np.asarray(searchItems['検索語'])
        sortType=SORT_ENUM(sortTypeCode)
        if sortType==SORT_ENUM.SORT_LENGTH_ASC or sortType==SORT_ENUM.SORT_LENGTH_DESC:
            searchItemIndex=[]
            for item in searchItemArray:
                searchItemIndex.append(len(item))
            searchSeries=pd.Series(searchItemIndex)
            serchItemDataFrame=pd.concat([searchItems['検索語'],searchSeries],axis=1)
            if sortType==SORT_ENUM.SORT_LENGTH_ASC:
                sortItems=serchItemDataFrame.sort_values(0,ascending=True)
            else:
                sortItems=serchItemDataFrame.sort_values(0,ascending=False)
            searchItemArray=np.asarray(sortItems['検索語'])
        regTypeCode=iniFile.get('info','regType')
        regType=REG_ENUM(regTypeCode)
        regStr=''
        for item in searchItemArray:
            if regStr!='':
                regStr=regStr+'|'
            sItem=item
            if REG_ENUM.REG_TYPE_CONTAIN==regType:
                sItem='.*'+item+'.*'
            elif REG_ENUM.REG_TYPE_FRONT==regType:
                sItem=item+'.*'
            elif REG_ENUM.REG_TYPE_BACKWARD==regType:
                sItem='*.'+item
            elif REG_ENUM.REG_TYPE_EXACT_MATCH==regType:
                sItem=item
            regStr=regStr+sItem
        return re.compile(regStr)


    def extract(self):
        reg=self.createReg()
        paths=glob.glob('data/*.csv')
        
        fileDict={}

        for pathName in paths:
            extractList=[]
            with open(pathName,encoding=iniFile.get('info','encoding')) as f:
                # targetStrs=f.read()
                for targetStr in f:
                    extractStr=reg.search(targetStr)
                    if extractStr:
                        extractList.append(targetStr)
            fileDict[os.path.basename(pathName)]=extractList
        outputPath=iniFile.get('info','outputPath')
        for key,data in fileDict.items():
            outputFile=outputPath+'extract_'+key+'.txt'
            with open(outputFile,encoding='utf-8',mode='w') as f:
                for d in data:
                    f.write(d)

if __name__ == "__main__":
    extract=ExtractReg()
    extract.extract()
    pass
