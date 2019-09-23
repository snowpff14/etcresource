import pandas as pd
import configparser
from pathlib import Path
import csv
from logger import LoggerObj 

iniFile=configparser.ConfigParser()
iniFile.read("resources/appConfig.ini","UTF-8")

class InputDataError(Exception):
    def __init__(self,message):
        self.message=message

class FileMerge():
  # CSVファイルをソートして取得する
    def getSortedCSVFiles(self,filePath):
        targetFilePath=Path(filePath)
        
        return sorted(targetFilePath.glob('*.csv'))

    # ファイルのマージを行う
    def fileMerge(self,targetDir,outputFileName):
        log=LoggerObj('fileMerge')
        log.info('処理開始')


        files=self.getSortedCSVFiles(targetDir)
        if len(files)==0:
            raise InputDataError(targetDir+' is empty')
        firstFile=files[0]
        encodeType=iniFile.get('encode','encodeType')

        dummy=pd.read_csv(firstFile, encoding=encodeType, dtype='str',engine='python')

        datas=pd.DataFrame(columns=dummy.columns)
        for file in files:
            log.info(str(file)+'開始')
            dataTemps = pd.read_csv(file, encoding=encodeType, dtype='str',engine='python')
            datas=pd.concat([datas,dataTemps])
            log.info(str(file)+'終了')
        
        # 重複したデータがあれば削除する
        datas=datas.drop_duplicates()
        datas.to_csv(outputFileName,index=None,encoding=iniFile.get('encode','outputEncodeType'),quoting=csv.QUOTE_ALL)
        log.info('処理終了')

if __name__ == "__main__":
    fileMerge =FileMerge()
    fileMerge.fileMerge(iniFile.get('file','inputFileDirectory'),iniFile.get('file','outputFile'))

