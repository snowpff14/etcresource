from logging import Formatter, handlers, StreamHandler, getLogger, DEBUG,INFO
from datetime import datetime
import configparser
import os

# ログの出力処理
class LoggerObj:
    def __init__(self, name=__name__,logFilename='log_'):
        self.logger = getLogger(name)
        self.logger.setLevel(DEBUG)
        formatter = Formatter("[%(levelname)s]:[%(asctime)s]: %(message)s")
        needoutput=True
        # stdout
        handler = StreamHandler()
        handler.setLevel(DEBUG)
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        logdirectory='logs'
        os.makedirs(logdirectory,exist_ok=True)
        fileNamePath=logdirectory+'/'+logFilename
        if needoutput:
            logFileName=fileNamePath+datetime.now().strftime("%Y%m%d%H%M%S")+'.log'
            # file
            handler = handlers.RotatingFileHandler(filename = logFileName,
                                                maxBytes = 1048576,
                                                backupCount = 3)
            handler.setLevel(DEBUG)
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)

    def debug(self, msg):
        self.logger.debug(msg)

    def info(self, msg):
        self.logger.info(msg)

    def warn(self, msg):
        self.logger.warning(msg)

    def error(self, msg):
        self.logger.error(msg)

    def critical(self, msg):
        self.logger.critical(msg)