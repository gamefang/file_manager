# -*- coding: utf-8 -*-

import os
import json
import codecs

from XlManager import XlManager as XlManager

class ConfManager():
    '''
    配置项的加载、读取、存储功能
    '''
    # 基础配置文件路径
    BASE_CONF_FILEPATH = 'FileManager.json'
    # 基础配置默认内容
    BASE_CONF_CONTENT = {
        'EXCEL_FILE_PATH':'FileManager.xlsx',
        'CONF_SHEET_NAME':'config',
        'LIST_SHEET_NAME':'list'
        }

    # 待引用的基础配置 CFG.BASE.XXX
    BASE = None
    # 待引用的EXCEL配置 CFG.EXCEL.XXX
    EXCEL = None

    @classmethod
    def load_basic_cfg(cls):
        '''
        加载或初始化基础配置
        '''
        # 检查文件是否存在，若无则按默认创建
        if not os.path.exists(cls.BASE_CONF_FILEPATH):
            with codecs.open(cls.BASE_CONF_FILEPATH,'w','utf8') as f:
                jsonstr = json.dumps(cls.BASE_CONF_CONTENT,ensure_ascii=False)
                f.write(jsonstr)
        # 加载基础配置，赋值类变量BASE
        with codecs.open(cls.BASE_CONF_FILEPATH,'r','utf8') as f:
            cls.BASE = json.load(f)

    @classmethod
    def save_basic_cfg(cls):
        '''
        储存基础配置
        '''
        if not cls.BASE:return
        with codecs.open(cls.BASE_CONF_FILEPATH,'w','utf8') as f:
            jsonstr = json.dumps(cls.BASE,ensure_ascii=False)
            f.write(jsonstr)

    @classmethod
    def load_excel_cfg(cls):
        '''
        初始化及加载Excel配置
        '''
        cls.EXCEL = {}
        cls.EXCEL["BASE_FOLDER"] = XlManager.fetch_name('BASE_FOLDER')
        cls.EXCEL["AUTO_BACKUP"] = XlManager.fetch_name('AUTO_BACKUP')
        cls.EXCEL["NO_HIDDEN_FILES_WIN"] = XlManager.fetch_name('NO_HIDDEN_FILES_WIN')
        cls.EXCEL["NO_HIDDEN_FILES_POINT"] = XlManager.fetch_name('NO_HIDDEN_FILES_POINT')
        cls.EXCEL["NO_FOLDERS"] = XlManager.fetch_name('NO_FOLDERS')
        cls.EXCEL["rIGNORE_FOLDERS"] = XlManager.fetch_name('rIGNORE_FOLDERS')
        cls.EXCEL["rEXT_BLACKLIST"] = XlManager.fetch_name('rEXT_BLACKLIST')
        cls.EXCEL["rEXT_WHITELIST"] = XlManager.fetch_name('rEXT_WHITELIST')
        cls.EXCEL["rKEY_MODE"] = XlManager.fetch_name('rKEY_MODE')

if __name__ == '__main__':
    ConfManager.load_basic_cfg()
    print(ConfManager.BASE)
    ConfManager.BASE['CONF_SHEET_NAME'] = 'conf'
    ConfManager.save_basic_cfg()
