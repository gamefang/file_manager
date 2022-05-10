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
        cls.EXCEL["NO_HIDDEN_FILES_WIN"] = XlManager.fetch_name('NO_HIDDEN_FILES_WIN')
        cls.EXCEL["NO_HIDDEN_FILES_POINT"] = XlManager.fetch_name('NO_HIDDEN_FILES_POINT')
        cls.EXCEL["NO_FOLDERS"] = XlManager.fetch_name('NO_FOLDERS')
        cls.EXCEL["rIGNORE_FOLDERS"] = XlManager.fetch_name('rIGNORE_FOLDERS')
        cls.EXCEL["rEXT_BLACKLIST"] = XlManager.fetch_name('rEXT_BLACKLIST')
        cls.EXCEL["rEXT_WHITELIST"] = XlManager.fetch_name('rEXT_WHITELIST')

# import configparser
#
# class ConfManager():
#     '''
#     配置项(ini)的加载、读取、存储功能
#     '''
#     # 待引用的配置 Conf.INI.XXX
#     INI = None
#
#     # 配置项解析
#     @staticmethod
#     def s2b(str):   # 配置项解析：字符串转bool
#         return True if str.lower() in ('true','1') else False
#     @staticmethod
#     def s2ls(str):  # 配置项解析：字符串分割为字符串list
#         return str.split(',')
#     @staticmethod
#     def s2ts(str):  # 配置项解析：字符串分割为字符串tuple
#         return tuple(str.split(','))
#     # 类型前缀处理方法的字典，默认为字符串
#     DIC_TYPE={
#             '<b>':s2b,
#             '<ts>':s2ts,
#             '<i>':int,
#             '<f>':float,
#     }
#
#     @classmethod
#     def odic_clean(cls,odic):
#         '''
#         清洗ini文件的配置字典
#         @param odic: ini文件解析的default或sections的有序字典数据
#         @return: 常规字典
#         '''
#         result = {}
#         for k,v in odic.items():
#             if not k.startswith( tuple(cls.DIC_TYPE.keys()) ):
#                 myk,myv = k,v
#             else:
#                 for pre in cls.DIC_TYPE.keys():
#                     if k.startswith(pre):
#                         myk = k[len(pre):]
#                         if v == 'None':   # 所有非字符串的None均为None
#                             myv = None
#                         else:
#                             myv = cls.DIC_TYPE[pre](v)
#             result[myk] = myv
#         return result
#
#     @classmethod
#     def load_ini_cfg(cls,cfg_file_path,get_obj=True):
#         '''
#         加载ini配置
#         @param cfg_file_path: ini配置文件路径
#         @param get_obj: 是否获取为Dobj对象
#         @return: 整理后的配置字典或Dobj对象
#         '''
#         cfg = configparser.ConfigParser()
#         cfg.read(cfg_file_path,encoding='utf8')
#         result = cls.odic_clean(cfg._defaults)   # DEFAULT节点解析
#         for sec,odic in cfg._sections.items():   # 其它节点解析
#             result[sec] = cls.odic_clean(odic)
#         if get_obj:
#             o = Dobj(result)
#             assert o.is_have,f'配置文件 {cfg_file_path} 错误或无配置文件！'
#             return o
#         INI = result
#         return result
#
# class Dobj(object):
#     '''
#     将嵌套的字典转为嵌套的对象
#     用法：
#         myDic={'a':1,'b':{'c':2,'d':{'e':3,'f':4}}}
#         myDicObj=Dobj(myDic)
#         print(myDicObj) #  {'a':1,'b':{'c':2,'d':{'e':3,'f':4}}}
#         print(myDicObj.b.d.e)   #  3
#     '''
#     def __init__(self, dic):
#         for k, v in dic.items():
#             if isinstance(v, list):
#                 setattr(self, k, [self.__cls__(x) if isinstance(x, dict) else x for x in v])
#             elif isinstance(v, tuple):
#                 setattr( self, k, tuple([self.__cls__(x) if isinstance(x, dict) else x for x in v]) )
#             else:
#                 setattr(self, k, Dobj(v) if isinstance(v, dict) else v)
#     def __repr__(self):
#         return repr(self.__dict__)
#     @property
#     def is_have(self):
#         return bool(self.__dict__)

if __name__ == '__main__':
    ConfManager.load_basic_cfg()
    print(ConfManager.BASE)
    ConfManager.BASE['CONF_SHEET_NAME'] = 'conf'
    ConfManager.save_basic_cfg()
