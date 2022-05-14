# -*- coding: utf-8 -*-

import os
import platform
# pip install pywin32
if 'Windows' in platform.system():
    import win32file
    import win32con

from DataManager import DataManager as DataManager

class FileManager():
    '''
    实际文件信息获取、递归文件目录获取等功能
    '''

    @classmethod
    def get_file_data(cls,CFG):
        '''
        获取文件全部数据
        @param CFG: 配置信息，来自ConfManager
        @return: 文件数据字典
        '''
        file_data = {}
        folder = CFG.EXCEL['BASE_FOLDER'] or os.curdir
        for root,dirs,files in os.walk(folder):
            if not CFG.EXCEL['NO_FOLDERS']: # 需处理文件夹
                for dir in dirs:
                    cur_dic = {
                        'filename':dir,
                        'ext':'',
                        'is_folder':True,
                        'size':0,
                        }
                    # 绝对、相对路径
                    full_path = os.path.join(root,dir)
                    if cls.need_pass(full_path,CFG):   # 跳过隐藏文件
                        continue
                    cur_dic['path'] = os.path.relpath(full_path,folder)
                    # 各种时间（整型化）
                    cur_dic['atime'] = int(os.path.getatime(full_path))
                    cur_dic['mtime'] = int(os.path.getmtime(full_path))
                    cur_dic['ctime'] = int(os.path.getctime(full_path))
                    # 确定key
                    list_keys = file_data.keys()
                    key = DataManager.get_key(cur_dic,CFG,list_keys)
                    # 数据记录
                    file_data[key] = cur_dic
            for fn in files:
                cur_dic = {'is_folder':False}
                # 绝对、相对路径
                full_path = os.path.join(root,fn)
                if cls.need_pass(full_path,CFG):   # 跳过隐藏文件
                    continue
                cur_dic['path'] = os.path.relpath(full_path,folder)
                # 文件名、扩展名
                cur_dic['filename'],ext = os.path.splitext(fn)
                cur_dic['ext'] = ext[1:]   # 去掉.
                if cls.need_pass_ext(cur_dic['ext'],CFG):   # 按扩展名跳过检测
                    continue
                # 文件大小、各种时间（整型化）
                cur_dic['size'] = os.path.getsize(full_path)
                cur_dic['atime'] = int(os.path.getatime(full_path))
                cur_dic['ctime'] = int(os.path.getctime(full_path))
                cur_dic['mtime'] = int(os.path.getmtime(full_path))
                # 确定key
                list_keys = file_data.keys()
                key = DataManager.get_key(cur_dic,CFG,list_keys)
                # 数据记录
                file_data[key] = cur_dic
        return file_data

    @classmethod
    def need_pass(cls,full_path,CFG):
        '''
        判断文件是否需跳过
        @param full_path: 文件完整路径
        @param CFG: 配置信息，来自ConfManager
        @return: bool是否需跳过
        '''
        full_path = os.path.normcase(full_path)
        # 忽略的关键字
        for item in CFG.EXCEL['rIGNORE_KEYWORDS']:
            if item in full_path:
                return True
        # 排除点开头的
        if CFG.EXCEL['NO_HIDDEN_FILES_POINT'] and cls.is_point_path(full_path):
            return True
        # 排除win隐藏文件
        if CFG.EXCEL['NO_HIDDEN_FILES_WIN'] and cls.is_win_hidden_file(full_path):
            return True

    @staticmethod
    def is_point_path(path):
        '''
        判断是否为以.开头的文件或递归文件夹中文件，视为隐藏文件
        '''
        li = path.split(os.sep)
        for item in li[1:]: # 跳过第一个.路径
            if item.startswith('.'):
                return True
        return False

    @staticmethod
    def is_win_hidden_file(fp):
        '''
        判断是否为win系统下的隐藏文件
        '''
        if 'Windows' in platform.system():
            file_attr = win32file.GetFileAttributes(fp)
            return file_attr & win32con.FILE_ATTRIBUTE_HIDDEN
        return False

    @staticmethod
    def need_pass_ext(ext,CFG):
        '''
        按扩展名判断文件是否需跳过
        @param ext: 文件扩展名，不含.
        @param CFG: 配置信息，来自ConfManager
        @return: bool是否需跳过
        '''
        # 扩展名白名单
        if CFG.EXCEL['rEXT_WHITELIST'] and ext not in CFG.EXCEL['rEXT_WHITELIST']:
            return True
        # 扩展名黑名单
        if ext in CFG.EXCEL['rEXT_BLACKLIST']:
            return True

if __name__ == '__main__':
    pass
