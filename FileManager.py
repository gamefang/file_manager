# -*- coding: utf-8 -*-

import os

class FileManager():
    '''
    实际文件信息获取、递归文件目录获取等功能
    '''

    @classmethod
    def get_file_data(cls,folder=None,no_hidden_files=True):
        '''
        获取文件全部数据
        @param folder: 待获取文件数据的目录
        @return: 文件数据字典
        '''
        file_data = {}
        folder = folder or os.curdir
        for root,dirs,files in os.walk(folder):
            for dir in dirs:
                # 绝对、相对路径
                full_path = os.path.join(root,dir)
                if no_hidden_files and cls.is_path_hidden(full_path):
                    continue    # 去除隐藏文件或递归被隐藏文件
                path = os.path.relpath(full_path,folder)
                # 各种时间
                atime = os.path.getatime(full_path)
                mtime = os.path.getmtime(full_path)
                ctime = os.path.getctime(full_path)
                # key，等于mtime
                # 数据记录
                file_data[mtime] = {
                    'filename':dir,
                    'ext':'',
                    'is_folder':True,
                    'path':path,
                    'size':0,
                    'ctime':ctime,
                    'mtime':mtime,
                    'atime':atime,
                }
            for fn in files:
                # 绝对、相对路径
                full_path = os.path.join(root,fn)
                if no_hidden_files and cls.is_path_hidden(full_path):
                    continue    # 去除隐藏文件或递归被隐藏文件
                path = os.path.relpath(full_path,folder)
                # 文件名、扩展名
                filename,ext = os.path.splitext(fn)
                ext = ext[1:]   # 去掉.
                # 文件大小、各种时间
                size = os.path.getsize(full_path)
                atime = os.path.getatime(full_path)
                mtime = os.path.getmtime(full_path)
                ctime = os.path.getctime(full_path)
                # key，文件大小与修改时间拼合
                key = mtime + size * 10000000000
                # 数据记录
                file_data[key] = {
                    'filename':filename,
                    'ext':ext,
                    'is_folder':False,
                    'path':path,
                    'size':size,
                    'ctime':ctime,
                    'mtime':mtime,
                    'atime':atime,
                }
        return file_data

    @staticmethod
    def is_path_hidden(path):
        '''
        根据路径判断是否为隐藏文件或递归式被隐藏文件
        暂以路径中是否含.开头的文件夹或文件判断
        '''
        path = os.path.normcase(path)
        li = path.split(os.sep)
        for item in li[1:]: # 跳过第一个.路径
            if item.startswith('.'):
                return True
        return False

    # win隐藏文件判断
    def is_hidden_file(self, file_path):
        """ 判断 windows 系统下，文件是否为隐藏文件,是则返回 True """
        import win32file  # 安装好 pywin32 后即可 ，操作windows的
        import win32con  # 安装好 pywin32 后即可 ，操作windows的
        import platform  # 判断电脑系统是什么系统
        if 'Windows' in platform.system():
            file_attr = win32file.GetFileAttributes(file_path)
            if file_attr & win32con.FILE_ATTRIBUTE_HIDDEN:
                return True
            return False
        return False

if __name__ == '__main__':
    pass
