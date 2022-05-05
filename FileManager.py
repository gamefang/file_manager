# -*- coding: utf-8 -*-

import os

class FileManager():
    '''
    实际文件信息获取、递归文件目录获取等功能
    '''

    @staticmethod
    def get_file_data(folder=None):
        '''
        获取文件全部数据
        @param folder: 待获取文件数据的目录
        @return: 文件数据字典
        '''
        file_data = {}
        folder = folder or os.curdir
        for root,dirs,files in os.walk(folder):
            for fn in files:
                # 文件名、扩展名
                filename,ext = os.path.splitext(fn)
                # 绝对、相对路径
                full_path = os.path.join(root,fn)
                path = os.path.relpath(full_path,folder)
                # 是文件夹、文件大小、各种时间
                is_folder = os.path.isdir(full_path)
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
                    'is_folder':is_folder,
                    'path':path,
                    'size':size,
                    'ctime':ctime,
                    'mtime':mtime,
                    'atime':atime,
                }
        return file_data

if __name__ == '__main__':
    pass
