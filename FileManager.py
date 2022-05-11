# -*- coding: utf-8 -*-

import os

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
            for dir in dirs:
                cur_dic = {
                    'filename':dir,
                    'ext':'',
                    'is_folder':True,
                    'size':0,
                    }
                # 绝对、相对路径
                full_path = os.path.join(root,dir)
                if CFG.EXCEL['NO_HIDDEN_FILES_POINT'] and cls.is_path_hidden(full_path):
                    continue    # 去除隐藏文件或递归被隐藏文件
                cur_dic['path'] = os.path.relpath(full_path,folder)
                # 各种时间（整型化）
                cur_dic['atime'] = int(os.path.getatime(full_path))
                cur_dic['mtime'] = int(os.path.getmtime(full_path))
                cur_dic['ctime'] = int(os.path.getctime(full_path))
                # 确定key
                list_keys = file_data.keys()
                key = cls.get_key(cur_dic,CFG,list_keys)
                # 数据记录
                file_data[key] = cur_dic
            for fn in files:
                cur_dic = {'is_folder':False}
                # 绝对、相对路径
                full_path = os.path.join(root,fn)
                if CFG.EXCEL['NO_HIDDEN_FILES_POINT'] and cls.is_path_hidden(full_path):
                    continue    # 去除隐藏文件或递归被隐藏文件
                cur_dic['path'] = os.path.relpath(full_path,folder)
                # 文件名、扩展名
                cur_dic['filename'],ext = os.path.splitext(fn)
                cur_dic['ext'] = ext[1:]   # 去掉.
                # 文件大小、各种时间（整型化）
                cur_dic['size'] = os.path.getsize(full_path)
                cur_dic['atime'] = int(os.path.getatime(full_path))
                cur_dic['ctime'] = int(os.path.getctime(full_path))
                cur_dic['mtime'] = int(os.path.getmtime(full_path))
                # 确定key
                list_keys = file_data.keys()
                key = cls.get_key(cur_dic,CFG,list_keys)
                # 数据记录
                file_data[key] = cur_dic
        return file_data

    @classmethod
    def get_merged_data(cls,excel_data,file_data,CFG):
        '''
        获取最终合并的数据
        @param excel_data: Excel数据
        @param file_data: 文件索引生成的数据
        @return: 合并后的数据
        '''
        final_data = {}
        excel_keys = excel_data.keys()
        # 遍历file_data
        for k,v in file_data.items():
            if k in excel_keys: # 有共同key
                excel_v = excel_data[k]
                is_changed = False
                # 合并同key字典
                excel_v_keys = excel_v.keys()
                for m,n in v.items():   # 遍历file字段
                    # excel中有同名字段，且值不相等，则被覆盖
                    if m in excel_v_keys and n != excel_v[m]:
                        is_changed = True
                        break
                excel_v.update(v)   # 以file为准，融入excel字典
                v = excel_v
                v['status'] = ('','mod')[is_changed]
            else:   # file有excel无，新增
                v['status'] = 'new'
            final_data[k] = v
        # 查漏excel_data
        final_keys = final_data.keys()
        for k,v in excel_data.items():
            if k not in final_keys: # excel有file无
                v['status'] = 'del'
                final_data[k] = v
        return final_data

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

    @staticmethod
    def get_key(cur_dic,CFG,list_keys,cur_key=None):
        '''
        根据用户配置及当前数据，生成唯一的key
        @param cur_dic: 当前数据的字典
        @param CFG: 配置信息，来自ConfManager
        @param list_keys: 现有大字典中，所有key的列表，用于查重
        @param cur_key: 现指定使用的key，需要查重，可能非最终状态
        @return: 确定的唯一key
        '''
        if cur_key:  # 使用已有
            key = cur_key
        else:   # 按规则新编
            key = ''
            key_list = [
                str(cur_dic[item])
                for item in CFG.EXCEL['rKEY_MODE']
            ]
            key = '|'.join(key_list)
        if key in list_keys: # 有重号key
            subfix = 1
            while 1:    # 持续顺序编号直至不重号
                try_new_key = f'{key}+{subfix}'
                if try_new_key in list_keys:
                    subfix += 1
                else:
                    key = try_new_key
                    break
        return key

if __name__ == '__main__':
    pass
