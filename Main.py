# -*- coding: utf-8 -*-

__version__ = 1.0
__author__ = 'gamefang'

from FileManager import FileManager as FileManager
from XlManager import XlManager as XlManager

excel_data = {}
file_data = {}

EXCEL_FILE = 'FileManager.xlsx'

def main():
    # 加载配置
    # 加载Excel数据
    # 解析递归文件信息
    global file_data
    file_data = FileManager.get_file_data()
    # 数据融合及冲突检查
    # 写入Excel数据
    XlManager.write_to_excel(file_data,EXCEL_FILE,'Sheet1')

if __name__ == '__main__':
    main()
