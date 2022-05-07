# -*- coding: utf-8 -*-

__version__ = 1.0
__author__ = 'gamefang'

from ConfManager import ConfManager as CFG
from XlManager import XlManager as XlManager
from FileManager import FileManager as FileManager

def main():
    # 加载配置
    CFG.load_basic_cfg()
    # 加载Excel数据
    CFG.load_excel_cfg()
    # 解析递归文件信息
    file_data = FileManager.get_file_data()
    # 数据融合及冲突检查
    # 写入Excel数据
    XlManager.write_to_excel(file_data,
        CFG.BASE['EXCEL_FILE_PATH'],
        CFG.BASE['LIST_SHEET_NAME'],
        )

if __name__ == '__main__':
    import time
    st = time.time()
    main()
    ed = time.time()
    print(ed-st)
