# -*- coding: utf-8 -*-

__version__ = 1.0
__author__ = 'gamefang'

from ConfManager import ConfManager as CFG
from XlManager import XlManager as XlManager
from FileManager import FileManager as FileManager

def main():
    # 加载基础配置
    CFG.load_basic_cfg()
    EXCEL_FILE_PATH = CFG.BASE['EXCEL_FILE_PATH']
    # 加载Excel配置
    if XlManager.is_excel_opened(EXCEL_FILE_PATH):
        raise Exception(f'请先关闭文件：{EXCEL_FILE_PATH}')
    XlManager.load_cur_file(EXCEL_FILE_PATH)
    CFG.load_excel_cfg()
    # 加载Excel数据
    excel_data = XlManager.load_excel_data(CFG)
    # TODO 解析递归文件信息
    real_data = FileManager.get_file_data()
    # TODO 数据融合及冲突检查
    file_data = real_data
    # 写入最终版本Excel数据
    XlManager.write_data_to_excel(file_data,CFG)

if __name__ == '__main__':
    import time
    st = time.time()
    main()
    ed = time.time()
    print(ed-st)
