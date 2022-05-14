# -*- coding: utf-8 -*-

__version__ = 1.0
__author__ = 'gamefang'

import time

import easygui as gui

from ConfManager import ConfManager as CFG
from XlManager import XlManager as XlManager
from FileManager import FileManager as FileManager
from DataManager import DataManager as DataManager

def main():
    st = time.time()
    # 加载基础配置
    CFG.load_basic_cfg()
    EXCEL_FILE_PATH = CFG.BASE['EXCEL_FILE_PATH']
    # 加载Excel配置
    if XlManager.is_excel_opened(EXCEL_FILE_PATH):
        gui.msgbox(f'请先关闭文件：{EXCEL_FILE_PATH}')
        return
    try:
        XlManager.load_cur_file(EXCEL_FILE_PATH)
    except Exception as e:
        gui.msgbox('加载Excel文件失败！\n',e)
        return
    try:
        CFG.load_excel_cfg()
    except Exception as e:
        gui.msgbox('加载Excel配置失败！\n',e)
        return
    # 加载Excel数据
    try:
        excel_data = XlManager.load_excel_data(CFG)
    except Exception as e:
        gui.msgbox('加载Excel数据失败！\n',e)
        return
    # 解析递归文件信息
    try:
        file_data = FileManager.get_file_data(CFG)
    except Exception as e:
        gui.msgbox('解析文件信息失败！\n',e)
        return
    # 数据融合及冲突检查
    try:
        final_data = DataManager.get_merged_data(excel_data,file_data,CFG)
    except Exception as e:
        gui.msgbox('数据融合及冲突检查错误！\n',e)
        return
    # 写入最终版本Excel数据
    try:
        XlManager.write_data_to_excel(final_data,CFG)
    except Exception as e:
        gui.msgbox('写入最终版本Excel失败！\n',e)
        return
    ed = time.time()
    gui.msgbox(f"已成功同步所有文件数据！用时：{ed-st}秒\n文件清单：{CFG.BASE['EXCEL_FILE_PATH']}\n如信息丢失，请查看备份数据：{DataManager.BACKUP_FILE}\n说明文档：https://github.com/gamefang/file_manager",'FileManger by gamefang')

if __name__ == '__main__':
    main()
