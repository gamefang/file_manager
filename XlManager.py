# -*- coding: utf-8 -*-

import pandas as pd
from openpyxl import load_workbook

class XlManager():
    '''
    实现Excel读写的功能
    TODO 隐藏文件过滤
    '''

    @staticmethod
    def write_to_excel(data,fp,sheet_name):
        '''
        写入数据至Excel文件中
        '''
        workbook = load_workbook(filename = fp)
        # 同名sheet备份
        if sheet_name in workbook.sheetnames:
            bak_sheet_name = sheet_name + '_bak'
            if bak_sheet_name in workbook.sheetnames:
                workbook.remove_sheet(workbook[bak_sheet_name])
            workbook[sheet_name].title = bak_sheet_name
        # 创建新sheet
        sheet = workbook.create_sheet(sheet_name,0)
        # 输出表头
        table_head = [
            '文件名','扩展名','是文件夹','文件路径','文件大小','创建时间','修改时间','访问时间'
            ]
        sheet.append(table_head)
        # 输出数据内容
        for k,v in data.items():
            line_data = [
                v['filename'],v['ext'],v['is_folder'],v['path'],v['size'],v['ctime'],v['mtime'],v['atime']
                ]
            sheet.append(line_data)
        workbook.save(filename = fp)

if __name__ == '__main__':
    pass
