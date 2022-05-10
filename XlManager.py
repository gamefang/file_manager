# -*- coding: utf-8 -*-

import os
from openpyxl import load_workbook

class XlManager():
    '''
    实现Excel读写的功能
    '''
    # 当前加载的Excel文件
    CUR_WB = None

    @classmethod
    def load_cur_file(cls,fp):
        '''
        加载当前使用的Excel文件，供随时访问
        @param fp: Excel文件路径
        '''
        fp = os.path.normcase(fp)
        if os.path.exists(fp):
            cls.CUR_WB = load_workbook(filename = fp)

    @classmethod
    def fetch_name(cls,defined_name,is_return_cell=False):
        '''
        从当前的Excel文件中，加载名称对应的数据
        @param defined_name: Excel中定义的名称
        @param is_return_cell: 是否返回cell对象
        @return: cell对象或列表；泛型数值或列表
        '''
        if not cls.CUR_WB:return
        dn = cls.CUR_WB.defined_names[defined_name]
        cells = []
        for k,v in dn.destinations:
            ws = cls.CUR_WB[k]
            cells.append(ws[v])
        cells = cells[0]    # 去掉无用的列表层
        # 返回cell对象
        if is_return_cell:
            return cells
        # 返回值
        if type(cells) is tuple: # 区域
            return [ cell.value
                for cell in cells
                if cell.value
                ][1:]
        else:   # 单独单元格
            return cells.value

    @staticmethod
    def write_to_excel(data,fp,sheet_name):
        '''
        TODO: 待优化，精确至cell
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

    @staticmethod
    def is_excel_opened(fp):
        '''
        判断Excel文件是否已打开（通过是否生成了~$文件判定）
        @param fp: Excel文件路径
        @return: bool
        '''
        fp = os.path.normcase(fp)
        dir_name,file_name = os.path.split(fp)
        hidden_fp = os.path.join(dir_name,'~$' + file_name)
        return os.path.exists(hidden_fp)

    # @staticmethod
    # def get_cell(fp,defined_name=None,sheet_name=None,cell=None):
    #     '''
    #     获取某Excel文件某Sheet的单独单元格数据
    #     @param fp: Excel文件
    #     @param defined_name: Excel中定义的名称。如无，则使用sheet+cell的定位方法
    #     @param sheet_name: Excel文件中的sheet名
    #     @param cell: Excel文件中的cell名
    #     @return: 泛型的Excel数据
    #     '''
    #     return result

if __name__ == '__main__':
    XlManager.load_cur_file('FileManager.xlsx')
    res = XlManager.fetch_name('NO_HIDDEN_FILES_WIN')
    print(res)
    res = XlManager.fetch_name('rEXT_WHITELIST')
    print(res)
