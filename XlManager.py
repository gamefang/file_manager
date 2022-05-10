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

    @classmethod
    def write_data_to_excel(cls,data,CFG):
        '''
        写入最终数据至Excel文件中
        @param data: 待写入data（来自FileManager的处理结果）
        @CFG: 配置信息，来自ConfManager
        '''
        workbook = cls.CUR_WB
        sheet_name = CFG.BASE['LIST_SHEET_NAME']
        if sheet_name not in workbook.sheetnames:
            raise Exception('Excel file wrong!')
        # sheet备份（自动加顺序号）
        if CFG.EXCEL['AUTO_BACKUP']:
            sheet_copy = workbook.copy_worksheet(workbook[sheet_name])
        # 使用原sheet
        ws = workbook[sheet_name]
        # 表头提取
        head_start_cell = cls.get_cell_by_value(ws,'key')
        head_row = ws[head_start_cell.row]
        # 根据表头确定输出字段顺序
        output_params = [cell.value for cell in head_row]
        # 清空无用表格行（保留表头下方文字表头行）
        ws.delete_rows(head_start_cell.row + 2, ws.max_row)
        # 按顺序输出
        p_row = head_start_cell.row + 2 # row指针
        for _,v in data.items():
            p_col = head_start_cell.column # col指针
            for item in output_params:
                cur_cell = ws.cell(
                    column = p_col,
                    row = p_row,
                    value = v.get(item,''), # 留空不存在数据
                )
                p_col += 1
            p_row += 1
        # 存储
        workbook.save(CFG.BASE['EXCEL_FILE_PATH'])

    @staticmethod
    def get_cell_by_value(sheet,value):
        '''
        按值获取单元格对象（先横后纵，取第一个值）
        @param sheet: 值所在的Excel工作表
        @param value: 所需的值
        @return: cell对象
        '''
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == value:
                    return cell

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

if __name__ == '__main__':
    XlManager.load_cur_file('FileManager.xlsx')
    res = XlManager.fetch_name('NO_HIDDEN_FILES_WIN')
    print(res)
    res = XlManager.fetch_name('rEXT_WHITELIST')
    print(res)
