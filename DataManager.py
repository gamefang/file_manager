# -*- coding: utf-8 -*-

import json
import codecs

class DataManager():
    '''
    数据融合、比较、冲突记录
    '''
    BACKUP_FILE = '.backup.json'

    @staticmethod
    def get_merged_data(excel_data,file_data,CFG):
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
        # 查漏excel_data（不自动删除的情况下）
        if not CFG.EXCEL['AUTO_DEL_IGNORED']:
            final_keys = final_data.keys()
            for k,v in excel_data.items():
                if k not in final_keys: # excel有file无
                    v['status'] = 'del'
                    final_data[k] = v
        return final_data

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

    @classmethod
    def backup_data(cls,data):
        '''
        备份数据至文件
        '''
        with codecs.open(cls.BACKUP_FILE,'w','utf8') as f:
            jsonstr = json.dumps(data,ensure_ascii=False)
            f.write(jsonstr)
