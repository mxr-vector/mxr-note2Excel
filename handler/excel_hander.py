from turtle import title
from numpy import void
import pandas as pd
'''
excel处理器
'''
class ExcelHandler:
    def __init__(self):
        self.field=''
        self.context=[]
        pass

    def json2entity(self,json: dict) -> void:
        if 'field' in json and 'context' in json:
            self.field = json['field']
            self.context = json['context']
        else:
            raise ValueError("JSON data missing required fields")
    '''
    写入excel文件
    @param file_name: 文件名称
    @param data: json数据
    @param filterd: 筛选字段
    @param Pagination: 是否分页
    '''
    def write_excel(self,file_name,data,filterd=[],Pagination=False) -> void:
        self.json2entity(data)
        field = self.field
        data = self.context
        df_all = pd.DataFrame(data,columns=field)
        if(not str.endswith(file_name,".xlsx")):
            file_name = file_name + ".xlsx"
        
        '是否开启分页模式，默认精简模式'
        if Pagination and (not filterd):
            '分页模式'
            df_filterd = df_all[df_all[filterd]]
            # 写入多个表单，使用 ExcelWriter
            with pd.ExcelWriter(file_name) as writer:
                df_filterd.to_excel(writer, sheet_name='Sheet1', index=False)
                df_all.to_excel(writer, sheet_name='Sheet2', index=False)
        else:
            '精简模式'
            # 将 DataFrame 写入 Excel 文件，写入 'Sheet1' 表单
            df_all.to_excel(file_name, sheet_name='Sheet1', index=False)