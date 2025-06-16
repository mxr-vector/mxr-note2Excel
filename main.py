import os
from handler import JsonHandler,ExcelHandler

'''
统一项目路径，以调用该方法的脚本位置为基础
'''
os.chdir(os.path.dirname(__file__))

if __name__ == "__main__":
    json_handler = JsonHandler()
    excel_handler = ExcelHandler()
    '加载json数据'
    data = json_handler.get_json_data("assets/data.json")
    '写入excel文件'
    # excel_handler.write_excel("assets/data.xlsx",data)
    
    excel_handler.write_excel("assets/data.xlsx",data,['姓名','手机号'],True)
