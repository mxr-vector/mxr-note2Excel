import os
from handler import JsonHandler,ExcelHandler
from web import create_UI
'''
统一项目路径，以调用该方法的脚本位置为基础
'''
os.chdir(os.path.dirname(__file__))

if __name__ == "__main__":
    create_UI()
