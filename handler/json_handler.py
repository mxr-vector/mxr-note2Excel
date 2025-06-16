
import json

'''
json处理器
'''
class JsonHandler:
    def __init__(self):
        pass

    def load_data(self,file_path: str) -> dict:
        with open(file_path, 'r', encoding='utf-8') as f:
                j = json.load(f)
                fields = j.get('field', [])
                data = j.get('context', [])
                return fields, data

    def save_data(self,file_path,fields,data):
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump({'field': fields, 'context': data}, f, ensure_ascii=False, indent=4)

