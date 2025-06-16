# mxr-note2Excel
一个按照规定格式写记事本，即可转为excel的工具

## 功能

- **记事本转Excel**：将特定格式的记事本内容转换为结构化的Excel文件。
- **字段自适应长度**：生成的Excel文件列宽将根据内容自动调整，提高可读性。
- **分页模式**：支持将数据分页写入Excel，方便管理大量数据。

## 安装

1. **克隆仓库**：

2. **安装依赖**：
```bash
uv sync
```
## 使用
3. **运行程序**：
```bash
uv run main.py
```

   程序将提示您输入记事本文件路径和输出Excel文件名称。

4. **查看生成的Excel文件**：

   生成的Excel文件将保存在指定路径，并具有自动调整的列宽和可选的分页。

3.使用 pyinstaller 代码带包为exe可执行文件的命令：
```bash
pyinstaller --name note2Excel_win64 --windowed --icon=assets/favicon.ico --add-data="assets/favicon.ico;./assets" --add-data="assets/data.json;./data" main.py
```

## 示例
请参考 `assets/data.json` 文件了解数据结构，以及 `handler/excel_hander.py` 中的 `write_excel` 方法了解Excel写入逻辑。
