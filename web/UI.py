import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLineEdit, QLabel, QMessageBox, QHeaderView, QDialog, QFormLayout, QFileDialog
)
from PyQt6.QtGui import QFont
import os

from handler import JsonHandler
from handler import ExcelHandler

DATA_PATH = os.path.join(os.path.dirname(__file__), '../assets/data.json')
JSON_DATA = JsonHandler()
class DataDialog(QDialog):
    def __init__(self, fields, data=None, parent=None):
        super().__init__(parent)
        self.fields = fields
        self.data = data or ['' for _ in fields]
        self.setWindowTitle('数据编辑')
        self.resize(400, 100)
        layout = QFormLayout()
        self.edits = []
        for i, field in enumerate(fields):
            edit = QLineEdit(self.data[i] if i < len(self.data) else '')
            layout.addRow(field, edit)
            self.edits.append(edit)
        self.setLayout(layout)
        self.result = None
        btns = QHBoxLayout()
        ok_btn = QPushButton('确定')
        cancel_btn = QPushButton('取消')
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btns.addWidget(ok_btn)
        btns.addWidget(cancel_btn)
        layout.addRow(btns)
    def getData(self):
        return [edit.text() for edit in self.edits]

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('现代化数据管理客户端')
        self.resize(900, 600)
        self.setStyleSheet('QMainWindow { background: #f5f6fa; }')
        font = QFont('微软雅黑', 10)
        self.setFont(font)
        self.fields, self.data = self.load_data()
        self.initUI()
    def load_data(self):
         if os.path.exists(DATA_PATH):
            return JSON_DATA.load_data(DATA_PATH)
    def save_data(self):
        JSON_DATA.save_data(DATA_PATH,self.fields,self.data)
    def initUI(self):
        central = QWidget()
        vbox = QVBoxLayout()
        # 字段编辑区
        field_box = QHBoxLayout()
        self.field_edit = QLineEdit(','.join(self.fields))
        self.field_edit.setPlaceholderText('用逗号分隔字段名，如：编号,姓名,地址...')
        update_fields_btn = QPushButton('更新字段')
        update_fields_btn.clicked.connect(self.update_fields)
        field_box.addWidget(QLabel('字段:'))
        field_box.addWidget(self.field_edit)
        field_box.addWidget(update_fields_btn)
        vbox.addLayout(field_box)
        # 表格区
        self.table = QTableWidget()
        self.table.setColumnCount(len(self.fields))
        self.table.setHorizontalHeaderLabels(self.fields)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.refresh_table()
        vbox.addWidget(self.table)
        # 操作按钮区
        btn_box = QHBoxLayout()
        add_btn = QPushButton('新增')
        edit_btn = QPushButton('编辑')
        del_btn = QPushButton('删除')
        export_simple_btn = QPushButton('精简导出')
        export_paged_btn = QPushButton('分页导出')
        add_btn.clicked.connect(self.add_row)
        edit_btn.clicked.connect(self.edit_row)
        del_btn.clicked.connect(self.del_row)
        export_simple_btn.clicked.connect(lambda: self.export_excel("simple"))
        export_paged_btn.clicked.connect(lambda: self.export_excel("paged"))
        btn_box.addWidget(add_btn)
        btn_box.addWidget(edit_btn)
        btn_box.addWidget(del_btn)
        btn_box.addWidget(export_simple_btn)
        btn_box.addWidget(export_paged_btn)
        vbox.addLayout(btn_box)
        central.setLayout(vbox)
        self.setCentralWidget(central)
    def refresh_table(self):
        self.table.setRowCount(len(self.data))
        self.table.setColumnCount(len(self.fields))
        self.table.setHorizontalHeaderLabels(self.fields)
        for row, row_data in enumerate(self.data):
            for col, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                self.table.setItem(row, col, item)
    def update_fields(self):
        fields = [f.strip() for f in self.field_edit.text().split(',') if f.strip()]
        if not fields:
            QMessageBox.warning(self, '警告', '字段不能为空')
            return
        self.fields = fields
        # 仅刷新表格，不操作json文件
        self.refresh_table()

    def export_excel(self, mode="simple"):
        file_name, _ = QFileDialog.getSaveFileName(self, f"保存{'精简' if mode=='simple' else '分页'}Excel文件", "data.xlsx", "Excel Files (*.xlsx)")
        if not file_name:
            return
        try:
            excel_handler = ExcelHandler()
            data = {'field': self.fields, 'context': self.data}
            if mode == "simple":
                # 精简导出，所有字段和数据导出到Sheet1
                excel_handler.write_excel(file_name, data)
                QMessageBox.information(self, '导出成功', f'精简数据已成功导出到 {file_name}')
            else:
                # 分页导出：sheet1用当前字段（部分字段），sheet2用全量字段和数据
                excel_handler.write_excel(file_name, data, self.fields, True)
                QMessageBox.information(self, '导出成功', f'分页数据已成功导出到 {file_name}')
        except Exception as e:
            QMessageBox.critical(self, '导出失败', f'导出Excel时发生错误: {e}')

    def export_simple_excel(self):
        pass
    def export_paged_excel(self):
        pass

    def add_row(self):
        dialog = DataDialog(self.fields, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_row = dialog.getData()
            if len(new_row) != len(self.fields):
                QMessageBox.warning(self, '警告', '数据长度与字段数不符')
                return
            self.data.append(new_row)
            self.refresh_table()
            self.save_data()

    def edit_row(self):
        selected = self.table.currentRow()
        if selected < 0 or selected >= len(self.data):
            QMessageBox.warning(self, '警告', '请先选择要编辑的行')
            return
        dialog = DataDialog(self.fields, self.data[selected], parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            updated_row = dialog.getData()
            if len(updated_row) != len(self.fields):
                QMessageBox.warning(self, '警告', '数据长度与字段数不符')
                return
            self.data[selected] = updated_row
            self.refresh_table()
            self.save_data()

    def del_row(self):
        selected = self.table.currentRow()
        if selected < 0 or selected >= len(self.data):
            QMessageBox.warning(self, '警告', '请先选择要删除的行')
            return
        reply = QMessageBox.question(self, '确认删除', '确定要删除选中行吗？', QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.data.pop(selected)
            self.refresh_table()
            self.save_data()

    def set_dark_mode(self, enable=True):
        if enable:
            dark_stylesheet = """
            QMainWindow { background: #232629; color: #f5f6fa; }
            QWidget { background: #232629; color: #f5f6fa; }
            QTableWidget { background: #232629; color: #f5f6fa; gridline-color: #444; }
            QHeaderView::section { background: #2c2f34; color: #f5f6fa; }
            QPushButton { background: #353941; color: #f5f6fa; border: 1px solid #444; }
            QLineEdit { background: #2c2f34; color: #f5f6fa; border: 1px solid #444; }
            QLabel { color: #f5f6fa; }
            QMessageBox { background: #232629; color: #f5f6fa; }
            """
            self.setStyleSheet(dark_stylesheet)
        else:
            self.setStyleSheet('QMainWindow { background: #f5f6fa; }')

def create_UI():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())