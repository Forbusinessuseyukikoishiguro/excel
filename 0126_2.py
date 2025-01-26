import pandas as pd
from openpyxl import load_workbook
import os

class ExcelOperator:
    def __init__(self, file_path):
        self.file_path = file_path
    
    def write_multiple_cells_and_save(self, cell_values):
        """
        複数のセルに値を書き込み、自動的に保存を行います
        
        Parameters:
        cell_values (dict): セル位置と値の辞書
        """
        try:
            workbook = load_workbook(self.file_path)
            sheet = workbook.active
            
            for cell_position, value in cell_values.items():
                sheet[cell_position] = value
                print(f"'{value}' をセル {cell_position} に書き込みました")
            
            workbook.save(self.file_path)
            print(f"\n{os.path.basename(self.file_path)}にすべての内容を保存しました")
            
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")

# 使用例
if __name__ == "__main__":
    input_file2 = r"C:\Users\yukik\Desktop\excel\ex2.xlsx"
    
    # 書き込む内容を辞書形式で定義
    cell_values = {
        'A1': 'イチゴ大福',
        'A2': 'ブドウ大福',
        'A3': '抹茶大福',
        'A4': '最中抹茶'
    }
    
    excel_op = ExcelOperator(input_file2)
    excel_op.write_multipl
#2025/1/26 ex2にA1にイチゴ大福、A2にブドウ大福、A3に抹茶大福、A4に最中抹茶を書き込みましたOK