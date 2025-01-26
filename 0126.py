import pandas as pd
from openpyxl import load_workbook
import os

class ExcelOperator:
    def __init__(self, input_file1):
        self.input_file1 = input_file1
    
    def write_multiple_cells_and_save(self, cell_values):
        """
        複数のセルに値を書き込み、自動的に保存を行います
        
        Parameters:
        cell_values (dict): セル位置と値の辞書
        """
        try:
            workbook = load_workbook(self.input_file1)
            sheet = workbook.active
            
            for cell_position, value in cell_values.items():
                sheet[cell_position] = value
                print(f"'{value}' をセル {cell_position} に書き込みました")
            
            workbook.save(self.input_file1)
            print("\nすべての内容を保存しました")
            
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")

# 使用例
if __name__ == "__main__":
    input_file1 = r"C:\Users\yukik\Desktop\excel\ex1.xlsx"
    
    # 書き込む内容を辞書形式で定義
    cell_values = {
        'A2': '抹茶大福',
        'A3': 'プリン大福',
        'A4': 'チョコ大福',
        'A5': 'イチゴ大福'
    }
    
    excel_op = ExcelOperator(input_file1)
    excel_op.write_multiple_cells_and_save(cell_values)

#2025/1/26 ex1_A5にいちご大福を書き込みましたOK