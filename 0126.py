import pandas as pd
from openpyxl import load_workbook
import os

class ExcelOperator:
    def __init__(self, input_file1):
        self.input_file1 = input_file1
    
    def write_and_save(self, cell_value, cell_position='A1'):
        """
        指定したセルに値を書き込み、自動的に保存を行います
        
        Parameters:
        cell_value (str): 書き込む値
        cell_position (str): 書き込み先のセル位置（デフォルト：'A1'）
        """
        try:
            workbook = load_workbook(self.input_file1)
            sheet = workbook.active
            sheet[cell_position] = cell_value
            workbook.save(self.input_file1)
            print(f"'{cell_value}' をセル {cell_position} に書き込み、保存しました")
            
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")

# 使用例
if __name__ == "__main__":
    input_file1 = r"C:\Users\yukik\Desktop\excel\ex1.xlsx"
    
    excel_op = ExcelOperator(input_file1)
    excel_op.write_and_save("こんにちは！イチゴ大福さん", "A1")
#2025.01.26保存機能追加A1に書き込み