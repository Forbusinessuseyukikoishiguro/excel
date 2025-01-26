import pandas as pd
from openpyxl import load_workbook
import os
from datetime import datetime

class ExcelOperator:
    def __init__(self):
        """
        ExcelOperatorクラスの初期化
        ファイルパスは各メソッド内で指定できるように変更
        """
        pass

    def write_multiple_cells_and_save(self, file_path, cell_values):
        """
        複数のセルに値を書き込み、自動的に保存を行います
        
        Parameters:
        file_path (str): 対象のExcelファイルパス
        cell_values (dict): セル位置と値の辞書
        """
        try:
            workbook = load_workbook(file_path)
            sheet = workbook.active
            
            print(f"\n{os.path.basename(file_path)}への書き込みを開始します：")
            for cell_position, value in cell_values.items():
                sheet[cell_position] = value
                print(f"'{value}' をセル {cell_position} に書き込みました")
            
            workbook.save(file_path)
            print(f"{os.path.basename(file_path)}にすべての内容を保存しました\n")
            
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")

def main():
    # ファイルパスの設定
    base_path = r"C:\Users\yukik\Desktop\excel"
    input_file1 = os.path.join(base_path, "ex1.xlsx")
    input_file2 = os.path.join(base_path, "ex2.xlsx")
    
    # ExcelOperatorのインスタンス作成
    excel_op = ExcelOperator()
    
    # ex1.xlsxの書き込み内容
    values_ex1 = {
        'A2': '抹茶チョコ',
        'A3': 'プリンチョコ',
        'A4': 'チョコ大福'
    }
    
    # ex2.xlsxの書き込み内容
    values_ex2 = {
        'A1': 'イチゴチョコ',
        'A2': 'ブドウチョコ',
        'A3': '抹茶あいす',
        'A4': '最中抹茶金時'
    }
    
    # 処理の実行
    print(f"処理開始: {datetime.now().strftime('%Y/%m/%d %H:%M:%S')}")
    
    # ex1.xlsxの処理
    excel_op.write_multiple_cells_and_save(input_file1, values_ex1)
    
    # ex2.xlsxの処理
    excel_op.write_multiple_cells_and_save(input_file2, values_ex2)
    
    print(f"処理完了: {datetime.now().strftime('%Y/%m/%d %H:%M:%S')}")

if __name__ == "__main__":
    main()
    
#2025/01/26 ex1とex2のA1~A4の入力処理＿まとめるテスト_インプット自動保存を統合して実行して入力（まとめてExcel２ファイル入力処理OK！）