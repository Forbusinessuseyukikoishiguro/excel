import pandas as pd
from openpyxl import load_workbook
import os

class ExcelProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
    
    def add_sale_message(self, keyword, message="お買い得！"):
        """
        指定したキーワードを含むセルの隣のセルにメッセージを追加します
        
        Parameters:
        keyword (str): 検索するキーワード
        message (str): 追加するメッセージ（デフォルト: お買い得！）
        """
        try:
            workbook = load_workbook(self.file_path)
            sheet = workbook.active
            update_count = 0
            
            # キーワードを含むセルを検索し、隣のセルにメッセージを追加
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and keyword in cell.value:
                        # 現在のセルの列番号を取得し、隣のセルを特定
                        next_cell = sheet.cell(row=cell.row, column=cell.column + 1)
                        next_cell.value = message
                        update_count += 1
                        print(f"{cell.coordinate}の隣のセル{next_cell.coordinate}に'{message}'を追加しました")
            
            # 変更を保存
            workbook.save(self.file_path)
            print(f"\n合計{update_count}箇所を更新し、保存しました")
            
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")

if __name__ == "__main__":
    input_file1 = r"C:\Users\yukik\Desktop\excel\ex1.xlsx"
    
    processor = ExcelProcessor(input_file1)
    processor.add_sale_message("抹茶")
#2025/1/26 抹茶の隣のセルB2に'お買い得！'を追加しました