import pandas as pd
from openpyxl import load_workbook
import os

class ExcelSearcher:
    def __init__(self, file_path):
        self.file_path = file_path
    
    def search_keyword(self, keyword):
        """
        指定したキーワードを検索し、出現回数とセル位置を返します
        
        Parameters:
        keyword (str): 検索するキーワード
        """
        try:
            workbook = load_workbook(self.file_path)
            sheet = workbook.active
            count = 0
            positions = []
            
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and keyword in cell.value:
                        count += 1
                        positions.append(f"{cell.coordinate}: {cell.value}")
            
            print(f"\n'{keyword}' の検索結果:")
            print(f"出現回数: {count}回")
            print("\n検出された位置と内容:")
            for position in positions:
                print(position)
            
            return count, positions
            
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")
            return 0, []

if __name__ == "__main__":
    input_file1 = r"C:\Users\yukik\Desktop\excel\ex1.xlsx"
    
    searcher = ExcelSearcher(input_file1)
    searcher.search_keyword("抹茶")
#2025/01/26 ex1の抹茶が含まれるセルの検索を開始しますOK