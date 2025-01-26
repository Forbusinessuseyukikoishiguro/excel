import pandas as pd
from openpyxl import load_workbook, Workbook
import os

class ExcelExtractor:
    def __init__(self, input_files, output_file):
        """
        ExcelExtractorクラスの初期化
        
        Parameters:
        input_files (list): 入力ファイルパスのリスト
        output_file (str): 出力ファイルパス
        """
        self.input_files = input_files
        self.output_file = output_file
    
    def extract_and_save_matcha_items(self):
        """
        入力ファイルから抹茶商品を抽出し、新しいファイルに保存します
        """
        try:
            # 出力用の新しいワークブックを作成
            output_wb = Workbook()
            output_sheet = output_wb.active
            output_sheet.title = "抹茶商品リスト"
            
            # ヘッダーを設定
            output_sheet['A1'] = "商品名"
            output_sheet['B1'] = "ファイル元"
            current_row = 2
            
            # 各入力ファイルを処理
            for input_file in self.input_files:
                input_wb = load_workbook(input_file)
                input_sheet = input_wb.active
                file_name = os.path.basename(input_file)
                
                # 各セルをチェック
                for row in input_sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and "抹茶" in cell.value:
                            # 抹茶商品を出力ファイルに追加
                            output_sheet[f'A{current_row}'] = cell.value
                            output_sheet[f'B{current_row}'] = file_name
                            print(f"抽出: {cell.value} (from {file_name})")
                            current_row += 1
            
            # 結果を保存
            output_wb.save(self.output_file)
            print(f"\n合計 {current_row - 2} 件の抹茶商品を抽出し、{self.output_file} に保存しました。")
            
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")

if __name__ == "__main__":
    # ファイルパスの設定
    base_path = r"C:\Users\yukik\Desktop\excel"
    input_file1 = os.path.join(base_path, "ex1.xlsx")
    input_file2 = os.path.join(base_path, "ex2.xlsx")
    output_file = os.path.join(base_path, "exoutput.xlsx")
    
    # ExcelExtractorのインスタンスを作成し、処理を実行
    extractor = ExcelExtractor([input_file1, input_file2], output_file)
    extractor.extract_and_save_matcha_items()
    
#2025/1/26 ex1.xlsxとex2.xlsxから抹茶商品を抽出し、exoutput.xlsxに保存しました。OK
