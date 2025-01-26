import pandas as pd
from openpyxl import load_workbook, Workbook
import os
from datetime import datetime

class ExcelRowSplitter:
    def __init__(self, input_files, output_dir):
        """
        行分割処理クラスの初期化
        
        Parameters:
        input_files (list): 入力ファイルパスのリスト
        output_dir (str): 出力ディレクトリのパス
        """
        self.input_files = input_files
        self.output_dir = output_dir
        self.all_data = []

    def extract_data_from_excel(self):
        """Excelファイルからデータを抽出"""
        try:
            for file_path in self.input_files:
                wb = load_workbook(file_path)
                sheet = wb.active
                
                for row in sheet.iter_rows():
                    row_data = []
                    for cell in row:
                        row_data.append(cell.value if cell.value is not None else "")
                    if any(row_data):  # 空の行を除外
                        self.all_data.append(row_data)
                            
            print(f"合計 {len(self.all_data)} 行のデータを抽出しました")
            
        except Exception as e:
            print(f"データ抽出中にエラーが発生しました: {str(e)}")

    def split_and_save_data(self, rows_per_file=3):
        """データを指定行数で分割して保存"""
        try:
            # データを指定行数ごとに分割
            for i in range(0, len(self.all_data), rows_per_file):
                # 分割したデータを取得
                chunk = self.all_data[i:i + rows_per_file]
                
                # 新規Excelファイルを作成
                wb = Workbook()
                sheet = wb.active
                
                # データを書き込み
                for row_idx, row_data in enumerate(chunk, 1):
                    for col_idx, value in enumerate(row_data, 1):
                        sheet.cell(row=row_idx, column=col_idx, value=value)
                
                # タイムスタンプを含むファイル名で保存
                file_number = (i // rows_per_file) + 1
                output_file = os.path.join(self.output_dir, f"split_data_{file_number:03d}.xlsx")
                wb.save(output_file)
                print(f"保存完了: {output_file}")
                
        except Exception as e:
            print(f"データ分割・保存中にエラーが発生しました: {str(e)}")

    def process_files(self):
        """一連の処理を実行"""
        print("処理を開始します...")
        # 出力ディレクトリの作成
        os.makedirs(self.output_dir, exist_ok=True)
        
        # データの抽出
        self.extract_data_from_excel()
        
        # データの分割と保存
        self.split_and_save_data()
        
        print("すべての処理が完了しました")

if __name__ == "__main__":
    # 基本パスの設定
    base_path = r"C:\Users\yukik\Desktop\excel"
    
    # 入力ファイルのパス設定
    input_file1 = os.path.join(base_path, "ex1.xlsx")
    input_file2 = os.path.join(base_path, "ex2.xlsx")
    
    # 出力ディレクトリの設定
    output_dir = os.path.join(base_path, "split_files_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    
    # 処理の実行
    splitter = ExcelRowSplitter([input_file1, input_file2], output_dir)
    splitter.process_files()
#2025/01/26 3行ごとに分割して保存処理を開始します...OK