import pandas as pd
from openpyxl import load_workbook, Workbook
import os
from datetime import datetime

class ExcelTextSplitter:
    def __init__(self, input_files, output_dir):
        """
        文字列分割処理クラスの初期化
        
        Parameters:
        input_files (list): 入力ファイルパスのリスト
        output_dir (str): 出力ディレクトリのパス
        """
        self.input_files = input_files
        self.output_dir = output_dir
        self.all_texts = []

    def extract_text_from_excel(self):
        """Excelファイルからテキストを抽出"""
        try:
            for file_path in self.input_files:
                wb = load_workbook(file_path)
                sheet = wb.active
                
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            self.all_texts.append(cell.value)
                            
            print(f"合計 {len(self.all_texts)} 件のテキストを抽出しました")
            
        except Exception as e:
            print(f"テキスト抽出中にエラーが発生しました: {str(e)}")

    def split_and_save_text(self, chunk_size=3):
        """テキストを指定サイズで分割して保存"""
        try:
            for text in self.all_texts:
                # テキストを指定サイズで分割
                chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
                
                # 新規Excelファイルを作成
                wb = Workbook()
                sheet = wb.active
                
                # 分割したテキストを書き込み
                for i, chunk in enumerate(chunks, 1):
                    sheet.cell(row=i, column=1, value=chunk)
                
                # タイムスタンプを含むファイル名で保存
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                output_file = os.path.join(self.output_dir, f"split_text_{timestamp}.xlsx")
                wb.save(output_file)
                print(f"保存完了: {output_file}")
                
        except Exception as e:
            print(f"テキスト分割・保存中にエラーが発生しました: {str(e)}")

    def process_files(self):
        """一連の処理を実行"""
        print("処理を開始します...")
        # 出力ディレクトリの作成
        os.makedirs(self.output_dir, exist_ok=True)
        
        # テキストの抽出
        self.extract_text_from_excel()
        
        # テキストの分割と保存
        self.split_and_save_text()
        
        print("すべての処理が完了しました")

if __name__ == "__main__":
    # 基本パスの設定
    base_path = r"C:\Users\yukik\Desktop\excel"
    
    # 入力ファイルのパス設定
    input_file1 = os.path.join(base_path, "ex1.xlsx")
    input_file2 = os.path.join(base_path, "ex2.xlsx")
    
    # 出力ディレクトリの設定
    output_dir = os.path.join(base_path, "split_texts_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    
    # 処理の実行
    splitter = ExcelTextSplitter([input_file1, input_file2], output_dir)
    splitter.process_files()
#2025/01/23 3文字ずつ分割して保存する処理を開始します...バッチ処理練習