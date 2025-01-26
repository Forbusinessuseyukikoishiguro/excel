import pandas as pd
from openpyxl import load_workbook
import os

class ExcelOperator:
    def __init__(self, input_file1, input_file2, output_file):
        """
        Excelファイル操作のためのクラスを初期化します
        
        Parameters:
        input_file1 (str): 1つ目の入力Excelファイルのパス
        input_file2 (str): 2つ目の入力Excelファイルのパス
        output_file (str): 出力Excelファイルのパス
        """
        self.input_file1 = input_file1
        self.input_file2 = input_file2
        self.output_file = output_file
        
    def read_excel_files(self):
        """
        Excelファイルを読み込みます
        
        Returns:
        tuple: (df1, df2) - 2つのDataFrame
        """
        try:
            df1 = pd.read_excel(self.input_file1)
            df2 = pd.read_excel(self.input_file2)
            print(f"Successfully read {self.input_file1} and {self.input_file2}")
            return df1, df2
        except Exception as e:
            print(f"Error reading Excel files: {str(e)}")
            return None, None
    
    def write_to_excel(self, dataframes_dict):
        """
        複数のDataFrameを1つのExcelファイルの異なるシートに書き込みます
        
        Parameters:
        dataframes_dict (dict): シート名をキー、DataFrameを値とする辞書
        """
        try:
            with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
                for sheet_name, df in dataframes_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Successfully wrote all data to {self.output_file}")
        except Exception as e:
            print(f"Error writing to Excel: {str(e)}")
    
    def process_excel_files(self):
        """
        Excelファイルの処理を実行します
        """
        df1, df2 = self.read_excel_files()
        if df1 is not None and df2 is not None:
            # 複数のDataFrameを辞書形式でまとめる
            dataframes = {
                'Sheet1': df1,
                'Sheet2': df2
            }
            # まとめて書き込み
            self.write_to_excel(dataframes)

# 使用例
if __name__ == "__main__":
    input_file1 = r"C:\Users\yukik\Desktop\excel\ex1.xlsx"
    input_file2 = r"C:\Users\yukik\Desktop\excel\ex2.xlsx"
    output_file = r"C:\Users\yukik\Desktop\excel\exoutput.xlsx"
    
    excel_op = ExcelOperator(input_file1, input_file2, output_file)
    excel_op.process_excel_files()