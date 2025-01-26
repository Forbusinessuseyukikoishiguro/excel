import win32com.client
import os
from datetime import datetime

class ExcelToPDFConverter:
    def __init__(self, input_file):
        """
        Excel to PDF変換クラスの初期化
        
        Parameters:
        input_file (str): 入力Excelファイルのパス
        """
        self.input_file = input_file
        # 出力PDFのパスを生成（入力ファイルと同じ場所）
        self.output_file = os.path.splitext(input_file)[0] + '.pdf'

    def convert_to_pdf(self):
        """
        ExcelファイルをPDFに変換します
        """
        try:
            # Excel アプリケーションを起動
            excel = win32com.client.Dispatch("Excel.Application")
            # Excelを非表示に設定
            excel.Visible = False
            
            print(f"ファイルを開いています: {self.input_file}")
            # Excelファイルを開く
            wb = excel.Workbooks.Open(os.path.abspath(self.input_file))
            
            print("PDFに変換中...")
            # PDFとして保存
            wb.ExportAsFixedFormat(0, os.path.abspath(self.output_file))
            
            # ファイルを閉じる
            wb.Close()
            excel.Quit()
            
            print(f"PDFファイルを作成しました: {self.output_file}")
            
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")
            # エラーが発生した場合もExcelを確実に終了
            try:
                wb.Close()
                excel.Quit()
            except:
                pass
            
if __name__ == "__main__":
    # 処理対象のExcelファイルのパス
    input_file = r"C:\Users\yukik\Desktop\excel\ex1.xlsx"
    
    # コンバーターのインスタンスを作成して変換を実行
    converter = ExcelToPDFConverter(input_file)
    converter.convert_to_pdf()
#2025/01/26_pdfファイルを作成しました: C:\Users\yukik\Desktop\excel\ex1.pdf_OK
