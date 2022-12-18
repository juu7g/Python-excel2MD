"""
エクセルファイルからMarkdown用表テキストを作成
"""

import openpyxl as oxl
import sys, datetime, csv

def Write_csv( file_name:str, headers:list, rows:list):
    """
    CSVファイルの出力   utf_8_sig, crlf

    Args:
        str:        ファイル名
        list:       ヘッダ用リスト
        list:       出力データ用リスト
    """
    try:
        with open(file_name, encoding="utf_8_sig", mode="w", newline="\n") as f:
            _writer = csv.writer(f, delimiter="|")
            for header in headers:
                _writer.writerow(header)
            _writer.writerows(rows)
    except Exception as e:
        print("CSVエラー", e)

# コマンドライン引数からエクセルファイル名を取得
if len(sys.argv) < 2:
    print("error no input file name")
    sys.exit(1)
input_file = sys.argv[1]

print("Convert file:{}".format(input_file))

_wb = oxl.load_workbook(filename=input_file, read_only=True, data_only=True)    # エクセルファイルの読み込み
_sheet = _wb.worksheets[0]            # 先頭のシートの取得
# 各セルの改行文字を置換して戻す
_rows = [[str(x).replace("\n", "<br>") for x in _rows] for _rows in _sheet.values]
kugiri = ["---" for _ in range(_sheet.max_column)]  # markdown表のアラインメント行の作成
dfs = _rows[:1] + [kugiri] +_rows[1:]               # アラインメント行の追加
_wb.close()

output_path = '{}.txt'.format(datetime.datetime.now().strftime("%m%d_%H%M_%S"))
Write_csv(output_path, [], dfs)                     # csvファイルへ出力
print("Convert end")
