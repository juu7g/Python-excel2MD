"""
エクセルファイルからMarkdown用表テキストを作成
"""

import pandas as pd
import sys, datetime

# コマンドライン引数からエクセルファイル名を取得
if len(sys.argv) < 2:
    print("error no input file name")
    sys.exit(1)
input_file = sys.argv[1]

print("Convert file:{}".format(input_file))

dfs = pd.read_excel(input_file, header=None)    # エクセルファイルの読み込み
dfs.replace("\n", "<br>", True, regex=True)     # 各セルの改行文字を置換
kugiri = {col:"---" for col in dfs.columns}     # markdown表のアラインメント行の作成
dfs = dfs.iloc[:1].append(kugiri, ignore_index=True).append(dfs.iloc[1:])   # アラインメント行の追加

output_path = '{}.txt'.format(datetime.datetime.now().strftime("%m%d_%H%M_%S"))
dfs.to_csv(output_path                          # csvファイルへ出力
                , sep="|"                       # 区切り文字
                , index=False                   # 行名を出力しない
                , header=False                  # 見出しを出力しない
                )
print("Convert end")
