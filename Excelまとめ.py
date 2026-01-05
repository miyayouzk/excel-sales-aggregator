from pathlib import Path
import pandas as pd

#フォルダのパスを作成
base_dir = Path(__file__).parent #現在のフォルダをベースに指定
input_dir = base_dir / "input" #inputフォルダの位置を指定
output_dir = base_dir / "output" #outputフォルダの位置を指定
output_dir.mkdir(exist_ok=True) #outputフォルダを作成、あるならスルー

#作成したパスからexcelファイルを取得
files = [
    f for f in input_dir.glob("*.xlsx") #excelファイルを取得、fに格納
    if not f.name.startswith("~$") #"~$"から始まるexcelファイルを除外してfilesに格納
]

#取得したexcelファイルを読み込み
dfs = [] #excelファイルをここに入れる。
for f in files: #ファイルを一つずつ処理
    df = pd.read_excel(f) #excelファイルを読み込み
    dfs.append(df) #excelファイルをdfsリスト内に格納

#excelファイルを縦に結合
all_df = pd.concat(dfs, ignore_index=True) #番号を振りなおして縦方向にexcelファイルを結合

#「¥や,」などの文字列を数値に変換する処理
all_df["合計金額"] = ( #¥や,を消して数字として扱えるように
    all_df["合計金額"] #DFから合計金額列だけを取り出す
    .replace("[¥,]", "", regex=True) #[]内の¥,どちらかにマッチしたら""空文字に置き換える
    .astype(int) #数値に変換
)

#担当者別に売り上げを集計・並び替え
staff_sales = (
    all_df
    .groupby("担当者")["合計金額"] #各担当者の合計金額をグループ化
    .sum() #合計金額を合算
    .reset_index() #表に戻す
    .sort_values("合計金額", ascending=False) #合計金額を基準に大きい順に並び替え
)

#商品別に売り上げを集計・並び替え
product_sales = (
    all_df
    .groupby("商品名")["合計金額"]
    .sum()
    .reset_index()
    .sort_values("合計金額", ascending=False)
)

#日付ごとに売り上げを集計・並び替え
daily_sales = (
    all_df
    .groupby("日付")["合計金額"]
    .sum()
    .reset_index()
    .sort_values("日付")
)

#ピポッドテーブルを作成
pivot_df = pd.pivot_table(
    all_df,
    index = "日付", #行
    columns = "担当者", #列
    values = "合計金額", #集計対象
    aggfunc = "sum", #合計
    fill_value = 0 #データなしは0
)

pivot_df = pivot_df.sort_index()

#合計金額を通貨表示に変換する処理
staff_sales["合計金額"] = staff_sales["合計金額"].map("¥{:,.0f}".format)
product_sales["合計金額"] = product_sales["合計金額"].map("¥{:,.0f}".format)
daily_sales["合計金額"] = daily_sales["合計金額"].map("¥{:,.0f}".format)
pivot_df = pivot_df.applymap(lambda x: f"¥{x:,.0f}")

#excelファイルを出力
output_path = output_dir / "売上集計.xlsx"

#一つのexcelファイルにまとめるためにExcelWriterを使って書き出し
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    all_df.to_excel(writer, sheet_name="売上明細", index=False)
    staff_sales.to_excel(writer, sheet_name="担当者別売上", index=False)
    product_sales.to_excel(writer, sheet_name="商品別売上", index=False)
    daily_sales.to_excel(writer, sheet_name="日別売上", index=False)
    pivot_df.to_excel(writer, sheet_name="担当者×日付")

#個別に書き出す場合は以下のような書き方をする
# staff_sales.to_excel(output_dir / "担当者別売上.xlsx", index=False)
# product_sales.to_excel(output_dir / "商品別売上.xlsx", index=False)

print("売上集計 完了！")