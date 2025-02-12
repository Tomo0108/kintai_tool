import os
import argparse
import pandas as pd
from datetime import datetime
from utils import read_csv, process_data, write_to_excel
from config import DEFAULT_EMPLOYEE_NAME

# データフォルダとテンプレートのパス
data_dir = "data"
template_path = "templates/勤怠表雛形.xlsx"

# コマンドライン引数の処理
parser = argparse.ArgumentParser(description="勤怠データをExcelに変換するツール")
parser.add_argument("--name", type=str, default=DEFAULT_EMPLOYEE_NAME, help="従業員名を指定")
args = parser.parse_args()

# 最新のCSVファイルを取得（dataディレクトリ内の最新ファイル）
csv_files = [f for f in os.listdir(data_dir) if f.endswith(".csv")]
if not csv_files:
    raise FileNotFoundError("データフォルダにCSVファイルが見つかりません。")
latest_csv = sorted(csv_files)[-1]  # 一番新しいファイルを選択
csv_path = os.path.join(data_dir, latest_csv)

# CSVデータを読み込み
df = read_csv(csv_path)

# CSVから月情報を取得
first_date = df["日付"].min()
year_month = datetime.fromordinal(int(first_date) - 2).strftime("%Y%m")

# データの整形
df_processed = process_data(df)

# 出力ファイル名を作成
output_filename = f"勤怠表_{year_month}_{args.name}.xlsx"
output_path = os.path.join(data_dir, output_filename)

# Excelに書き込み
write_to_excel(template_path, output_path, df_processed)

print(f"✅ 勤怠表を作成しました: {output_path}")
