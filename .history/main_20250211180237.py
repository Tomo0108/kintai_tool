import os
import argparse
import pandas as pd
from datetime import datetime
from utils import read_csv, process_data, write_to_excel
from config import DEFAULT_EMPLOYEE_NAME, TEMPLATE_PATH

# 入力・出力フォルダ
input_dir = "input"
output_dir = "output"
template_path = TEMPLATE_PATH

# 必要なフォルダを作成
os.makedirs(input_dir, exist_ok=True)
os.makedirs(output_dir, exist_ok=True)

# 最新のCSVファイルを取得
csv_files = [f for f in os.listdir(input_dir) if f.endswith(".csv")]
if not csv_files:
    raise FileNotFoundError("inputフォルダにCSVファイルが見つかりません。")

latest_csv = sorted(csv_files)[-1]  # 一番新しいファイルを選択
csv_path = os.path.join(input_dir, latest_csv)

# CSVデータを読み込み
df = read_csv(csv_path)

# CSVから月情報を取得
first_date = df["日付"].min()
year_month = datetime.fromordinal(int(first_date) - 2).strftime("%Y%m")

# データの整形
df_processed = process_data(df)

# 出力ファイル名を作成
output_filename = f"勤怠表_{year_month}_{DEFAULT_EMPLOYEE_NAME}.xlsx"
output_path = os.path.join(output_dir, output_filename)

# Excelに書き込み
write_to_excel(template_path, output_path, df_processed)

print(f"✅ 勤怠表を作成しました: {output_path}")
