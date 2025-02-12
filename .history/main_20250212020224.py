import os
import argparse
import pandas as pd
from datetime import datetime
from utils import read_csv, process_data, write_to_excel
from config import DEFAULT_EMPLOYEE_NAME, INPUT_DIR, OUTPUT_DIR

# 必要なフォルダが存在しない場合は作成
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# コマンドライン引数の処理
parser = argparse.ArgumentParser(description="勤怠データをExcelに変換するツール")
parser.add_argument("--name", type=str, default=DEFAULT_EMPLOYEE_NAME, help="従業員名を指定")
parser.add_argument("--file", type=str, required=True, help="処理するCSVファイルのパス")
parser.add_argument("--template", type=str, required=True, help="テンプレートExcelファイルのパス")
args = parser.parse_args()

# CSVデータを読み込み
df = read_csv(args.file)

# CSVから月情報を取得
first_date = df["日付"].min()
year_month = first_date.to_pydatetime().strftime("%Y%m")

# データの整形
df_processed = process_data(df)

# 出力ファイル名を作成
output_filename = f"勤怠表_{year_month}_{args.name}.xlsx"
output_path = os.path.join(OUTPUT_DIR, output_filename)

# Excelに書き込み（不足していた csv_filename を追加）
write_to_excel(args.template, output_path, df_processed, args.file)  

print(f"✅ 勤怠表を作成しました: {output_path}")
