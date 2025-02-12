import os
import argparse
import pandas as pd
from datetime import datetime
from utils import read_csv, process_data, write_to_excel
from config import DEFAULT_EMPLOYEE_NAME, TEMPLATE_PATH, INPUT_DIR, OUTPUT_DIR

# 必要なフォルダが存在しない場合は作成
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# コマンドライン引数の処理
parser = argparse.ArgumentParser(description="勤怠データをExcelに変換するツール")
parser.add_argument("--name", type=str, default=DEFAULT_EMPLOYEE_NAME, help="従業員名を指定")
args = parser.parse_args()

# 最新のCSVファイルを取得
csv_files = [f for f in os.listdir(INPUT_DIR) if f.endswith(".csv")]
if not csv_files:
    raise FileNotFoundError("inputフォルダにCSVファイルが見つかりません。")

latest_csv = sorted(csv_files)[-1]  # 一番新しいファイルを選択
csv_path = os.path.join(INPUT_DIR, latest_csv)

# CSVデータを読み込み
df = read_csv(csv_path)

# CSVから月情報を取得
first_date = df["日付"].min()
year_month = first_date.to_pydatetime().strftime("%Y%m")

# データの整形
df_processed = process_data(df)

# 出力ファイル名を作成（「氏名」に統一）
output_filename = f"勤怠表_{year_month}_氏名.xlsx"
output_path = os.path.join(OUTPUT_DIR, output_filename)

# Excelに書き込み
write_to_excel(TEMPLATE_PATH, output_path, df_processed)

print(f"✅ 勤怠表を作成しました: {output_path}")
