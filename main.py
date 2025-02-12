import os
import argparse
import pandas as pd
from datetime import datetime
from utils import read_csv, process_data, write_to_excel
from config import (get_default_employee_name, get_input_dir, 
                   get_output_dir)

def process_attendance(csv_path, template_path, employee_name):
    """勤怠データの処理を行う関数"""
    # CSVデータを読み込み
    df = read_csv(csv_path)

    # CSVから月情報を取得
    first_date = df["日付"].min()
    year_month = first_date.to_pydatetime().strftime("%Y%m")

    # データの整形
    df_processed = process_data(df)

    # 出力ファイル名を作成
    output_filename = f"勤怠表_{year_month}_{employee_name}.xlsx"
    output_path = os.path.join(get_output_dir(), output_filename)

    # Excelに書き込み
    write_to_excel(template_path, output_path, df_processed, csv_path)

    return output_path

def main():
    """コマンドライン実行用のメイン関数"""
    # 必要なフォルダが存在しない場合は作成
    os.makedirs(get_input_dir(), exist_ok=True)
    os.makedirs(get_output_dir(), exist_ok=True)

    # コマンドライン引数の処理
    parser = argparse.ArgumentParser(description="勤怠データをExcelに変換するツール")
    parser.add_argument("--name", type=str, default=get_default_employee_name(), help="従業員名を指定")
    parser.add_argument("--file", type=str, required=True, help="処理するCSVファイルのパス")
    parser.add_argument("--template", type=str, required=True, help="テンプレートExcelファイルのパス")
    args = parser.parse_args()

    # 処理実行
    output_path = process_attendance(args.file, args.template, args.name)
    print(f"✅ 勤怠表を作成しました: {output_path}")

if __name__ == "__main__":
    main()
