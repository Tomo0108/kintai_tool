import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from config import CSV_ENCODING, DATE_FORMAT

def read_csv(csv_path):
    """
    CSVファイルを読み込み、DataFrameとして返す
    """
    df = pd.read_csv(csv_path, encoding=CSV_ENCODING)
    df["日付"] = pd.to_datetime(df["日付"], format=DATE_FORMAT)
    return df

def process_data(df):
    """
    勤怠データを整形する
    """
    # 必要なカラムの選択
    columns_to_keep = ["日付", "始業時刻", "終業時刻", "総勤務時間", "法定内残業", "時間外労働", "深夜労働", "勤怠種別"]
    df_filtered = df[columns_to_keep].copy()
    
    # 不要データの削除（未入力、休日）
    # 未入力・休日データも残すように変更
    # df_filtered = df_filtered[~df_filtered["勤怠種別"].isin(["未入力", "所定休日", "法定休日"])]
    
    # 時間を小数時間に変換
    def time_to_hours(time_str):
        if isinstance(time_str, str) and ":" in time_str:
            h, m = map(int, time_str.split(":"))
            return h + m / 60  # 分を時間に変換
        return 0

    df_filtered["総勤務時間"] = df_filtered["総勤務時間"].apply(time_to_hours)
    
    return df_filtered

def write_to_excel(template_path, output_path, df):
    """
    ひな型Excelに勤怠データを書き込む
    """
    wb = openpyxl.load_workbook(template_path)
    sheet = wb["勤務表"]
    
    # H5セルの月を取得し、それに基づいてA列の日付を設定
    import re
        month_value = df["日付"].dt.month.iloc[0]  # CSVデータから月を取得
        year_value = df["日付"].dt.year.iloc[0]  # CSVデータから年を取得
        sheet["H5"] = month_value
            heet["F5"] = year_value

    
    for index, row in enumerate(df.itertuples(), start=11):  # C列から開始
        day_value = index - 10  # A11に1日から入力
        sheet[f"A{index}"] = f"=DATE({year_value},{month_value},{day_value})"
        sheet[f"C{index}"] = row.始業時刻
        sheet[f"D{index}"] = row.終業時刻
        sheet[f"E{index}"] = "1:00"  # 休憩時間固定（E列に記述）
        sheet[f"F{index}"] = row.総勤務時間  # 総勤務時間の列を維持
    
    wb.save(output_path)
    print(f"✅ Excelファイルを保存しました: {output_path}")
