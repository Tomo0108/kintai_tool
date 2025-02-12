import pandas as pd
import openpyxl
from datetime import datetime
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
    df_filtered = df_filtered[~df_filtered["勤怠種別"].isin(["未入力", "所定休日", "法定休日"])]
    
    # 時間を小数時間に変換
    def time_to_hours(time_str):
        if isinstance(time_str, str) and ":" in time_str:
            h, m = map(int, time_str.split(":"))
            return h + m / 60  # 分を時間に変換
        return 0

    df_filtered["総勤務時間"] = df_filtered["総勤務時間"].apply(time_to_hours)
    
    # 日付をExcelのシリアル値に変換
    df_filtered["日付"] = df_filtered["日付"].apply(lambda x: x.toordinal() + 2)
    
    return df_filtered

def write_to_excel(template_path, output_path, df):
    """
    ひな型Excelに勤怠データを書き込む
    """
    wb = openpyxl.load_workbook(template_path)
    sheet = wb["勤務表"]
    
    # データを書き込む開始位置（例: A9 以降）
    start_row = 9
    for index, row in df.iterrows():
        sheet[f"A{start_row + index}"] = row["日付"]
        sheet[f"B{start_row + index}"] = row["始業時刻"]
        sheet[f"C{start_row + index}"] = row["終業時刻"]
        sheet[f"D{start_row + index}"] = row["総勤務時間"]
    
    wb.save(output_path)
    print(f"✅ Excelファイルを保存しました: {output_path}")
