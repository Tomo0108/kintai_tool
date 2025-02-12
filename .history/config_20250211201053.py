# config.py

import os

CONFIG_FILE = "config.ini"

# デフォルトの従業員名（config.ini から読み込み）
DEFAULT_EMPLOYEE_NAME = "氏名"  # 初期値

if os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            if "DEFAULT_EMPLOYEE_NAME" in line:
                DEFAULT_EMPLOYEE_NAME = line.split("=")[1].strip().strip('"')
                break

# データフォルダ（入力・出力）
INPUT_DIR = "input"
OUTPUT_DIR = "output"

# テンプレートExcelのパス
TEMPLATE_PATH = "templates/勤怠表雛形_2025年版.xlsx"

# CSVファイルのエンコーディング
CSV_ENCODING = "utf-8"

# 日付フォーマット
DATE_FORMAT = "%Y-%m-%d"

def update_config(name):
    """config.ini に氏名を保存"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(f'DEFAULT_EMPLOYEE_NAME="{name}"\n')
