import os
import sys
import configparser

# `.exe` 実行時のパスを適切に設定
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.ini")

# デフォルト設定
DEFAULT_EMPLOYEE_NAME = "氏名"
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates/勤怠表雛形_2025年版.xlsx")
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
CSV_ENCODING = "utf-8"
DATE_FORMAT = "%Y-%m-%d"

# config.ini を読み込み、存在しない場合は作成
config = configparser.ConfigParser()

if not os.path.exists(CONFIG_FILE):
    config["Settings"] = {
        "DEFAULT_EMPLOYEE_NAME": DEFAULT_EMPLOYEE_NAME,
        "TEMPLATE_PATH": TEMPLATE_PATH,
        "CSV_ENCODING": CSV_ENCODING,
        "DATE_FORMAT": DATE_FORMAT
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

# `config.ini` を読み込む
config.read(CONFIG_FILE, encoding="utf-8")

# `[Settings]` セクションがない場合は追加
if "Settings" not in config:
    config["Settings"] = {
        "DEFAULT_EMPLOYEE_NAME": DEFAULT_EMPLOYEE_NAME,
        "TEMPLATE_PATH": TEMPLATE_PATH,
        "CSV_ENCODING": CSV_ENCODING,
        "DATE_FORMAT": DATE_FORMAT
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

# 設定を変数に格納
DEFAULT_EMPLOYEE_NAME = config["Settings"].get("DEFAULT_EMPLOYEE_NAME", DEFAULT_EMPLOYEE_NAME)
TEMPLATE_PATH = config["Settings"].get("TEMPLATE_PATH", TEMPLATE_PATH)
CSV_ENCODING = config["Settings"].get("CSV_ENCODING", CSV_ENCODING)
DATE_FORMAT = config["Settings"].get("DATE_FORMAT", DATE_FORMAT)

def update_config(name, template):
    """config.ini に氏名とテンプレートファイルのパスを保存"""
    config["Settings"]["DEFAULT_EMPLOYEE_NAME"] = name
    config["Settings"]["TEMPLATE_PATH"] = template
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)
