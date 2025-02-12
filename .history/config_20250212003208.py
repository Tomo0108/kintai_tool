import os
import sys
import configparser

CONFIG_FILE = "config.ini"  # `.exe` でも書き換え可能な `config.ini` を使用

# 設定を読み込む
config = configparser.ConfigParser()
if not os.path.exists(CONFIG_FILE):
    config["Settings"] = {
        "DEFAULT_EMPLOYEE_NAME": "氏名",
        "TEMPLATE_PATH": "templates/勤怠表雛形_2025年版.xlsx"
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

config.read(CONFIG_FILE, encoding="utf-8")

# 設定の取得
DEFAULT_EMPLOYEE_NAME = config["Settings"].get("DEFAULT_EMPLOYEE_NAME", "氏名")
TEMPLATE_PATH = config["Settings"].get("TEMPLATE_PATH", "templates/勤怠表雛形_2025年版.xlsx")

# 設定の更新
def update_config(name, template):
    """`config.ini` に設定を保存"""
    config["Settings"]["DEFAULT_EMPLOYEE_NAME"] = name
    config["Settings"]["TEMPLATE_PATH"] = template
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)
