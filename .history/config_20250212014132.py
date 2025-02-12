import os
import sys

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

# **CSVのエンコーディングを追加**
CSV_ENCODING = "utf-8"

# config.ini が存在しない場合は自動作成
if not os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write("[Settings]\n")
        f.write(f'DEFAULT_EMPLOYEE_NAME="{DEFAULT_EMPLOYEE_NAME}"\n')
        f.write(f'TEMPLATE_PATH="{TEMPLATE_PATH}"\n')

def update_config(name, template):
    """config.ini に氏名とテンプレートファイルのパスを保存"""
    config_lines = []
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            if line.startswith("DEFAULT_EMPLOYEE_NAME"):
                config_lines.append(f'DEFAULT_EMPLOYEE_NAME="{name}"\n')
            elif line.startswith("TEMPLATE_PATH"):
                config_lines.append(f'TEMPLATE_PATH="{template}"\n')
            else:
                config_lines.append(line)

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.writelines(config_lines)
