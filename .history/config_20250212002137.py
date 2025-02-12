import os
import sys

CONFIG_FILE = __file__ if not getattr(sys, 'frozen', False) else "config.py"

# デフォルトの従業員名
DEFAULT_EMPLOYEE_NAME = "氏名"

# データフォルダ（入力・出力）
INPUT_DIR = "input"
OUTPUT_DIR = "output"

# テンプレートExcelのパス
TEMPLATE_PATH = "templates/勤怠表雛形_2025年版.xlsx"

def update_config(name, template):
    """config.py に氏名とテンプレートファイルのパスを保存"""
    global DEFAULT_EMPLOYEE_NAME, TEMPLATE_PATH
    DEFAULT_EMPLOYEE_NAME = name
    TEMPLATE_PATH = template

    # .exe 実行時は書き換えをスキップ
    if getattr(sys, 'frozen', False):
        return

    # config.py を直接編集
    config_lines = []
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            if line.startswith("DEFAULT_EMPLOYEE_NAME"):
                config_lines.append(f'DEFAULT_EMPLOYEE_NAME = "{name}"\n')
            elif line.startswith("TEMPLATE_PATH"):
                config_lines.append(f'TEMPLATE_PATH = "{template}"\n')
            else:
                config_lines.append(line)

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.writelines(config_lines)
