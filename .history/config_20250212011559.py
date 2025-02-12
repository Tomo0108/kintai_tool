import os
import sys

# .exe として実行された場合、適切なパスを取得
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS  # PyInstaller の一時ディレクトリ
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.ini")

# デフォルトの設定
DEFAULT_EMPLOYEE_NAME = "氏名"
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates/勤怠表雛形_2025年版.xlsx")
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# config.ini がない場合は自動作成
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
