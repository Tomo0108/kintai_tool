import os
import configparser
from pathlib import Path

def get_config_path():
    """実行ファイルと同じディレクトリのconfig.iniのパスを返す"""
    if getattr(sys, 'frozen', False):
        # exe実行時
        base_path = Path(sys.executable).parent
    else:
        # 通常実行時
        base_path = Path(__file__).parent
    return base_path / 'config.ini'

def load_config():
    """設定を読み込む"""
    config = configparser.ConfigParser()
    config.read(get_config_path(), encoding='utf-8')
    return config

def save_config(config):
    """設定を保存する"""
    with open(get_config_path(), 'w', encoding='utf-8') as f:
        config.write(f)

def get_default_employee_name():
    return load_config()['DEFAULT']['employee_name']

def get_template_path():
    return load_config()['DEFAULT']['template_path']

def get_input_dir():
    return load_config()['PATHS']['input_dir']

def get_output_dir():
    return load_config()['PATHS']['output_dir']

def get_csv_encoding():
    return load_config()['CSV']['encoding']

def get_date_format():
    return load_config()['CSV']['date_format']

def update_config(name, template):
    """config.iniに氏名とテンプレートファイルのパスを保存"""
    config = load_config()
    config['DEFAULT']['employee_name'] = name
    config['DEFAULT']['template_path'] = template
    save_config(config)