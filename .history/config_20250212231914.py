import os
import sys
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

def create_default_config():
    """デフォルトの設定ファイルを作成する"""
    config = configparser.ConfigParser()
    
    # デフォルト設定
    config['DEFAULT'] = {
        'employee_name': '小島知将',
        'template_path': 'templates/勤怠表雛形_2025年版.xlsx'
    }
    
    config['PATHS'] = {
        'input_dir': 'input',
        'output_dir': 'output'
    }
    
    config['CSV'] = {
        'encoding': 'utf-8',
        'date_format': '%Y-%m-%d'
    }
    
    # 設定ファイルを保存
    with open(get_config_path(), 'w', encoding='utf-8') as f:
        config.write(f)

def load_config():
    """設定を読み込む。ない場合は作成する"""
    config = configparser.ConfigParser()
    config_path = get_config_path()
    
    # config.iniが存在しない場合、デフォルト設定で作成
    if not config_path.exists():
        create_default_config()
    
    config.read(config_path, encoding='utf-8')
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