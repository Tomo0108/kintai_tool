import os
import sys
import shutil
from pathlib import Path
import configparser

def get_mac_dirs():
    """Mac用の各種ディレクトリパスを取得"""
    app_name = '勤怠表自動作成ツール'
    home = Path.home()
    return {
        'config': home / 'Library' / 'Application Support' / app_name,
        'documents': home / 'Documents' / app_name,
        'templates': home / 'Documents' / app_name / 'templates'
    }

def create_default_config(config_path):
    """デフォルトのconfig.iniを作成"""
    config = configparser.ConfigParser(interpolation=None)
    
    config['DEFAULT'] = {
        'employee_name': '',
        'template_path': str(get_mac_dirs()['templates'] / '勤怠表雛形_2025年版.xlsx')
    }
    
    config['PATHS'] = {
        'input_dir': str(get_mac_dirs()['documents'] / 'input'),
        'output_dir': str(get_mac_dirs()['documents'] / 'output')
    }
    
    config['CSV'] = {
        'encoding': 'utf-8',
        'date_format': '%Y-%m-%d'
    }
    
    with open(config_path, 'w', encoding='utf-8') as f:
        config.write(f)

def get_config_path():
    """プラットフォームに応じたconfig.iniのパスを取得"""
    if sys.platform == 'darwin':
        return get_mac_dirs()['config'] / 'config.ini'
    else:
        if getattr(sys, 'frozen', False):
            base_path = Path(sys.executable).parent
        else:
            base_path = Path(__file__).parent
        return base_path / 'config.ini'

def initialize_mac_environment():
    """Mac用の初期環境をセットアップ"""
    if sys.platform != 'darwin':
        return

    dirs = get_mac_dirs()
    
    # 必要なディレクトリを作成
    for dir_path in dirs.values():
        dir_path.mkdir(parents=True, exist_ok=True)
    
    # input/outputディレクトリを作成
    docs_dir = dirs['documents']
    (docs_dir / 'input').mkdir(exist_ok=True)
    (docs_dir / 'output').mkdir(exist_ok=True)
    
    # テンプレートファイルをコピー
    if getattr(sys, 'frozen', False):
        bundle_templates = Path(sys._MEIPASS) / 'templates'
        if bundle_templates.exists():
            for template in bundle_templates.glob('*.xlsx'):
                shutil.copy2(template, dirs['templates'])
    
    # 初期設定ファイルの作成
    config_path = dirs['config'] / 'config.ini'
    if not config_path.exists():
        create_default_config(config_path)

def get_default_employee_name():
    """デフォルトの従業員名を取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('DEFAULT', 'employee_name', fallback='')

def get_template_path():
    """テンプレートファイルのパスを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('DEFAULT', 'template_path', fallback='')

def get_input_dir():
    """入力ディレクトリのパスを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('PATHS', 'input_dir', fallback='input')

def get_output_dir():
    """出力ディレクトリのパスを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('PATHS', 'output_dir', fallback='output')

def get_csv_encoding():
    """CSVファイルのエンコーディングを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('CSV', 'encoding', fallback='utf-8')

def get_date_format():
    """日付フォーマットを取得"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    return config.get('CSV', 'date_format', fallback='%Y-%m-%d')

def update_config(employee_name, template_path):
    """設定を更新"""
    config = configparser.ConfigParser(interpolation=None)
    config.read(get_config_path(), encoding='utf-8')
    
    if 'DEFAULT' not in config:
        config['DEFAULT'] = {}
    
    config['DEFAULT']['employee_name'] = employee_name
    config['DEFAULT']['template_path'] = template_path
    
    with open(get_config_path(), 'w', encoding='utf-8') as f:
        config.write(f)

def ensure_directories():
    """必要なディレクトリを作成"""
    dirs = ['input', 'output', 'templates']
    for dir_name in dirs:
        os.makedirs(dir_name, exist_ok=True)

# アプリケーション起動時に実行
initialize_mac_environment()

if sys.platform == 'darwin':  # Mac OS の場合
    from tkinter import ttk
    style = ttk.Style()
    style.theme_use('aqua')  # Mac スタイルを適用